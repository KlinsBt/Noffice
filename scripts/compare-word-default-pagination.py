"""Gate actual default-font DOCX/PDF outputs against independently authored Word."""
from pathlib import Path
import argparse, hashlib, json, sys

base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
from word_pdf_glyphs import glyphs

parser = argparse.ArgumentParser()
parser.add_argument('--inherited', action='store_true')
parser.add_argument('--negative-control', choices=['pixel', 'glyph', 'page', 'text'])
args = parser.parse_args()
root = base / '.local' / ('word-default-inheritance' if args.inherited else 'word-default-pagination')
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
receipt = read(root / 'native-report.json')
browser = read(root / 'browser-report.json')
assert receipt['captured'] and browser['passed']
assert receipt['browserHash'] == sha(root / 'browser-report.json')
assert receipt['referenceHash'] == sha(base / '.local/word-default-fonts/variable-embedded-native.docx')
assert receipt['scriptHash'] == sha(base / 'scripts/verify-word-default-pagination.ps1')
assert receipt['executableHash'] == next(a['sha256'] for a in read(base / 'docs/parity/baseline.json')['applications'] if a['executable'] == 'WINWORD.EXE')
expected_stages = ['source', 'edited', 'restored', 'reimported'] if args.inherited else ['empty', 'source', 'edited', 'restored', 'reimported']
assert [row['name'] for row in receipt['cases']] == expected_stages
rows = []
controlled = False
for row in receipt['cases']:
    name = row['name']
    first, last = row['states']['expected'], row['states']['actual']
    assert first['inputHash'] == sha(root / f'native-{name}.docx')
    assert last['inputHash'] == browser['outputs'][f'{name}.docx'] == sha(root / f'{name}.docx')
    assert browser['outputs'][f'{name}.pdf'] == sha(root / f'{name}.pdf')
    assert first['pdfHash'] == sha(root / f'{name}-expected.pdf')
    assert last['pdfHash'] == sha(root / f'{name}-actual.pdf')
    native = first['snapshot'] == last['snapshot']
    expected_positions = [[{'text': c['text'].replace('\v', '\n'), 'page': c['page']} for c in p['characters']] for p in first['snapshot']['paragraphs']]
    editable = json.loads(json.dumps(browser[name])) if name != 'empty' else [[]]
    control_case = name != 'empty' and not controlled
    if control_case and args.negative_control == 'page':
        editable[0][0]['page'] += 1
    page_assignments = editable == expected_positions
    no_spaces = all(' ' not in p['text'] for p in first['snapshot']['paragraphs'])
    pages = []
    with pdfium.PdfDocument(root / f'{name}-expected.pdf') as a, pdfium.PdfDocument(root / f'{name}-actual.pdf') as b, pdfium.PdfDocument(root / f'{name}.pdf') as c:
        counts = [len(a), len(b), len(c)]
        for i in range(min(counts)):
            p, q, r = a[i], b[i], c[i]
            left, right = p.render(scale=2), q.render(scale=2)
            try:
                pixels = bytearray(right.buffer)
                if control_case and i == 0 and args.negative_control == 'pixel':
                    pixels[0] ^= 1
                exact = p.get_size() == q.get_size() and bytes(left.buffer) == bytes(pixels)
                expected, actual = glyphs(p, no_spaces), glyphs(r, no_spaces)
                if control_case and i == 0 and args.negative_control == 'text':
                    actual[0]['text'] = '!'
                text = [g['text'] for g in expected] == [g['text'] for g in actual]
                pairs = [(g, h) for g, h in zip(expected, actual) if not (g.get('generated') or h.get('generated'))]
                if control_case and i == 0 and args.negative_control == 'glyph':
                    pairs[0][1]['y'] += 1
                differences = {axis: max((abs(g[axis]-h[axis]) for g, h in pairs), default=0) for axis in ['x', 'y', 'size']}
                color = all(g['color'] == h['color'] for g, h in pairs)
                paper = all(abs(u-v) <= .15 for u, v in zip(p.get_size(), r.get_size()))
                pages.append({'page': i+1, 'exactNativePixels': exact, 'text': text, 'color': color, 'paper': paper, 'glyphs': len(pairs), 'differences': differences,
                              'passed': exact and text and color and paper and all(v <= .15 for v in differences.values())})
            finally:
                left.close(); right.close(); p.close(); q.close(); r.close()
    rows.append({'name': name, 'nativeSnapshot': native, 'editablePages': page_assignments, 'pageCounts': counts, 'pages': pages,
                 'passed': native and page_assignments and len(set(counts)) == 1 and all(p['passed'] for p in pages)})
    if control_case:
        controlled = True
report = {'passed': all(r['passed'] for r in rows), 'nativeReceiptHash': sha(root / 'native-report.json'), 'comparerHash': sha(Path(__file__)), 'rows': rows}
output = f'negative-{args.negative_control}.json' if args.negative_control else 'comparison.json'
(root / output).write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps(report, indent=2))
sys.exit(0 if report['passed'] else 1)
