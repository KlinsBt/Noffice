"""Compare all independent Inter cases, real DOCX exports and downloaded PDF glyphs."""
from pathlib import Path
import argparse, hashlib, json, sys
base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
from word_pdf_glyphs import glyphs

parser = argparse.ArgumentParser()
parser.add_argument('--negative-control', choices=['pixel', 'glyph', 'page', 'text'])
selection = parser.add_mutually_exclusive_group()
selection.add_argument('--long', action='store_true')
selection.add_argument('--boundary', action='store_true')
args = parser.parse_args()
matrix = 'word-inter-long-metrics' if args.long else 'word-leading-boundary' if args.boundary else 'word-inter-metrics'
root = base / '.local' / matrix
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
oracle = read(base / 'tests/fixtures' / ('native-' + matrix + '.json'))
receipt = read(root / 'native-report.json')
assert receipt['captured']
assert receipt['scriptHash'] == sha(base / 'scripts/verify-word-inter-metrics.ps1')
assert receipt['oracleHash'] == sha(base / 'tests/fixtures' / ('native-' + matrix + '.json'))
assert receipt['fontHash'] == oracle['fontHash']
assert receipt['executableHash'] == oracle['executableHash'] == next(
    a['sha256'] for a in read(base / 'docs/parity/baseline.json')['applications']
    if a['executable'] == 'WINWORD.EXE')
assert [c['name'] for c in receipt['cases']] == [c['name'] for c in oracle['cases']]
rows = []
for case, reference in zip(receipt['cases'], oracle['cases']):
    folder = root / case['name']
    browser = read(folder / 'browser-report.json')
    assert browser['passed'] and sha(folder / 'browser-report.json') == case['browserHash']
    expected, actual = case['states']['expected'], case['states']['actual']
    assert expected['sourceHash'] == reference['sourceHash'] == sha(
        base / 'tests/fixtures' / matrix / (case['name'] + '.docx'))
    assert actual['sourceHash'] == browser['outputs']['docx'] == sha(folder / 'restored.docx')
    assert sha(folder / 'restored.pdf') == browser['outputs']['pdf']
    assert sha(folder / 'expected.pdf') == expected['pdfHash']
    assert sha(folder / 'actual.pdf') == actual['pdfHash']
    semantic = expected['paragraphs'] == actual['paragraphs'] and expected['pages'] == actual['pages']
    editable = browser['linePages'].copy()
    if args.negative_control == 'page' and not rows:
        editable[0] += 1
    page_assignments = editable == [c['page'] for c in reference['starts']]
    pages = []
    with pdfium.PdfDocument(folder / 'expected.pdf') as a, pdfium.PdfDocument(folder / 'actual.pdf') as b, pdfium.PdfDocument(folder / 'restored.pdf') as c:
        counts = [len(a), len(b), len(c)]
        for i in range(min(counts)):
            p, q, r = a[i], b[i], c[i]
            left, right = p.render(scale=2), q.render(scale=2)
            try:
                pixels = bytearray(right.buffer)
                if args.negative_control == 'pixel' and not rows and i == 0:
                    pixels[0] ^= 1
                exact = p.get_size() == q.get_size() and bytes(left.buffer) == bytes(pixels)
                # These independently authored samples contain no spaces.
                # Ignore only PDFium-generated word gaps; real PDF spaces
                # and all source/export text identities remain checked.
                no_spaces = all(' ' not in paragraph['text'] for paragraph in expected['paragraphs'])
                first, last = glyphs(p, no_spaces), glyphs(r, no_spaces)
                if args.negative_control == 'text' and not rows and i == 0:
                    last[0]['text'] = ' '
                text = [g['text'] for g in first] == [g['text'] for g in last]
                pairs = [(g, h) for g, h in zip(first, last) if not (g.get('generated') or h.get('generated'))]
                if args.negative_control == 'glyph' and not rows and i == 0:
                    pairs[0][1]['y'] += 1
                differences = {axis: max((abs(g[axis]-h[axis]) for g, h in pairs), default=0) for axis in ['x', 'y', 'size']}
                colors = all(g['color'] == h['color'] for g, h in pairs)
                paper = all(abs(u-v) <= .15 for u, v in zip(p.get_size(), r.get_size()))
                pages.append({'page': i+1, 'exactNativePixels': exact, 'text': text, 'colors': colors,
                              'glyphs': len(pairs), 'differences': differences, 'paperPassed': paper,
                              'passed': exact and text and colors and paper and all(v <= .15 for v in differences.values())})
            finally:
                left.close(); right.close(); p.close(); q.close(); r.close()
        rows.append({'name': case['name'], 'nativeSnapshot': semantic, 'editablePages': page_assignments,
                     'pageCounts': counts, 'pages': pages,
                     'passed': semantic and page_assignments and len(set(counts)) == 1 and all(p['passed'] for p in pages)})
report = {'passed': all(r['passed'] for r in rows), 'nativeReceiptHash': sha(root / 'native-report.json'),
          'comparerHash': sha(Path(__file__)), 'rows': rows}
name = f'negative-{args.negative_control}.json' if args.negative_control else 'comparison.json'
(root / name).write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps(report, indent=2))
sys.exit(0 if report['passed'] else 1)
