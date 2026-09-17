"""Compare independent authoring, actual DOCX downloads, editable pages and PDF glyphs."""
from pathlib import Path
import argparse, hashlib, json, sys
base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
from word_pdf_glyphs import glyphs

root = base / '.local/word-authored-pagination'
parser = argparse.ArgumentParser()
parser.add_argument('--negative-control', choices=['pixel', 'glyph', 'page', 'paragraph'])
parser.add_argument('--pdf-kind', choices=['print', 'download'], default='print')
args = parser.parse_args()
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
receipt = read(root / 'native-report.json')
browser = read(root / 'browser-report.json')
assert receipt['captured'] and browser['passed']
assert receipt['scriptHash'] == sha(base / 'scripts/verify-word-authored-pagination.ps1')
assert receipt['executableHash'] == next(a['sha256'] for a in read(base / 'docs/parity/baseline.json')['applications'] if a['executable'] == 'WINWORD.EXE')
for artifact in receipt['hashes']:
    assert sha(root / artifact['path']) == artifact['sha256'], 'Stale ' + artifact['path']
rows = []
for stage, key in [('source', 'initial'), ('edited', 'edited'), ('restored', 'restored'), ('source', 'reimported'), ('automatic', 'automatic'), ('minimum', 'minimum')]:
    expected, actual = receipt['stages'][stage]['expected'], receipt['stages'][stage]['actual']
    if args.negative_control == 'paragraph' and key == 'initial':
        actual['paragraphs'][0]['after'] += 1
    fields = ['lineRule', 'line', 'before', 'after', 'markFont', 'markSize']
    properties = len(expected['paragraphs']) == len(actual['paragraphs']) and all(
        all(p[field] == q[field] for field in fields)
        for p, q in zip(expected['paragraphs'], actual['paragraphs']))
    wanted = [c for p in expected['paragraphs'] for c in p['characters']]
    exported = [c for p in actual['paragraphs'] for c in p['characters']]
    editable = [c for p in browser[key] for c in p]
    if args.negative_control == 'page' and key == 'initial':
        editable[0]['page'] += 1
    semantics = len(expected['paragraphs']) == len(actual['paragraphs']) and [c['text'] for c in wanted] == [c['text'] for c in exported]
    native = semantics and properties and wanted == exported and expected['pages'] == actual['pages']
    dom = len(editable) == len(wanted) and all(a['text'] == b['text'].replace('\v', '\n') and a['page'] == b['page'] for a, b in zip(editable, wanted))
    printed = 'reimported' if key == 'reimported' else stage
    pdf_name = f'{"download" if args.pdf_kind == "download" else "browser"}-{printed}.pdf'
    if args.pdf_kind == 'download':
        assert sha(root / pdf_name) == browser['pdfExports'][printed], 'Stale PDF download'
    with pdfium.PdfDocument(root / f'expected-{stage}.pdf') as a, pdfium.PdfDocument(root / f'native-{stage}.pdf') as b, pdfium.PdfDocument(root / pdf_name) as c:
        pages = []
        for index in range(min(len(a), len(b), len(c))):
            p, q, r = a[index], b[index], c[index]
            x, y = p.render(scale=2), q.render(scale=2)
            try:
                output = bytearray(y.buffer)
                if args.negative_control == 'pixel' and key == 'initial' and index == 0:
                    output[0] ^= 1
                pixels = p.get_size() == q.get_size() and bytes(x.buffer) == bytes(output)
                first, last = glyphs(p, False), glyphs(r, False)
                text = [g['text'] for g in first] == [g['text'] for g in last]
                pairs = [(g, h) for g, h in zip(first, last) if not (g.get('generated') or h.get('generated'))]
                if args.negative_control == 'glyph' and key == 'initial' and index == 0:
                    pairs[0][1]['y'] += 1
                differences = {axis: max((abs(g[axis]-h[axis]) for g,h in pairs), default=0) for axis in ['x','y','size']}
                paper = all(abs(u-v) <= .15 for u,v in zip(p.get_size(), r.get_size()))
                painted = text and all(g['color']==h['color'] for g,h in pairs) and all(v <= .15 for v in differences.values())
                pages.append({'page': index+1, 'nativePixels': pixels, 'text': text, 'glyphs': len(pairs), 'differences': differences,
                              'nativePaperPt': p.get_size(), 'browserPaperPt': r.get_size(), 'paperPassed': paper, 'glyphPassed': painted,
                              'passed': pixels and painted and paper})
            finally:
                x.close(); y.close(); p.close(); q.close(); r.close()
        rows.append({'stage': printed, 'nativeSnapshot': native, 'paragraphProperties': properties, 'editablePages': dom, 'pageCounts': [len(a),len(b),len(c)], 'pages': pages,
                     'passed': native and dom and len(a)==len(b)==len(c) and all(p['passed'] for p in pages)})
report = {'passed': all(r['passed'] for r in rows), 'pdfKind': args.pdf_kind, 'nativeReceiptHash': sha(root / 'native-report.json'), 'comparerHash': sha(Path(__file__)), 'rows': rows}
suffix = '-download' if args.pdf_kind == 'download' else ''
output = f'negative-{args.negative_control}{suffix}.json' if args.negative_control else f'comparison{suffix}.json'
(root / output).write_text(json.dumps(report, indent=2)+'\n', encoding='utf-8')
print(json.dumps(report, indent=2))
sys.exit(0 if report['passed'] else 1)
