"""Compare actual native/browser flow PDFs at every authored glyph origin.

This gate supplements exact native DOCX-export pixels and browser page/text
checks. It does not certify browser glyph outlines or unmeasured fonts.
"""
from pathlib import Path
import argparse
import hashlib
import json
import sys

base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
from word_pdf_glyphs import glyphs


parser = argparse.ArgumentParser()
parser.add_argument('--family', required=True, choices=['section', 'inline', 'forced', 'ui'])
parser.add_argument('--case', required=True)
parser.add_argument('--negative-control', choices=['baseline', 'color', 'size'])
parser.add_argument('--zoom', type=int, choices=[50, 150])
parser.add_argument('--transition', action='store_true')
parser.add_argument('--pdf-kind', choices=['print','download'], default='print')
args = parser.parse_args()
root = base / f'.local/word-{args.family}-acceptance' / args.case
if args.transition:
    assert args.family == 'section' and args.case.startswith(('oddPage-', 'evenPage-'))
    root = root.with_name('transition-' + args.case)
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
read = lambda path: json.loads(path.read_text(encoding='utf-8-sig'))
prefix = 'word-forced-keep' if args.family == 'forced' else f'word-{args.family}-flow'
reference = base / f'tests/fixtures/native-{prefix}.json'
assert args.case in read(reference)['cases'], 'Unknown flow fixture'
receipt_path = root / 'native-report.json'
receipt = read(receipt_path)
assert receipt['passed'] and receipt['scriptHash'] == sha(base / 'scripts/verify-word-section-flow.ps1')
assert receipt['sourceHash'] == sha(base / f'tests/fixtures/{prefix}-{args.case}.docx')
assert receipt['referenceHash'] == sha(reference)
for binding in receipt['hashes']:
    assert sha(root / binding['path']) == binding['sha256'], 'Stale ' + binding['path']
if args.zoom:
    zoom = next(row for row in read(root / 'browser-report.json')['zoomPrints'] if row['zoom'] == args.zoom)
    for stage, prefix in [('source', ''), ('edited', 'edited-')]:
        assert sha(root / f'browser-{prefix}print-zoom{args.zoom}.pdf') == zoom[f'{stage}Hash'], 'Stale zoom PDF'

rows, control = [], False
for stage, native_name, browser_name in [
    ('source', 'source-native.pdf', 'browser-print.pdf'),
    ('edited', 'expected-edited.pdf', 'browser-edited-print.pdf'),
]:
    if args.pdf_kind == 'download':
        assert not args.zoom, 'Download PDF uses physical layout, not print zoom variants'
        browser_name = f'download-{stage}.pdf'
        browser_receipt = read(root/'browser-report.json')
        assert sha(root/browser_name) == browser_receipt['pdfDownloadHash' if stage=='source' else 'editedPdfDownloadHash'], 'Stale PDF download'
    if args.zoom:
        browser_name = browser_name.replace('.pdf', f'-zoom{args.zoom}.pdf')
    with pdfium.PdfDocument(root / native_name) as native, pdfium.PdfDocument(root / browser_name) as browser:
        assert len(native) == len(browser), f'{stage}: page count differs'
        for index in range(len(native)):
            a, b = native[index], browser[index]
            try:
                if args.pdf_kind == 'download':
                    assert all(abs(x-y)<=.15 for x,y in zip(a.get_size(), b.get_size())), 'PDF paper differs'
                no_spaces = not any(c['text'] == ' ' for p in receipt[stage]['paragraphs']
                                    for c in p['characters'] if c['page'] == index + 1)
                expected, actual = glyphs(a, no_spaces), glyphs(b, no_spaces)
                assert [g['text'] for g in expected] == [g['text'] for g in actual], f'{stage}/{index+1}: glyph sequence differs'
                pairs = [(n, c) for n, c in zip(expected, actual) if not (n.get('generated') or c.get('generated'))]
                if args.negative_control and pairs and not control:
                    if args.negative_control == 'baseline': pairs[0][1]['y'] += 1
                    if args.negative_control == 'size': pairs[0][1]['size'] += 1
                    if args.negative_control == 'color': pairs[0][1]['color'][0] ^= 1
                    control = True
                errors = {key: max((abs(n[key] - c[key]) for n, c in pairs), default=0)
                          for key in ['x', 'y', 'size']}
                assert all(error <= .15 for error in errors.values()), f'{stage}/{index+1}: glyph geometry differs: {errors}'
                assert all(n['color'] == c['color'] for n, c in pairs), f'{stage}/{index+1}: glyph color differs'
                rows.append({'stage': stage, 'page': index+1, 'glyphs': len(pairs), 'reconstructedSpaces': len(expected)-len(pairs), 'maxErrorPt': errors})
            finally:
                a.close(); b.close()
assert not args.negative_control, 'Negative control did not exercise a glyph'
report = {'passed': True, 'pdfKind': args.pdf_kind, 'nativeReceiptHash': sha(receipt_path), 'scriptHash': sha(Path(__file__)),
          'glyphHelperHash': sha(base / 'scripts/word_pdf_glyphs.py'),
          'scope': 'Exact text sequence; painted glyph color, physical X/Y baseline and effective font size within .15pt. Reconstructed spaces retain text identity and are measured through adjacent painted glyph positions. Browser glyph outlines remain open.',
          'pages': rows}
(root / ('glyph-download-report.json' if args.pdf_kind=='download' else f'glyph-zoom{args.zoom}-report.json' if args.zoom else 'glyph-report.json')).write_text(json.dumps(report, indent=2) + '\n', encoding='utf8')
print(json.dumps({'passed': True, 'pages': len(rows), 'glyphs': sum(r['glyphs'] for r in rows)}))
