"""Check actual DOCX story semantics and exact native rendering against separate edits."""
import argparse
import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '.local/pdf-tools'))
import pypdfium2 as pdfium
from word_pdf_glyphs import glyphs
from word_story_artifacts import check_build

parser = argparse.ArgumentParser()
parser.add_argument('--negative-control', choices=['text', 'property', 'pixel', 'glyph'])
parser.add_argument('--require-native-return', action='store_true')
args = parser.parse_args()
root = ROOT / '.local/word-story-options'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
prefix = 'native-export'
receipt = json.loads((root / 'native-export-report.json').read_text(encoding='utf-8-sig'))
assert receipt['scriptHash'] == sha(ROOT / 'scripts/verify-word-story-options.ps1'), 'Stale native verifier'
baseline = json.loads((ROOT / 'docs/parity/baseline.json').read_text(encoding='utf-8-sig'))
assert receipt['executableHash'] == next(a['sha256'] for a in baseline['applications'] if a['executable'] == 'WINWORD.EXE')
for entry in receipt['hashes']:
    assert sha(root / entry['path']) == entry['sha256'], f"Stale artifact: {entry['path']}"
    if entry['path'].endswith('/browser-report.json'):
        check_build(json.loads((root / entry['path']).read_text(encoding='utf-8-sig')))
reference = json.loads((ROOT / 'tests/fixtures/native-word-story-options.json').read_text(encoding='utf-8'))
assert len(reference['cases']) == 12 and len(receipt['rows']) == 48, 'Incomplete section matrix'
assert reference['nativeReceiptHash'] == sha(root / 'native-verified.json'), 'Changed discovery receipt'
assert reference['nativeScriptHash'] == sha(ROOT / '.local/reverify-word-story-options.ps1'), 'Changed discovery script'
for name, sample in reference['cases'].items():
    browser = json.loads((root / f'browser/{name}/browser-report.json').read_text(encoding='utf-8'))
    assert not browser['errors']
    assert sha(ROOT / 'tests/fixtures' / reference['sourceFile']) == reference['sourceSha256'] == browser['hashes']['source']
    assert sha(root / f'verified-{name}.pdf') == sample['nativePdfSha256']
assert {(r['name'], r['stage']) for r in receipt['rows']} == {(n,s) for n in reference['cases'] for s in ['undo','edited','reloaded','restored']}
assert reference['initialReceiptHash'] == sha(root / 'native-discovery.json'), 'Initial discovery changed'
for supplemental in reference['supplementalReceipts']:
    assert sha(ROOT / supplemental['path']) == supplemental['sha256']
    assert sha(ROOT / supplemental['scriptPath']) == supplemental['scriptHash']
distances = json.loads((ROOT / 'tests/fixtures/native-word-story-distances.json').read_text(encoding='utf-8'))
assert distances['executableHash'] == reference['executableHash']
observations = []
for binding in distances['receipts']:
    assert sha(ROOT / binding['path']) == binding['sha256'], 'Distance receipt changed'
    assert sha(ROOT / binding['scriptPath']) == binding['scriptHash'], 'Distance probe changed'
    captured = json.loads((ROOT / binding['path']).read_text(encoding='utf-8-sig'))
    assert captured['executableHash'] == distances['executableHash']
    assert captured['scriptHash'] == binding['scriptHash']
    observations.extend({key: row[key] for key in ('kind', 'input', 'accepted', 'observed')} for row in captured['rows'])
assert observations == distances['cases'] and len(observations) == 42, 'Incomplete native distance observations'
results = []
returns = []
if args.require_native_return:
    assert len(receipt['returns']) == 12, 'Missing native re-edit cases'
    for row in receipt['returns']:
        folder = root / row['folder']
        report = json.loads((folder / 'return-browser-report.json').read_text(encoding='utf-8-sig'))
        check_build(report)
        assert report['nativeReceiptHash'] == sha(root / 'native-export-report.json'), 'Stale native return receipt'
        assert not report['errors'], 'Native reimport reported browser errors'
        assert sha(folder / 'native-return.docx') == sha(folder / 'returned.docx') == report['sourceHash'], 'Native re-edit was not preserved'
        returns.append({'name': row['name'], 'passed': True, 'browserReportSha256': sha(folder / 'return-browser-report.json')})
for row_index, row in enumerate(receipt['rows']):
    name, stage = row['name'], row['stage']
    wanted, actual = row['expected'], row['actual']
    if row_index == 0 and args.negative_control == 'text':
        actual['body'] += 'invalid'
    if row_index == 0 and args.negative_control == 'property':
        actual['sections'][0]['header'] += 1
    pages = []
    with pdfium.PdfDocument(root / row['folder'] / f'expected-{stage}.pdf') as a, pdfium.PdfDocument(root / row['folder'] / f'native-{stage}.pdf') as b:
        for index in range(min(len(a), len(b))):
            p, q = a[index], b[index]
            x, y = p.render(scale=2), q.render(scale=2)
            try:
                pixels = bytearray(y.buffer)
                if args.negative_control == 'pixel' and row_index == 0 and index == 0:
                    pixels[0] ^= 1
                pages.append(p.get_size() == q.get_size() and bytes(x.buffer) == bytes(pixels))
            finally:
                x.close(); y.close(); p.close(); q.close()
        pixels_pass = len(a) == len(b) == len(pages) and len(a) == wanted['pages'] and len(b) == actual['pages'] and all(pages)
    native_pages = pages
    downloaded = []
    comparisons = [(f'{stage}.pdf', f'expected-{stage}.pdf')] if stage in ['edited','restored'] else []
    for pdf_name, reference_name in comparisons:
        expected_path = root / row['folder'] / reference_name if stage in ['edited','restored'] else root / reference_name
        pages = []
        with pdfium.PdfDocument(expected_path) as a, pdfium.PdfDocument(root / row['folder'] / pdf_name) as b:
            for index in range(min(len(a), len(b))):
                p, q = a[index], b[index]
                try:
                    first, last = glyphs(p, False), glyphs(q, False)
                    pairs = [(g,h) for g,h in zip(first,last) if not(g.get('generated') or h.get('generated'))]
                    if args.negative_control == 'glyph' and index == 0 and pairs: pairs[0][1]['y'] += 1
                    differences = {axis: max((abs(g[axis]-h[axis]) for g,h in pairs), default=0) for axis in ['x','y','size']}
                    text = [g['text'] for g in first] == [g['text'] for g in last]
                    paper = all(abs(u-v) <= .15 for u,v in zip(p.get_size(),q.get_size()))
                    color = all(g['color'] == h['color'] for g,h in pairs)
                    pages.append({'page':index+1,'glyphs':len(pairs),'differences':differences,'text':text,'paper':paper,'color':color,'passed':text and paper and color and all(v <= .15 for v in differences.values())})
                finally: p.close(); q.close()
            passed = len(a) == len(b) == len(pages) and all(p['passed'] for p in pages)
            pixels_pass = pixels_pass and passed
        downloaded.append({'file': pdf_name, 'pages': pages, 'passed': passed})
    results.append({'name': name, 'stage': stage, 'properties': wanted == actual, 'pages': native_pages, 'downloadedPdf': downloaded, 'passed': wanted == actual and pixels_pass})
report = {'passed': bool(results) and all(r['passed'] for r in results), 'rows': results, 'nativeReturns': returns, 'negativeControl': args.negative_control, 'nativeReceiptSha256': sha(root / 'native-export-report.json')}
if not args.negative_control:
    (root / f'{prefix}-comparison.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
print(json.dumps(report, indent=2))
sys.exit(0 if report['passed'] else 1)
