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
root = ROOT / '.local/word-story-create'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
prefix = 'native-export'
receipt = json.loads((root / 'native-export-report.json').read_text(encoding='utf-8-sig'))
assert receipt['scriptHash'] == sha(ROOT / 'scripts/verify-word-story-create.ps1'), 'Stale native verifier'
baseline = json.loads((ROOT / 'docs/parity/baseline.json').read_text(encoding='utf-8-sig'))
assert receipt['executableHash'] == next(a['sha256'] for a in baseline['applications'] if a['executable'] == 'WINWORD.EXE')
for entry in receipt['hashes']:
    assert sha(root / entry['path']) == entry['sha256'], f"Stale artifact: {entry['path']}"
    if entry['path'].endswith('/browser-report.json'):
        check_build(json.loads((root / entry['path']).read_text(encoding='utf-8-sig')))
reference = json.loads((ROOT / 'tests/fixtures/native-word-story-create.json').read_text())
assert receipt['referenceHash'] == sha(ROOT / 'tests/fixtures/native-word-story-create.json'), 'Stale native matrix'
for probe in reference['probes']:
    assert probe['nativeReceiptHash'] == sha(ROOT / '.local' / probe['root'] / 'native-discovery.json')
    assert probe['nativeScriptHash'] == sha(ROOT / '.local' / ('probe-' + probe['root'] + '.ps1'))
    assert probe['executableHash'] == receipt['executableHash']
assert reference['activationProbe']['nativeReceiptHash'] == sha(ROOT / '.local/word-story-create-activate/native-discovery.json')
assert len(reference['cases']) == 54 and len(receipt['rows']) == 224, 'Incomplete creation matrix'
expected_rows = set()
for name, sample in reference['cases'].items():
    browser = json.loads((root / f'browser/{name}/browser-report.json').read_text())
    assert not browser['errors']
    source = reference['sources'][sample['mode']]
    assert sha(ROOT / 'tests/fixtures' / source['file']) == source['sha256'] == browser['hashes']['source']
    stages = ['blank', 'typed', 'reloaded', 'reimported'] + (['activated'] if 'activation' in sample else [])
    for stage in stages:
        expected_rows.add((name, stage))
        wanted = sample['activation']['native'] if stage == 'activated' else sample['blank' if stage == 'blank' else 'typed']
        matching = [r for r in receipt['rows'] if r['name'] == name and r['stage'] == stage]
        assert len(matching) == 1 and matching[0]['expected'] == wanted, 'Independent properties changed'
assert {(r['name'], r['stage']) for r in receipt['rows']} == expected_rows

from zipfile import ZipFile
import xml.etree.ElementTree as ET
w = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
rel = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
def story_references(path):
    with ZipFile(path) as z:
        main = ET.fromstring(z.read('word/document.xml'))
        relationships = ET.fromstring(z.read('word/_rels/document.xml.rels'))
        targets = {r.attrib['Id']: r.attrib['Target'] for r in relationships}
        sections, previous = [], {}
        for section in main.iter(w + 'sectPr'):
            current = {}
            for kind in ['header', 'footer']:
                for slot in ['default', 'first', 'even']:
                    key = kind + ':' + slot
                    direct = next((r for r in section.findall(w + kind + 'Reference') if r.attrib.get(w + 'type') == slot), None)
                    if direct is not None:
                        path = 'word/' + targets[direct.attrib[rel + 'id']]
                        value = ''.join(t.text or '' for t in ET.fromstring(z.read(path)).iter(w + 't'))
                        current[key] = {'text': value, 'linked': False}
                    else:
                        current[key] = {'text': previous.get(key, {}).get('text', ''), 'linked': bool(sections)}
            sections.append(current); previous = current
    return sections
results = []
returns = []
if args.require_native_return:
    assert len(receipt['returns']) == 54, 'Missing native re-edit cases'
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
    sample = reference['cases'][name]
    expected_refs = sample['activation']['references'] if stage == 'activated' else sample['references']['blank' if stage == 'blank' else 'typed']
    references_match = story_references(root / row['folder'] / f'{stage}.docx') == expected_refs
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
    comparisons = [(f'{stage}.pdf', f'expected-{stage}.pdf')] if stage in ['typed', 'activated'] else []
    for pdf_name, reference_name in comparisons:
        expected_path = root / row['folder'] / reference_name if stage in ['typed', 'activated'] else root / reference_name
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
    results.append({'name': name, 'stage': stage, 'properties': wanted == actual, 'references': references_match, 'pages': native_pages, 'downloadedPdf': downloaded, 'passed': wanted == actual and references_match and pixels_pass})
recovery = []
assert {r['name'] for r in receipt['recovery']} == {'late-abort', 'other-tab'}, 'Missing creation recovery matrix'
for row in receipt['recovery']:
    sample = reference['cases'][row['expectedCase']]
    references_match = story_references(root / row['folder'] / 'unsaved.docx') == sample['references']['typed']
    browser = json.loads((root / row['folder'] / 'browser-report.json').read_text())
    assert not browser['errors']
    assert browser['sourceHash'] == reference['sources']['missing-styles']['sha256']
    assert browser['exportHash'] == sha(root / row['folder'] / 'unsaved.docx')
    pages = []
    with pdfium.PdfDocument(root / 'browser' / row['expectedCase'] / 'expected-typed.pdf') as a, pdfium.PdfDocument(root / row['folder'] / 'native-unsaved.pdf') as b:
        for index in range(min(len(a), len(b))):
            p, q = a[index], b[index]; x, y = p.render(scale=2), q.render(scale=2)
            try: pages.append(p.get_size() == q.get_size() and bytes(x.buffer) == bytes(y.buffer))
            finally: x.close(); y.close(); p.close(); q.close()
        passed = len(a) == len(b) == len(pages) and all(pages) and row['actual'] == sample['typed'] and references_match
    recovery.append({'name':row['name'], 'pages':pages, 'properties':row['actual'] == sample['typed'], 'references':references_match, 'passed':passed})
report = {'passed': bool(results) and all(r['passed'] for r in results) and all(r['passed'] for r in recovery), 'rows': results, 'recovery':recovery, 'nativeReturns': returns, 'negativeControl': args.negative_control, 'nativeReceiptSha256': sha(root / 'native-export-report.json')}
if not args.negative_control:
    (root / f'{prefix}-comparison.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
print(json.dumps(report, indent=2))
sys.exit(0 if report['passed'] else 1)
