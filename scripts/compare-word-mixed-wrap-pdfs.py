"""Gate bounded mixed-wrap line boxes, glyph baselines and actual DOCX exports."""
from pathlib import Path
import hashlib
import json
import sys

base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium

root = base / '.local/word-mixed-wrap'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
native = read(root / 'native-report.json')
browser = read(root / 'browser-report.json')
oracle = read(root / 'native-lines.json')
fixture = read(base / 'tests/fixtures/native-word-mixed-wrap.json')
assert native['passed'] and browser['passed']
assert native['sourceSha256'] == browser['sourceSha256'] == oracle['sourceSha256'] == sha(base / 'tests/fixtures/word-mixed-wrap.docx')
assert native['scriptSha256'] == sha(base / 'scripts/verify-word-mixed-wrap.ps1')
assert oracle['scriptSha256'] == sha(base / 'scripts/inspect-word-mixed-wrap.py')
assert oracle['nativeReceiptPath'] == 'native-report.json'
assert oracle['nativeReceiptSha256'] == sha(root / 'native-report.json')
assert oracle['probeSha256'] == sha(root / 'boxes/native-report.json')
assert oracle['stages'] == fixture['stages'], 'Regenerate/review the changed native oracle before acceptance'
for artifact in native['artifacts']:
    assert sha(root / artifact['path']) == artifact['sha256'], 'Stale ' + artifact['path']
pages = []
for stage in native['stages']:
    name = stage['stage']
    assert stage['exportSha256'] == browser['exports'][name] == sha(root / f'browser-{name}.docx')
    assert stage['expectedSha256'] == sha(root / f'expected-{name}.docx')
    with pdfium.PdfDocument(root / f'expected-{name}.pdf') as a, pdfium.PdfDocument(root / f'browser-{name}.pdf') as b:
        assert len(a) == len(b)
        for index in range(len(a)):
            p, q = a[index], b[index]
            x, y = p.render(scale=2), q.render(scale=2)
            try:
                pages.append({'stage':name,'page':index+1,'exact':p.get_size()==q.get_size() and bytes(x.buffer)==bytes(y.buffer)})
            finally:
                x.close(); y.close(); p.close(); q.close()
assert native['nativeReedit']['passed']
assert native['nativeReedit']['inputExpectedSha256'] == sha(root/'expected-typing.docx')
assert native['nativeReedit']['inputBrowserSha256'] == sha(root/'browser-typing.docx')
with pdfium.PdfDocument(root/'native-reedit-expected.pdf') as a, pdfium.PdfDocument(root/'native-reedit-browser.pdf') as b:
    assert len(a) == len(b) == 1
    p, q = a[0], b[0]
    x, y = p.render(scale=2), q.render(scale=2)
    try:
        pages.append({'stage':'native-reedit','page':1,'exact':p.get_size()==q.get_size() and bytes(x.buffer)==bytes(y.buffer)})
    finally:
        x.close(); y.close(); p.close(); q.close()
negative = '--negative-control' in sys.argv
if negative:
    browser['initial'][0]['height'] += 0.3
checks, diagnostics, line_boxes = [], [], []
for name, expected in oracle['stages'].items():
    actual = browser['initial'] if name == 'source' else browser['stages'][name]
    for index, (p, q) in enumerate(zip(expected, actual)):
        difference = q['height']-p['height']
        identities = [l['words'] for l in p['lines']] == [l['words'] for l in q['lines']]
        checks.append({'stage':name,'paragraph':index+2,'lines':len(p['lines']),'identitiesMatch':identities,'heightDifferencePt':difference,'tolerancePt':0.15,'passed':identities and abs(difference)<0.15})
        assert len(q['boxes']) == len(p['lines'])
        for number, (line, box) in enumerate(zip(p['lines'], q['boxes'])):
            top = box['top']-line['boxTop']
            height = box['height']-line['boxHeight']
            line_boxes.append({'stage':name,'paragraph':index+2,'line':number+1,'topDifferencePt':top,'heightDifferencePt':height,'tolerancePt':0.15,'passed':abs(top)<0.15 and abs(height)<0.15})
        for number, (x, y) in enumerate(zip(p['lines'], q['lines'])):
            difference = y['baseline']-x['baseline']
            relative = (y['baseline']-q['lines'][0]['baseline'])-(x['baseline']-p['lines'][0]['baseline'])
            diagnostics.append({'stage':name,'paragraph':index+2,'line':number+1,'absoluteDifferencePt':difference,'relativeDifferencePt':relative,'tolerancePt':0.15,'withinTolerance':abs(difference)<0.15 and abs(relative)<0.15})
report = {'negativeControl':negative,'passed':all(c['passed'] for c in checks) and all(p['exact'] for p in pages) and all(l['passed'] for l in line_boxes) and all(d['withinTolerance'] for d in diagnostics),'scope':'Bounded regular Arial mixed-size soft-line identities, paragraph heights and paragraph-relative absolute/relative glyph baselines; exact native-rendered actual DOCX exports. Other fonts, pagination and glyph outlines remain open.','paragraphs':checks,'lineBoxes':line_boxes,'pages':pages,'glyphChecks':diagnostics,'nativeReceiptSha256':sha(root/'native-report.json'),'browserReceiptSha256':sha(root/'browser-report.json'),'nativeLinesSha256':sha(root/'native-lines.json'),'scriptSha256':sha(Path(__file__))}
(root/('negative-control.json' if negative else 'pdf-report.json')).write_text(json.dumps(report,indent=2)+'\n')
print(json.dumps({'passed':report['passed'],'paragraphChecks':len(checks),'lineIdentities':sum(c['lines'] for c in checks),'lineBoxChecks':len(line_boxes),'exactPages':sum(p['exact'] for p in pages),'glyphChecks':len(diagnostics),'failedGlyphChecks':sum(not d['withinTolerance'] for d in diagnostics)},indent=2))
sys.exit(0 if report['passed'] else 1)
