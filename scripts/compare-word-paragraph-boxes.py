"""Validate native paragraph origins without mixing COM range and PDF coordinates."""
from pathlib import Path
import ctypes
import hashlib
import json
import re
import sys

base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
import pypdfium2.raw as raw

root = base / '.local/word-paragraph-boxes'
wrapped = base / '.local/word-wrapped-leading'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
probe = read(root / 'native-report.json')
native = read(wrapped / 'native-report.json')
browser = read(wrapped / 'browser-report.json')
assert native['passed'] and browser['passed']
assert probe['sourceSha256'] == native['sourceSha256'] == browser['sourceSha256'] == sha(base / 'tests/fixtures/word-wrapped-leading.docx')
assert probe['scriptSha256'] == sha(base / 'scripts/probe-word-paragraph-boxes.ps1')
assert native['scriptSha256'] == sha(base / 'scripts/verify-word-mixed-lines.ps1')
baseline = read(base / 'docs/parity/baseline.json')
assert probe['executableSha256'] == native['executableSha256'] == next(a['sha256'] for a in baseline['applications'] if a['executable'] == 'WINWORD.EXE')
for artifact in native['artifacts']:
    assert sha(wrapped / artifact['path']) == artifact['sha256'], 'Stale ' + artifact['path']
for stage in native['stages']:
    assert stage['exportSha256'] == browser['exports'][stage['stage']] == sha(wrapped / ('browser-' + stage['stage'] + '.docx'))
negative = '--negative-control' in sys.argv
if negative:
    # Deliberate in-memory displacement; canonical browser evidence stays intact.
    browser['stages']['double'][1]['lines'][2]['top'] += 0.3

def matrices(text):
    result = []
    for i in range(text.count_chars()):
        matrix = raw.FS_MATRIX()
        if raw.FPDFText_GetMatrix(text.raw, i, ctypes.byref(matrix)):
            result.append((text.get_text_range(i, 1), *[getattr(matrix, key) for key in ('a','b','c','d','e','f')]))
    return result

origins, offsets, checks, artifacts = {}, {}, [], []
for stage in probe['stages']:
    name = stage['stage']
    for suffix, key in [('.docx','docxSha256'),('.pdf','pdfSha256'),('-boxes.pdf','boxesPdfSha256')]:
        path = root / (name + suffix)
        assert stage[key] == sha(path)
        artifacts.append({'path':path.name,'sha256':sha(path)})
    observed = browser['initial'] if name == 'source' else browser['stages'][name]
    paragraphs = native['source'] if name == 'source' else next(s['expected'] for s in native['stages'] if s['stage'] == name)
    counterpart = wrapped / ('source-native.pdf' if name == 'source' else f'expected-{name}.pdf')
    with pdfium.PdfDocument(root / f'{name}.pdf') as plain, pdfium.PdfDocument(root / f'{name}-boxes.pdf') as marked, pdfium.PdfDocument(counterpart) as original:
        assert len(plain) == len(marked) == len(original) == 1
        p, q, saved = plain[0], marked[0], original[0]
        text, marked_text, saved_text = p.get_textpage(), q.get_textpage(), saved.get_textpage()
        assert matrices(text) == matrices(marked_text) == matrices(saved_text), 'Shading or the probe changed native glyph layout'
        boxes = {2:[], 3:[]}
        for obj in q.get_objects():
            color = [ctypes.c_uint() for _ in range(4)]
            if raw.FPDFPageObj_GetFillColor(obj.raw, *[ctypes.byref(c) for c in color]):
                rgb = [c.value for c in color[:3]]
                paragraph = 2 if rgb == [0,255,255] else 3 if rgb == [255,0,255] else None
                if paragraph:
                    boxes[paragraph].append(q.get_height() - obj.get_bounds()[3])
        origins[name] = {str(i):min(boxes[i]) for i in (2,3)}
        content = text.get_text_range()
        offsets[name] = []
        for index in (1,2):
            lines = observed[index-1]['lines']
            assert len(boxes[index+1]) == len(lines)
            baselines = []
            for line in lines:
                token = re.match(r'\w+', paragraphs[index]['text'][line['from']:]).group()
                matches = list(re.finditer(r'\b' + re.escape(token) + r'\b', content))
                assert len(matches) == 1
                matrix = raw.FS_MATRIX()
                assert raw.FPDFText_GetMatrix(text.raw, matches[0].start(), ctypes.byref(matrix))
                baselines.append(p.get_height()-matrix.f)
            offsets[name].append(baselines[0]-origins[name][str(index+1)])
            for i, baseline in enumerate(baselines):
                expected = baseline-baselines[0]
                actual = lines[i]['top']-lines[0]['top']
                checks.append({'stage':name,'paragraph':index+1,'line':i+1,'nativeAdvancePt':expected,'browserAdvancePt':actual,'differencePt':actual-expected,'tolerancePt':0.15,'passed':abs(actual-expected)<0.15})

leading = []
for name in ('double','minimum','single'):
    for index in range(2):
        expected = offsets[name][index]-offsets['source'][index]
        actual = browser['stages'][name][index]['lines'][0]['top']-browser['initial'][index]['lines'][0]['top']
        leading.append({'stage':name,'paragraph':index+2,'nativeShiftPt':expected,'browserShiftPt':actual,'differencePt':actual-expected,'tolerancePt':0.15,'passed':abs(actual-expected)<0.15})
report = {'passed':all(c['passed'] for c in checks+leading),'scope':'Per-state line advances and relative leading from layout-neutral native PDF paragraph backgrounds; absolute browser baselines remain open','negativeControl':negative,'nativeGlyphMatricesUnchanged':True,'nativeOrigins':origins,'lineAdvances':checks,'leading':leading,'artifacts':artifacts,'nativeProbeSha256':sha(root/'native-report.json'),'nativeReceiptSha256':sha(wrapped/'native-report.json'),'browserReceiptSha256':sha(wrapped/'browser-report.json'),'scriptSha256':sha(Path(__file__))}
(root/('negative-control.json' if negative else 'comparison.json')).write_text(json.dumps(report,indent=2)+'\n')
print(json.dumps({'passed':report['passed'],'lineChecks':len(checks),'leadingChecks':len(leading),'maximumDifferencePt':max(abs(c['differencePt']) for c in checks+leading),'failures':[c for c in checks+leading if not c['passed']]},indent=2))
sys.exit(0 if report['passed'] else 1)
