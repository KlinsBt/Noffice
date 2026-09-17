"""Exact native rendering of real exports; browser baselines remain diagnostic."""
from pathlib import Path
import ctypes
import hashlib
import json
import sys
import re

base = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
import pypdfium2.raw as raw

probe = '--probe' in sys.argv
uniform = '--uniform-leading' in sys.argv
wrapped = '--wrapped-leading' in sys.argv
root = base / ('.local/word-wrapped-leading' if wrapped else '.local/word-uniform-leading' if uniform else '.local/word-baselines' if probe else '.local/word-mixed-lines')
digest = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
native = read(root / 'native-report.json')
browser = read(root / 'browser-report.json')
assert native['passed'] and browser['passed']
assert native['scriptSha256'] == digest(base / 'scripts/verify-word-mixed-lines.ps1')
assert native['sourceSha256'] == browser['sourceSha256'] == digest(root / 'source.docx')
for artifact in native['artifacts']:
    assert digest(root / artifact['path']) == artifact['sha256'], 'Stale ' + artifact['path']
pages = []
for stage in native['stages']:
    name = stage['stage']
    assert stage['exportSha256'] == browser['exports'][name] == digest(root / f'browser-{name}.docx')
    assert stage['expectedSha256'] == digest(root / f'expected-{name}.docx')
    with pdfium.PdfDocument(root / f'expected-{name}.pdf') as a, pdfium.PdfDocument(root / f'browser-{name}.pdf') as b:
        assert len(a) == len(b)
        for i in range(len(a)):
            p, q = a[i], b[i]
            x, y = p.render(scale=2), q.render(scale=2)
            try:
                pages.append({'stage': name, 'page': i+1, 'exact': p.get_size() == q.get_size() and bytes(x.buffer) == bytes(y.buffer)})
            finally:
                x.close(); y.close(); p.close(); q.close()
baselines = []
with pdfium.PdfDocument(root / 'source-native.pdf') as doc:
    page = doc[0]
    text = page.get_textpage()
    for i in range(text.count_chars()):
        if probe or uniform or wrapped or text.get_text_range(i, 1) != 'S' or len(baselines) == 6:
            continue
        matrix = raw.FS_MATRIX()
        assert raw.FPDFText_GetMatrix(text.raw, i, ctypes.byref(matrix))
        n = len(baselines)
        expected = page.get_height() - matrix.f - native['source'][n]['chars'][0]['y']
        actual = browser['baselines'][n]
        baselines.append({'paragraph': n+1, 'nativePt': expected, 'browserPt': actual, 'differencePt': actual-expected})
placements = []
if uniform:
    def target_offset(path, paragraphs):
        with pdfium.PdfDocument(path) as doc:
            page = doc[0]
            text = page.get_textpage()
            positions = [i for i in range(text.count_chars()) if text.get_text_range(i, 1) == 'S']
            assert len(positions) == 18
            matrix = raw.FS_MATRIX()
            assert raw.FPDFText_GetMatrix(text.raw, positions[9], ctypes.byref(matrix))
            return page.get_height() - matrix.f - paragraphs[9]['chars'][0]['y']
    initial = target_offset(root / 'source-native.pdf', native['source'])
    for stage in native['stages']:
        name = stage['stage']
        delta = target_offset(root / f'expected-{name}.pdf', stage['expected']) - initial
        actual = browser['stages'][name]['textTop'] - browser['initial']['textTop']
        placements.append({'stage': name, 'nativeDeltaPt': delta, 'browserDeltaPt': actual,
                           'tolerancePt': 0.15, 'passed': abs(actual-delta) < 0.15})
wrapped_lines = []
wrapped_placement = []
wrapped_diagnostics = []
if wrapped:
    def line_starts(paragraph):
        lines = []
        for i, c in enumerate(paragraph['chars']):
            if c['text'].isspace():
                continue
            if not lines or lines[-1]['y'] != c['y'] or lines[-1]['page'] != c['page']:
                lines.append({'from': i, 'text': '', 'y': c['y'], 'page': c['page']})
            lines[-1]['text'] += c['text']
        return lines
    def glyph_offsets(path, paragraphs):
        result = []
        with pdfium.PdfDocument(path) as doc:
            assert len(doc) == 1
            page = doc[0]
            text = page.get_textpage()
            content = text.get_text_range()
            for index in (1, 2):
                paragraph = paragraphs[index]
                offsets = []
                for line in line_starts(paragraph):
                    token = re.match(r'\w+', paragraph['text'][line['from']:]).group()
                    matches = list(re.finditer(r'\b'+re.escape(token)+r'\b', content))
                    assert len(matches) == 1
                    matrix = raw.FS_MATRIX()
                    assert raw.FPDFText_GetMatrix(text.raw, matches[0].start(), ctypes.byref(matrix))
                    offsets.append(page.get_height()-matrix.f-paragraph['chars'][0]['y'])
                result.append(offsets)
        return result
    source_offsets = glyph_offsets(root/'source-native.pdf', native['source'])
    cases = [('source', native['source'], browser['initial'])] + [(s['stage'], s['expected'], browser['stages'][s['stage']]) for s in native['stages']]
    for name, paragraphs, observed in cases:
        for index in (1, 2):
            expected = [{'from': l['from'], 'text': l['text']} for l in line_starts(paragraphs[index])]
            actual = [{'from': l['from'], 'text': l['text']} for l in observed[index-1]['lines']]
            wrapped_lines.append({'stage': name, 'paragraph': index+1, 'lines': len(expected), 'passed': expected == actual})
    for stage in native['stages']:
        name = stage['stage']
        offsets = glyph_offsets(root/f'expected-{name}.pdf', stage['expected'])
        for paragraph in range(2):
            for line, expected in enumerate(offsets[paragraph]):
                delta = expected-source_offsets[paragraph][line]
                actual = browser['stages'][name][paragraph]['lines'][line]['top']-browser['initial'][paragraph]['lines'][line]['top']
                row = {'stage': name, 'paragraph': paragraph+2, 'line': line+1, 'nativeDeltaPt': delta, 'browserDeltaPt': actual, 'differencePt': actual-delta, 'tolerancePt': 0.15, 'withinTolerance': abs(actual-delta)<0.15}
                # Only the first wrapped line has the same independently fixed
                # paragraph origin across all edits. Keep every other coordinate
                # diagnostic visible; their COM/PDF origin quantization is open.
                if paragraph == 0 and line == 0:
                    wrapped_placement.append(row)
                else:
                    wrapped_diagnostics.append(row)
report = {'passed': all(p['exact'] for p in pages) and all(p['passed'] for p in placements) and all(p['passed'] for p in wrapped_lines) and all(p['withinTolerance'] for p in wrapped_placement), 'scope': ('Exact native export pixels, line identities and first wrapped-line relative leading; remaining browser coordinates are diagnostic and open' if wrapped else 'Exact native export pixels; browser baseline differences are diagnostic and remain open'),
          'nativeReceiptSha256': digest(root / 'native-report.json'), 'browserReceiptSha256': digest(root / 'browser-report.json'),
          'wrappedLineIdentities': wrapped_lines, 'wrappedFirstLinePlacement': wrapped_placement, 'wrappedCoordinateDiagnostics': wrapped_diagnostics,
          'pages': pages, 'browserBaselines': baselines, 'browserLeadingPlacement': placements,
          'pdfHashes': {p.name: digest(p) for p in root.glob('*.pdf')}}
(root / 'pdf-report.json').write_text(json.dumps(report, indent=2)+'\n', encoding='utf-8')
print(json.dumps(report, indent=2))
sys.exit(0 if report['passed'] else 1)
