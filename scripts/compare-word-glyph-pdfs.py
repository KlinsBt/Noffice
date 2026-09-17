"""Compare real Chromium PDF glyph baselines with an independent Word oracle.

Paragraph backgrounds establish the same coordinate origin. This gates vertical
baselines and identities, not font outlines, horizontal positions or pagination.
Native shaded/unshaded matrix equality is independently checked by the inspector.
"""
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

sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
root = base / '.local/word-glyph'
mixed = base / '.local/word-mixed-wrap'
browser = read(mixed / 'browser-report.json')
native = read(mixed / 'native-report.json')
oracle = read(mixed / 'native-lines.json')
accepted = read(mixed / 'pdf-report.json')
fixture = read(base / 'tests/fixtures/native-word-mixed-wrap.json')
assert browser['passed'] and native['passed'] and accepted['passed']
assert oracle['stages'] == fixture['stages']
assert browser['sourceSha256'] == native['sourceSha256'] == oracle['sourceSha256'] == sha(base / 'tests/fixtures/word-mixed-wrap.docx')
assert oracle['nativeReceiptSha256'] == accepted['nativeReceiptSha256'] == sha(mixed / 'native-report.json')
assert accepted['browserReceiptSha256'] == sha(mixed / 'browser-report.json')
assert accepted['nativeLinesSha256'] == sha(mixed / 'native-lines.json')
assert accepted['scriptSha256'] == sha(base / 'scripts/compare-word-mixed-wrap-pdfs.py')
assert oracle['scriptSha256'] == sha(base / 'scripts/inspect-word-mixed-wrap.py')
assert native['scriptSha256'] == sha(base / 'scripts/verify-word-mixed-wrap.ps1')
assert oracle['probeSha256'] == sha(mixed / 'boxes/native-report.json')
for artifact in native['artifacts']:
    assert artifact['sha256'] == sha(mixed / artifact['path']), 'Stale native artifact ' + artifact['path']

negative = '--negative-control' in sys.argv
checks = []
for stage, paragraphs in oracle['stages'].items():
    path = root / f'{stage}.pdf'
    assert browser['prints'][stage] == sha(path), 'Stale browser print ' + stage
    with pdfium.PdfDocument(path) as doc:
        assert len(doc) == 1
        page = doc[0]
        text = page.get_textpage()
        try:
            for index, prefix, rgb in [(0, 'alpha', [0, 255, 255]), (1, 'beta', [255, 0, 255])]:
                bounds = []
                for obj in page.get_objects():
                    color = [ctypes.c_uint() for _ in range(4)]
                    if raw.FPDFPageObj_GetFillColor(obj.raw, *[ctypes.byref(c) for c in color]) and [c.value for c in color[:3]] == rgb:
                        bounds.append(obj.get_bounds())
                assert bounds, (stage, prefix, 'No paragraph background')
                top = page.get_height() - max(b[3] for b in bounds)
                lines = []
                for match in re.finditer(prefix + r'\d+', text.get_text_range()):
                    matrix = raw.FS_MATRIX()
                    assert raw.FPDFText_GetMatrix(text.raw, match.start(), ctypes.byref(matrix))
                    baseline = page.get_height() - matrix.f - top
                    if not lines or abs(lines[-1]['baseline'] - baseline) > 0.15:
                        lines.append({'baseline': baseline, 'words': []})
                    lines[-1]['words'].append(match.group())
                expected = paragraphs[index]['lines']
                assert sum(len(line['words']) for line in lines) == 20
                assert [line['words'] for line in lines] == [line['words'] for line in expected], (stage, prefix, 'Changed line identities')
                # Perturb only an in-memory coordinate; preserve canonical PDFs.
                if negative and stage == 'source' and index == 0:
                    lines[0]['baseline'] += 0.3
                for number, (actual, reference) in enumerate(zip(lines, expected)):
                    absolute = actual['baseline'] - reference['baseline']
                    relative = (actual['baseline'] - lines[0]['baseline']) - (reference['baseline'] - expected[0]['baseline'])
                    checks.append({'stage': stage, 'paragraph': index + 2, 'line': number + 1, 'words': actual['words'], 'absoluteDifferencePt': absolute, 'relativeDifferencePt': relative, 'tolerancePt': 0.15, 'passed': abs(absolute) < 0.15 and abs(relative) < 0.15})
        finally:
            text.close()
            page.close()
report = {'passed': all(c['passed'] for c in checks), 'negativeControl': negative, 'scope': __doc__.strip(), 'checks': checks, 'prints': browser['prints'], 'browserReceiptSha256': sha(mixed / 'browser-report.json'), 'nativeReceiptSha256': sha(mixed / 'native-report.json'), 'nativeLinesSha256': sha(mixed / 'native-lines.json'), 'exportComparisonSha256': sha(mixed / 'pdf-report.json'), 'scriptSha256': sha(Path(__file__))}
(root / ('negative-control.json' if negative else 'pdf-report.json')).write_text(json.dumps(report, indent=2) + '\n')
print(json.dumps({'passed': report['passed'], 'printedPages': len(browser['prints']), 'lineBaselineChecks': len(checks), 'maxAbsoluteDifferencePt': max(abs(c['absoluteDifferencePt']) for c in checks), 'maxRelativeDifferencePt': max(abs(c['relativeDifferencePt']) for c in checks)}, indent=2))
sys.exit(0 if report['passed'] else 1)
