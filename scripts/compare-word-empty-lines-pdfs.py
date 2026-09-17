"""Compare independently edited Word copies with actual browser DOCX exports."""
from pathlib import Path
import hashlib
import json
import sys

base = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium

root = base / '.local/word-empty-lines'
digest = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
native = read(root / 'native-report.json')
browser = read(root / 'browser-report.json')
assert native['passed'] and browser['passed']
assert native['scriptSha256'] == digest(base / 'scripts/verify-word-empty-lines.ps1')
assert native['sourceSha256'] == browser['sourceSha256'] == digest(root / 'source.docx')
pages = []
artifacts = []
for stage in native['stages']:
    name = stage['stage']
    assert stage['expectedSha256'] == digest(root / f'expected-{name}.docx')
    assert stage['expectedPdfSha256'] == digest(root / f'expected-{name}.pdf')
    if name == 'source':
        continue
    assert stage['exportSha256'] == browser['exports'][name] == digest(root / f'browser-{name}.docx')
    assert stage['exportPdfSha256'] == digest(root / f'browser-{name}.pdf')
    for prefix in ('expected', 'browser'):
        for extension in ('docx', 'pdf'):
            p = root / f'{prefix}-{name}.{extension}'
            artifacts.append({'path': p.name, 'sha256': digest(p)})
    with pdfium.PdfDocument(root / f'expected-{name}.pdf') as a, pdfium.PdfDocument(root / f'browser-{name}.pdf') as b:
        assert len(a) == len(b)
        for i in range(len(a)):
            p, q = a[i], b[i]
            x, y = p.render(scale=2), q.render(scale=2)
            try:
                exact = p.get_size() == q.get_size() and bytes(x.buffer) == bytes(y.buffer)
                pages.append({'stage': name, 'page': i + 1, 'exact': exact})
                assert exact, f'{name} page {i + 1}'
            finally:
                x.close(); y.close(); p.close(); q.close()
report = {'passed': True, 'nativeSha256': digest(root / 'native-report.json'), 'browserSha256': digest(root / 'browser-report.json'), 'scriptSha256': digest(Path(__file__)), 'pages': pages, 'artifacts': artifacts}
(root / 'pdf-report.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(f'{len(pages)} empty-line native export PDF pages match exactly at 144 DPI.')
