"""Saved-file rendering comparison; this does not certify browser line metrics."""
from pathlib import Path
import hashlib, json, sys
base = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
root = base / ('.local/word-font-cascade' if '--cascade' in sys.argv else '.local/word-line-metrics')
digest = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
receipt = json.loads((root / 'native-report.json').read_text(encoding='utf-8-sig'))
assert receipt['scriptHash'] == digest(base / 'scripts/verify-word-line-metrics.ps1')
for artifact in receipt['hashes']:
    assert digest(root / artifact['path']) == artifact['sha256'], 'Stale ' + artifact['path']
pages = []
for case in receipt['cases']:
    name = case['name']
    with pdfium.PdfDocument(root / ('expected-' + name + '.pdf')) as a, pdfium.PdfDocument(root / ('browser-' + name + '.pdf')) as b:
        assert len(a) == len(b) == case['expected']['pages']
        for i in range(len(a)):
            p, q = a[i], b[i]
            x, y = p.render(scale=2), q.render(scale=2)
            try:
                pages.append({'case': name, 'page': i + 1, 'exact': p.get_size() == q.get_size() and bytes(x.buffer) == bytes(y.buffer)})
            finally:
                x.close(); y.close(); p.close(); q.close()
report = {'passed': all(p['exact'] for p in pages), 'nativeReceiptHash': digest(root / 'native-report.json'), 'pages': pages, 'dpi': 144, 'scope': 'Native saved-file pixels; browser line metrics remain open'}
(root / 'pdf-report.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps(report))
sys.exit(0 if report['passed'] else 1)
