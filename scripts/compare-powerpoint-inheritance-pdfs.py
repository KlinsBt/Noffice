"""Compare native rendering of the actual browser export to independent native edits."""
import hashlib
import argparse
import json
import sys
from pathlib import Path

root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(root / '.local/pdf-tools'))
import pypdfium2 as pdfium
parser = argparse.ArgumentParser()
parser.add_argument('--fixture', choices=['inheritance', 'groups', 'group-edits', 'group-bounds'], default='inheritance')
fixture = parser.parse_args().fixture
directory = root / f'.local/powerpoint-{fixture}'
pages = 2 if fixture == 'inheritance' else 4
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
receipt = json.loads((directory / 'native-report.json').read_text(encoding='utf8'))
assert receipt['passed']
assert receipt['scriptSha256'] == sha(root / f'scripts/verify-powerpoint-{fixture}.ps1')
assert receipt['exportSha256'] == sha(directory / 'edited.pptx')
assert receipt['expectedSha256'] == sha(directory / 'native-expected.pptx')
expected_path = directory / 'native-expected.pdf'
actual_path = directory / 'browser-export.pdf'
with pdfium.PdfDocument(expected_path) as expected, pdfium.PdfDocument(actual_path) as actual:
    assert len(expected) == len(actual) == pages
    for index in range(pages):
        p, q = expected[index], actual[index]
        left, right = p.render(scale=2), q.render(scale=2)
        try:
            assert p.get_size() == q.get_size() and bytes(left.buffer) == bytes(right.buffer), f'Native rendered slide {index + 1} differs'
        finally:
            left.close(); right.close(); p.close(); q.close()
report = {'passed': True, 'pages': pages, 'dpi': 144, 'scope': 'Exact native-rendered exported slides; not browser glyph/layout parity', 'nativeReceiptSha256': sha(directory / 'native-report.json'), 'expectedPdfSha256': sha(expected_path), 'actualPdfSha256': sha(actual_path)}
(directory / 'pdf-report.json').write_text(json.dumps(report, indent=2), encoding='utf8')
print(f'All {pages} actual exported slides match independent native rendering exactly at 144 DPI.')
