"""Bind actual exports, native UI and re-edit evidence, then compare native-rendered pages."""
import argparse
import hashlib
import json
from pathlib import Path
import sys

root = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(root / '.local/pdf-tools'))
import pypdfium2 as pdfium

parser = argparse.ArgumentParser()
parser.add_argument('--negative-control', action='store_true')
negative = parser.parse_args().negative_control
directory = root / '.local/powerpoint-child-bounds'
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
receipt = json.loads((directory / 'native-report.json').read_text(encoding='utf8'))
assert receipt['passed']
assert receipt['scriptSha256'] == sha(root / 'scripts/verify-powerpoint-child-bounds.ps1')
assert receipt['sourceSha256'] == sha(root / 'tests/fixtures/powerpoint-child-bounds.pptx')
assert receipt['browserReceiptSha256'] == sha(directory / 'browser-report.json')
assert receipt['uiReceiptSha256'] == sha(root / '.local/powerpoint-child-ui/ui-edit-report.json')
pages = 0
for stage in receipt['stages']:
    name = stage['stage']
    for index, origin in enumerate(['native', 'browser']):
        hashes = stage['hashes'][index]
        pptx = f'{name}-native.pptx' if origin == 'native' else f'{name}.pptx'
        for key, file in [('pptxSha256', pptx), ('pdfSha256', f'{name}-{origin}.pdf'),
                          ('reeditSha256', f'{name}-{origin}-reedit.pptx'),
                          ('reeditPdfSha256', f'{name}-{origin}-reedit.pdf')]:
            assert hashes[key] == sha(directory / file), f'Stale artifact: {file}'
    for suffix in ['', '-reedit']:
        with pdfium.PdfDocument(directory / f'{name}-native{suffix}.pdf') as expected, \
                pdfium.PdfDocument(directory / f'{name}-browser{suffix}.pdf') as actual:
            assert len(expected) == len(actual) == 4
            for index in range(4):
                p, q = expected[index], actual[index]
                left, right = p.render(scale=2), q.render(scale=2)
                try:
                    actual_pixels = bytearray(right.buffer)
                    if negative and pages == 0:
                        actual_pixels[0] ^= 1
                    assert p.get_size() == q.get_size() and bytes(left.buffer) == actual_pixels, \
                        f'Native rendered slide differs: {name}{suffix} slide {index + 1}'
                    pages += 1
                finally:
                    left.close(); right.close(); p.close(); q.close()
report = {'passed': True, 'pages': pages, 'dpi': 144,
          'scope': 'Exact native-rendered browser exports and native re-edits; browser glyph parity is not certified',
          'nativeReceiptSha256': sha(directory / 'native-report.json'),
          'scriptSha256': sha(Path(__file__))}
(directory / 'pdf-report.json').write_text(json.dumps(report, indent=2), encoding='utf8')
print(f'All {pages} native-rendered export/re-edit pages match exactly at 144 DPI.')
