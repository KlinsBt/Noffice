"""Compare Word-rendered exports exactly; does not certify browser pagination."""
from pathlib import Path
import hashlib
import json
import sys

base = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium

root = base / '.local/word-sections'
def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()
def read(name):
    return json.loads((root / name).read_text(encoding='utf-8-sig'))

native = read('native-report.json')
assert not native['mismatches'], 'Native geometry comparison has failures'
for entry in native['hashes']:
    assert digest(root / entry['path']) == entry['sha256'].lower(), 'Stale native geometry receipt'
pdfs = []
for name, docx, receipt in [('native-expected.pdf', 'native-expected.docx', 'native-pdf.json'), ('browser.pdf', 'browser.docx', 'browser-pdf.json')]:
    binding = read(receipt)
    assert digest(root / docx) == binding['docxHash'].lower(), 'Stale PDF source'
    assert digest(root / name) == binding['pdfHash'].lower(), 'Stale PDF output'
    pdfs.append({'path': name, 'sha256': digest(root / name)})

pages = []
with pdfium.PdfDocument(root / 'native-expected.pdf') as expected, pdfium.PdfDocument(root / 'browser.pdf') as actual:
    assert len(expected) == len(actual) == native['actual']['pages'], 'Page counts differ'
    for i in range(len(expected)):
        left, right = expected[i], actual[i]
        a = left.render(scale=2, rev_byteorder=True, force_bitmap_format=pdfium.raw.FPDFBitmap_BGRA)
        b = right.render(scale=2, rev_byteorder=True, force_bitmap_format=pdfium.raw.FPDFBitmap_BGRA)
        try:
            first, second = bytes(a.buffer), bytes(b.buffer)
            sizes_match = (a.width, a.height, a.stride) == (b.width, b.height, b.stride)
            different_bytes = sum(x != y for x, y in zip(first, second)) + abs(len(first) - len(second))
            pages.append({'page': i + 1, 'points': left.get_size(), 'width': a.width, 'height': a.height, 'sameSize': sizes_match and left.get_size() == right.get_size(), 'differentBytes': different_bytes, 'expectedPixelHash': hashlib.sha256(first).hexdigest(), 'actualPixelHash': hashlib.sha256(second).hexdigest()})
            a.to_pil().save(root / f'native-page-{i + 1}.png')
            b.to_pil().save(root / f'browser-export-page-{i + 1}.png')
        finally:
            a.close(); b.close(); left.close(); right.close()

passed = all(p['sameSize'] and p['differentBytes'] == 0 for p in pages)
report = {'passed': passed, 'scope': 'Both DOCX files rendered by native Word; not browser pagination', 'dpi': 144, 'tolerance': 0, 'nativeReceiptHash': digest(root / 'native-report.json'), 'pdfs': pdfs, 'pages': pages}
(root / 'pdf-report.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps(report, indent=2))
sys.exit(0 if passed else 1)
