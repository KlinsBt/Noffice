"""Exact native-rendered split/join export comparison, not browser pagination."""
from pathlib import Path
import hashlib
import json
import sys

base = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium

root = base / '.local/word-ranges'
digest = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
receipt = json.loads((root / 'native-report.json').read_text(encoding='utf-8-sig'))
assert receipt['passed']
assert digest(root / 'browser-source.docx') == receipt['browser']['sourceHash']
pages = []
for case in receipt['liveResults']:
    operation = case['operation']
    for name, field in [(f'native-browser-{operation}.docx', 'nativeHash'), (f'browser-{operation}.docx', 'browserHash'), (f'native-{operation}.pdf', 'nativePdfHash'), (f'browser-{operation}.pdf', 'browserPdfHash')]:
        assert digest(root / name) == case[field], f'Stale artifact: {name}'
    with pdfium.PdfDocument(root / f'native-{operation}.pdf') as left, pdfium.PdfDocument(root / f'browser-{operation}.pdf') as right:
        assert len(left) == len(right), f'Page count differs: {operation}'
        for i in range(len(left)):
            a, b = left[i], right[i]
            first = a.render(scale=2, rev_byteorder=True, force_bitmap_format=pdfium.raw.FPDFBitmap_BGRA)
            second = b.render(scale=2, rev_byteorder=True, force_bitmap_format=pdfium.raw.FPDFBitmap_BGRA)
            try:
                same = a.get_size() == b.get_size() and (first.width, first.height, first.stride) == (second.width, second.height, second.stride) and bytes(first.buffer) == bytes(second.buffer)
                pages.append({'operation': operation, 'page': i + 1, 'samePixels': same, 'nativePixelHash': hashlib.sha256(bytes(first.buffer)).hexdigest(), 'browserPixelHash': hashlib.sha256(bytes(second.buffer)).hexdigest()})
                first.to_pil().save(root / f'native-{operation}-{i+1}.png')
                second.to_pil().save(root / f'browser-{operation}-{i+1}.png')
            finally:
                first.close(); second.close(); a.close(); b.close()
result = {'passed': bool(pages) and all(p['samePixels'] for p in pages), 'dpi': 144, 'tolerance': 0, 'scope': 'Both files rendered by installed Word; browser pagination not certified.', 'nativeReceiptHash': digest(root / 'native-report.json'), 'pages': pages}
(root / 'live-pdf-report.json').write_text(json.dumps(result, indent=2), encoding='utf-8')
print(json.dumps(result, indent=2))
assert result['passed'], 'Native-rendered export pixels differ'
