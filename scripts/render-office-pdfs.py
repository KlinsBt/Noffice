"""Render Office-produced PDFs for review; optional local pypdfium2 test dependency."""
from pathlib import Path
import hashlib
import json
import struct
import sys
import zlib
sys.path.insert(0, str(Path('.local/pdf-tools').resolve()))
import pypdfium2 as pdfium

root = Path('.local/office-corpus/desktop')
def png(bitmap):
    raw = bytes(bitmap.buffer)
    pixels = b''.join(b'\0' + raw[y*bitmap.stride:y*bitmap.stride+bitmap.width*4] for y in range(bitmap.height))
    def chunk(name, data):
        return struct.pack('>I',len(data)) + name + data + struct.pack('>I',zlib.crc32(name+data))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR',struct.pack('>IIBBBBB',bitmap.width,bitmap.height,8,6,0,0,0)) + chunk(b'IDAT',zlib.compress(pixels)) + chunk(b'IEND',b'')

reports = []
for path in sorted(root.glob('*.pdf')):
    with pdfium.PdfDocument(path) as document:
        row = {'file':path.name, 'pages':len(document), 'pagePixelHashes':[], 'inspectedPages':min(len(document),40)}
        for i in range(row['inspectedPages']):
            page = document[i]
            bitmap = page.render(scale=1, rev_byteorder=True, force_bitmap_format=pdfium.raw.FPDFBitmap_BGRA)
            row['pagePixelHashes'].append(hashlib.sha256(bytes(bitmap.buffer)).hexdigest())
            if i == 0: path.with_suffix('.png').write_bytes(png(bitmap))
            bitmap.close()
            page.close()
        reports.append(row)
(root/'render-report.json').write_text(json.dumps(reports,indent=2)+'\n',encoding='utf-8')
print(f'Rendered {len(reports)} PDFs; first-page PNGs and per-page pixel hashes saved. At most 40 pages checked per PDF.')
