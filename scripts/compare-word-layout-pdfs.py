"""Native export pixels and browser print page sizes, with separate expectations."""
from pathlib import Path
import hashlib
import json
import sys
base=Path(__file__).resolve().parent.parent
sys.path.insert(0,str(base/'.local/pdf-tools'))
import pypdfium2 as pdfium
root=base/'.local/word-layout'
native=json.loads((root/'native-report.json').read_text(encoding='utf-8-sig'))
assert not native['mismatches'], 'Native geometry comparison failed'
for entry in native['hashes']:
    assert hashlib.sha256((root/entry['path']).read_bytes()).hexdigest()==entry['sha256'], 'Stale native receipt'
pages=[]
with pdfium.PdfDocument(root/'native-preset.pdf') as a, pdfium.PdfDocument(root/'browser-native.pdf') as b:
    assert len(a)==len(b)==1, 'Unexpected native page count'
    for i in range(len(a)):
        left,right=a[i],b[i]
        first,second=left.render(scale=2),right.render(scale=2)
        try:
            same=left.get_size()==right.get_size() and bytes(first.buffer)==bytes(second.buffer)
            pages.append({'page':i+1,'exactPixels':same,'dimensions':left.get_size()})
            first.to_pil().save(root/'native-preset.png')
            second.to_pil().save(root/'browser-native.png')
        finally:
            first.close();second.close();left.close();right.close()
print_sizes=[]
for name,geometry in [('custom-print.pdf',native['nativeCustom']),('browser-print.pdf',native['nativePreset'])]:
    with pdfium.PdfDocument(root/name) as pdf:
        assert len(pdf)==1, 'Unexpected browser print page count'
        page=pdf[0]
        actual=page.get_size();expected=[geometry['width']/20,geometry['height']/20]
        # Chromium PDF media boxes quantize CSS physical sizes; this checks the page
        # rectangle only, not typography or native browser pagination equivalence.
        passed=all(abs(x-y)<=1 for x,y in zip(actual,expected))
        print_sizes.append({'path':name,'actualPoints':actual,'expectedPoints':expected,'tolerancePoints':1,'passed':passed})
        page.close()
report={'nativePixels':pages,'browserPrintRectangles':print_sizes,'passed':all(p['exactPixels'] for p in pages) and all(p['passed'] for p in print_sizes),'nativeReceiptHash':hashlib.sha256((root/'native-report.json').read_bytes()).hexdigest()}
(root/'pdf-report.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps(report,indent=2))
sys.exit(0 if report['passed'] else 1)
