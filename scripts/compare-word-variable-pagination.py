"""Require exact native export pixels and independent browser print page/text assignments."""
from pathlib import Path
import argparse,hashlib,json,sys
base=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(base/'.local/pdf-tools'))
import pypdfium2 as pdfium
parser=argparse.ArgumentParser();parser.add_argument('--case',choices=['mixed','single','keep'],default='mixed');parser.add_argument('--negative-control',action='store_true');args=parser.parse_args()
root=base/f'.local/word-variable-{args.case}'
sha=lambda path:hashlib.sha256(path.read_bytes()).hexdigest()
read=lambda path:json.loads(path.read_text(encoding='utf-8-sig'))
receipt=read(root/'native-report.json')
assert receipt['passed'] and receipt['scriptHash']==sha(base/'scripts/verify-word-variable-pagination.ps1')
assert receipt['sourceHash']==sha(base/f'tests/fixtures/word-variable-{args.case}.docx')
assert receipt['referenceHash']==sha(base/'tests/fixtures/native-word-variable-pagination.json')
for binding in receipt['hashes']:assert sha(root/binding['path'])==binding['sha256'],'Stale '+binding['path']
control=read(root/'native-source-reference.json')
assert control['passed'] and control['scriptHash']==sha(base/'scripts/compare-word-variable-reference.py')
assert control['glyphHelperHash']==sha(base/'scripts/word_pdf_glyphs.py')
assert control['referenceHash']==receipt['referenceHash'] and control['actualHash']==sha(root/'source-native.pdf')
reference=read(base/'tests/fixtures/native-word-variable-pagination.json')['cases'][args.case]
assert control['expectedHash']==reference['nativeReceipt']['pdfSha256']==sha(base/control['expected'])
assert all(all(page[key] for key in ['pixels','glyphs','paper']) for page in control['pages'])
stages=[('source','source-native.pdf','browser-native.pdf','browser-print.pdf'),('edited','expected-edited.pdf','browser-edited.pdf','browser-edited-print.pdf')]
if receipt.get('keep'):stages.append(('keep','expected-keep.pdf','browser-keep.pdf','browser-keep-print.pdf'))
pixels=[];printed=[];control=False
for stage,expected,actual,print_file in stages:
 with pdfium.PdfDocument(root/expected) as native,pdfium.PdfDocument(root/actual) as export,pdfium.PdfDocument(root/print_file) as browser:
  assert len(native)==len(export)==len(browser)==receipt[stage]['pages'],f'{stage}: page count differs'
  for index in range(len(native)):
   a,b,c=native[index],export[index],browser[index]
   x,y=a.render(scale=2),b.render(scale=2)
   try:
    left,right=bytes(x.buffer),bytearray(y.buffer)
    if args.negative_control and not control:right[0]^=1;control=True
    exact=a.get_size()==b.get_size() and left==bytes(right)
    assert exact,f'Native exported pixels differ: {stage}, page {index+1}'
    pixels.append({'stage':stage,'page':index+1,'exact':True})
   finally:x.close();y.close()
   at,ct=a.get_textpage(),c.get_textpage()
   try:
    before,after=at.get_text_range().split(),ct.get_text_range().split()
    matched=before==after and all(abs(u-v)<=1 for u,v in zip(a.get_size(),c.get_size()))
    printed.append({'stage':stage,'page':index+1,'passed':matched,'nativeText':before,'browserText':after,'nativeSize':a.get_size(),'browserSize':c.get_size()})
   finally:at.close();ct.close();a.close();b.close();c.close()
report={'passed':all(p['passed'] for p in printed),'nativeReceiptHash':sha(root/'native-report.json'),'scriptHash':sha(Path(__file__)),'nativePixels':pixels,'browserPrint':printed,'nativeSourceControlHash':sha(root/'native-source-reference.json'),'historicalCoordinatesMatch':receipt['historicalCoordinatesMatch'],'referenceCoordinateDiagnostics':receipt['referenceCoordinateDiagnostics'],'nativeUiVerified':False}
(root/'pdf-report.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf8')
print(json.dumps({'passed':report['passed'],'nativeExactPages':len(pixels),'browserPrintPages':len(printed),'browserPrintPassed':sum(p['passed'] for p in printed)}))
sys.exit(0 if report['passed'] else 1)
