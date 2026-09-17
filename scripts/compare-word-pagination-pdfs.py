"""Independent Word pixels and browser print page/text assignments; never equate the two."""
from pathlib import Path
import hashlib,json,sys
base=Path(__file__).resolve().parent.parent
sys.path.insert(0,str(base/'.local/pdf-tools'))
import pypdfium2 as pdfium
widow='--widow' in sys.argv
terminal='--terminal' in sys.argv
orphan='--orphan' in sys.argv
root=base/('.local/word-orphan' if orphan else '.local/word-terminal' if terminal else '.local/word-widow' if widow else '.local/word-pagination')
digest=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
receipt=json.loads((root/'native-report.json').read_text(encoding='utf-8-sig'))
assert receipt['scriptHash']==digest(base/'scripts/verify-word-pagination.ps1')
for b in receipt['hashes']: assert digest(root/b['path'])==b['sha256'], 'Stale '+b['path']
pixels=[]
for first,second in [('source-native.pdf','browser-native.pdf'),('expected-edited.pdf','browser-edited.pdf')]:
 with pdfium.PdfDocument(root/first) as a,pdfium.PdfDocument(root/second) as b:
  assert len(a)==len(b)==receipt['source' if first=='source-native.pdf' else 'edited']['pages']
  for i in range(len(a)):
   p,q=a[i],b[i];x,y=p.render(scale=2),q.render(scale=2)
   try: pixels.append({'case':second,'page':i+1,'exact':p.get_size()==q.get_size() and bytes(x.buffer)==bytes(y.buffer)})
   finally:x.close();y.close();p.close();q.close()
printed=[]
cases=[('source-native.pdf','browser-print.pdf','source')]
if widow or terminal or orphan: cases.append(('expected-edited.pdf','browser-edited-print.pdf','edited'))
for native_name,browser_name,snapshot in cases:
 with pdfium.PdfDocument(root/native_name) as native,pdfium.PdfDocument(root/browser_name) as browser:
  assert len(native)==len(browser)==receipt[snapshot]['pages'], 'Browser print page count differs from Word'
  for i in range(len(native)):
   a,b=native[i],browser[i]
   at,bt=a.get_textpage(),b.get_textpage()
   try:
    expected=at.get_text_range().split();actual=bt.get_text_range().split()
    printed.append({'case':browser_name,'page':i+1,'expectedSize':a.get_size(),'actualSize':b.get_size(),'expectedText':expected,'actualText':actual,
      'passed':expected==actual and all(abs(x-y)<=1 for x,y in zip(a.get_size(),b.get_size()))})
   finally:at.close();bt.close();a.close();b.close()
report={'passed':all(p['exact'] for p in pixels) and all(p['passed'] for p in printed),
 'nativeReceiptHash':digest(root/'native-report.json'),'nativePixels':pixels,'browserPrint':printed}
(root/'pdf-report.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'passed':report['passed'],'nativePages':len(pixels),'browserPages':len(printed),'nativeExact':sum(p['exact'] for p in pixels),'browserPrintPassed':sum(p['passed'] for p in printed)}))
sys.exit(0 if report['passed'] else 1)
