"""Baseline diagnostics keep PDF text origins and native COM line origins separate."""
from pathlib import Path
import ctypes, json, hashlib, sys
base=Path(__file__).resolve().parent.parent
sys.path.insert(0,str(base/'.local/pdf-tools'))
import pypdfium2 as pdf
import pypdfium2.raw as raw
root=base/'.local/word-baselines'
read=lambda p:json.loads(p.read_text(encoding='utf-8-sig'))
receipt=root/('native-report.json' if (root/'native-report.json').exists() else 'source-report.json')
native=read(receipt);cases=read(root/'cases.json')
assert native['sourceSha256']==hashlib.sha256((root/'source.docx').read_bytes()).hexdigest()
for artifact in native['artifacts']:
 assert artifact['sha256']==hashlib.sha256((root/artifact['path']).read_bytes()).hexdigest(),artifact['path']
rows=[]
with pdf.PdfDocument(root/'source-native.pdf') as doc:
 for pageNo in range(len(doc)):
  page=doc[pageNo];text=page.get_textpage()
  for i in range(text.count_chars()):
   if text.get_text_range(i,1)!='S':continue
   matrix=raw.FS_MATRIX();assert raw.FPDFText_GetMatrix(text.raw,i,ctypes.byref(matrix))
   n=len(rows);p=native['source'][n];assert p['chars'][0]['page']==pageNo+1
   top=p['chars'][0]['y'];baseline=page.get_height()-matrix.f
   rows.append({'case':cases[n],'topPt':top,'pdfBaselinePt':baseline,'offsetPt':baseline-top})
assert len(rows)==len(cases)
(root/'baseline-report.json').write_text(json.dumps({'nativeReceiptSha256':hashlib.sha256(receipt.read_bytes()).hexdigest(),'rows':rows,'status':'diagnostic, exact browser baselines remain open'},indent=2)+'\n',encoding='utf-8')
for r in rows:print(r)
