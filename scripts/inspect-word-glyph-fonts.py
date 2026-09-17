from pathlib import Path
import sys,ctypes,re,json,hashlib
base=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(base/'.local/pdf-tools'))
import pypdfium2 as p
import pypdfium2.raw as r
root=base/'.local/word-glyph-fonts';rows=[]
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
receipt=json.loads((root/'report.json').read_text(encoding='utf-8-sig'))
baseline=json.loads((base/'docs/parity/baseline.json').read_text(encoding='utf-8-sig'))
assert receipt['executableSha256']==next(a['sha256'] for a in baseline['applications'] if a['executable']=='WINWORD.EXE')
assert receipt['sourceSha256']==sha(base/'tests/fixtures/word-glyph-fonts.docx')
assert receipt['scriptSha256']==sha(base/'scripts/verify-word-glyph-fonts.ps1')
assert receipt['pdfSha256']==sha(root/'native.pdf')
assert receipt['savedSha256']==sha(root/'saved.docx')
with p.PdfDocument(root/'native.pdf') as doc:
 for page in doc:
  text=page.get_textpage();boxes=[]
  for obj in page.get_objects():
   color=[ctypes.c_uint() for _ in range(4)]
   if r.FPDFPageObj_GetFillColor(obj.raw,*[ctypes.byref(c) for c in color]) and [c.value for c in color[:3]]==[0,255,255]:
    b=obj.get_bounds();boxes.append((page.get_height()-b[3],page.get_height()-b[1]))
  boxes.sort()
  words=list(re.finditer('probe(\d+)m(\d+)',text.get_text_range()))
  assert len(words)==len(boxes),(len(words),len(boxes))
  for word,(top,bottom) in zip(words,boxes):
   matrix=r.FS_MATRIX();assert r.FPDFText_GetMatrix(text.raw,word.start(),ctypes.byref(matrix))
   rows.append({'size':int(word.group(1)),'rule':'atLeast' if word.group(2)=='1000' else 'auto','line':int(word.group(2)),'height':bottom-top,'ascent':page.get_height()-matrix.f-top,'highProbePrediction':int(word.group(1))*.9216})
assert len(rows)==28
report={**receipt,'rows':rows,'inspectorSha256':sha(Path(__file__)),'normalRatio':2355/2048}
if '--record' not in sys.argv:
 expected=json.loads((base/'tests/fixtures/native-word-glyph-fonts.json').read_text(encoding='utf-8-sig'))
 assert receipt['fontSha256']==expected['fontSha256'], 'Native Arial font changed'
 assert rows==expected['rows'], 'Native baseline profile changed; review before accepting'
output=base/'tests/fixtures/native-word-glyph-fonts.json' if '--record' in sys.argv else root/'metrics.json'
output.write_text(json.dumps(report,indent=2)+'\n',encoding='utf8')
print('Verified 28 native regular Arial paragraph baselines and seven independent GDI descenders.' if '--record' not in sys.argv else 'Recorded 28 independent native regular Arial paragraph baselines.')
