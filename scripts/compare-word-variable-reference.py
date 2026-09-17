"""Pin native layout to its independent rendered reference, separate from COM caret diagnostics."""
from pathlib import Path
import argparse,hashlib,json,sys
base=Path(__file__).resolve().parents[1];sys.path[:0]=[str(base/'.local/pdf-tools'),str(base/'scripts')]
import pypdfium2 as pdfium
from word_pdf_glyphs import glyphs
parser=argparse.ArgumentParser();parser.add_argument('--case',choices=['mixed','single','keep'],required=True);args=parser.parse_args()
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
referencePath=base/'tests/fixtures/native-word-variable-pagination.json';reference=json.loads(referencePath.read_text(encoding='utf-8'))['cases'][args.case]
root=base/f'.local/word-variable-{args.case}'
expected=base/('.local/word-variable-pagination/native.pdf' if args.case=='mixed' else f'.local/word-variable-{args.case}/native.pdf')
actual=root/'source-native.pdf';assert sha(expected)==reference['nativeReceipt']['pdfSha256']
pages=[]
with pdfium.PdfDocument(expected) as a,pdfium.PdfDocument(actual) as b:
 assert len(a)==len(b)==reference['snapshot']['pages']
 for i in range(len(a)):
  p,q=a[i],b[i];x,y=p.render(scale=2),q.render(scale=2)
  try:pages.append(dict(page=i+1,pixels=bytes(x.buffer)==bytes(y.buffer),glyphs=glyphs(p,False)==glyphs(q,False),paper=p.get_size()==q.get_size()))
  finally:x.close();y.close();p.close();q.close()
passed=all(all(row[key] for key in ['pixels','glyphs','paper']) for row in pages)
report=dict(passed=passed,pages=pages,expected=expected.relative_to(base).as_posix(),expectedHash=sha(expected),actualHash=sha(actual),referenceHash=sha(referencePath),scriptHash=sha(Path(__file__)),glyphHelperHash=sha(base/'scripts/word_pdf_glyphs.py'))
(root/'native-source-reference.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print('Pinned native source render',args.case,passed,len(pages),'exact pages');sys.exit(not passed)
