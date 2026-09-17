"""Compare actual kerning exports and browser print glyphs to native Word."""
from pathlib import Path
import argparse, hashlib, json, sys
base=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(base/'.local/pdf-tools'))
import pypdfium2 as pdfium
from word_pdf_glyphs import glyphs

parser=argparse.ArgumentParser()
parser.add_argument('--negative-control',choices=['pixel','baseline'])
args=parser.parse_args()
root=base/'.local/word-kerning-acceptance'
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p:json.loads(p.read_text(encoding='utf-8-sig'))
receiptPath=root/'native-report.json';receipt=read(receiptPath)
assert receipt['passed'] and receipt['scriptHash']==sha(base/'scripts/verify-word-kerning.ps1')
assert receipt['referenceHash']==sha(base/'tests/fixtures/native-word-kerning.json')
assert receipt['sourceHash']==sha(base/'tests/fixtures/word-kerning.docx')
for binding in receipt['hashes']:assert sha(root/binding['path'])==binding['sha256'],'Stale '+binding['path']
rows=[];control=False
for stage in receipt['stages']:
    name=stage['name']
    with pdfium.PdfDocument(root/f'expected-{name}.pdf') as native,pdfium.PdfDocument(root/f'native-{name}.pdf') as export,pdfium.PdfDocument(root/f'{name}.pdf') as browser:
        assert len(native)==len(export)==len(browser),f'{name}: physical page count differs'
        for i in range(len(native)):
            a,b,c=native[i],export[i],browser[i]
            left,right=a.render(scale=2),b.render(scale=2)
            try:
                actual=bytearray(right.buffer)
                if args.negative_control=='pixel' and not control:actual[0]^=1;control=True
                assert a.get_size()==b.get_size() and bytes(left.buffer)==bytes(actual),f'{name}/{i+1}: native pixels differ'
            finally:left.close();right.close()
            expected,actual=glyphs(a,False),glyphs(c,False)
            assert [g['text'] for g in expected]==[g['text'] for g in actual],f'{name}/{i+1}: printed glyph sequence differs'
            pairs=[(n,v) for n,v in zip(expected,actual) if not (n.get('generated') or v.get('generated'))]
            if args.negative_control=='baseline' and actual and not control:actual[0]['y']+=1;control=True
            errors={key:max((abs(n[key]-v[key]) for n,v in pairs),default=0) for key in ['x','y','size']}
            assert all(e<=.15 for e in errors.values()),f'{name}/{i+1}: printed glyph geometry differs: {errors}'
            assert all(n['color']==v['color'] for n,v in pairs),f'{name}/{i+1}: printed color differs'
            rows.append({'stage':name,'page':i+1,'nativePixelsExact':True,'glyphs':len(pairs),'reconstructedSpaces':len(expected)-len(pairs),'maxErrorPt':errors})
            a.close();b.close();c.close()
assert not args.negative_control,'Negative control did not exercise a glyph'
report={'passed':True,'nativeReceiptHash':sha(receiptPath),'scriptHash':sha(Path(__file__)),
        'glyphHelperHash':sha(base/'scripts/word_pdf_glyphs.py'),'pages':rows,
        'scope':'Actual native exported-file pixels and browser text sequence, painted glyph color/origin/size within .15pt. Reconstructed spaces retain text identity and are measured through adjacent painted glyph positions. Glyph outlines and the full typography matrix remain open.'}
(root/'pdf-report.json').write_text(json.dumps(report,indent=2)+'\n')
print(json.dumps({'passed':True,'pages':len(rows),'glyphs':sum(r['glyphs'] for r in rows),'maxErrorPt':{k:max(r['maxErrorPt'][k] for r in rows) for k in ['x','y','size']}}))
