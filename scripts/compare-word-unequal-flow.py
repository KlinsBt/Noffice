"""Independent unequal-column exported-file, glyph and separator comparison."""
from pathlib import Path
import argparse,ctypes,hashlib,json,sys
base=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(base/'.local/pdf-tools'))
import pypdfium2 as pdfium
import pypdfium2_raw as raw
from word_pdf_glyphs import glyphs

def rules(page):
    result=[]
    for obj in page.get_objects(filter=[raw.FPDF_PAGEOBJ_PATH]):
        left,bottom,right,top=obj.get_bounds()
        if right-left>2 or top-bottom<5:continue
        fill,stroke=ctypes.c_int(),ctypes.c_int();color=[ctypes.c_uint() for _ in range(4)]
        assert raw.FPDFPath_GetDrawMode(obj,fill,stroke)
        assert raw.FPDFPageObj_GetFillColor(obj,*color)
        if not fill.value or [v.value for v in color] != [0,0,0,255]:continue
        result.append({'x':(left+right)/2,'top':page.get_height()-top,
                       'bottom':page.get_height()-bottom,'width':right-left})
    return sorted(result,key=lambda r:(r['x'],r['top']))

parser=argparse.ArgumentParser()
parser.add_argument('--case',required=True)
parser.add_argument('--negative-control',choices=['pixel','glyph','separator','space'])
parser.add_argument('--pdf-kind',choices=['print','download'],default='print')
args=parser.parse_args()
root=base/'.local/word-unequal-acceptance'/args.case
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p:json.loads(p.read_text(encoding='utf-8-sig'))
reference=base/'tests/fixtures/native-word-unequal-flow.json'
referenceCase=read(reference)['cases'][args.case]
wrapped=referenceCase['configuration']['flow']=='wrapped'
receiptPath=root/'native-report.json';receipt=read(receiptPath)
assert receipt['passed'] and receipt['scriptHash']==sha(base/'scripts/verify-word-unequal-flow.ps1')
assert receipt['referenceHash']==sha(reference)
assert receipt['sourceHash']==sha(base/f'tests/fixtures/word-unequal-flow-{args.case}.docx')
for binding in receipt['hashes']:assert sha(root/binding['path'])==binding['sha256'],'Stale '+binding['path']
rows=[];control=False
browserReceipt=read(root/'browser-report.json')
for stage,aName,bName,cName,dName in [
    ('source','source-native.pdf','browser-native.pdf','browser-print.pdf','expected-restored.pdf'),
    ('edited','expected-edited.pdf','browser-edited.pdf','browser-edited-print.pdf','expected-edited.pdf'),
]:
    if args.pdf_kind=='download':
        cName=f'download-{stage}.pdf'
        assert sha(root/cName)==browserReceipt['pdfDownloadHash' if stage=='source' else 'editedPdfDownloadHash'],'Stale PDF download'
    lineStarts=[];paintedByPage={}
    for paragraphIndex,paragraph in enumerate(receipt[stage]['paragraphs']):
        previous=None;lineIndex=0
        for char in paragraph['characters']:
            if char['text'] in '\v\f\x0e':previous=None;continue
            page=char['page'];painted=paintedByPage.setdefault(page,[])
            if previous is None or previous['page']!=page or abs(previous['y']-char['y'])>.01 or char['x']<previous['x']-.01:
                lineStarts.append({'page':page,'glyph':len(painted),'paragraph':paragraphIndex,'line':lineIndex});lineIndex+=1
            painted.append(char['text']);previous=char
    domLines=browserReceipt['lines' if stage=='source' else 'editedLines']
    assert sum(len(p) for p in domLines)==len(lineStarts),'DOM line count differs'
    with pdfium.PdfDocument(root/aName) as expected,pdfium.PdfDocument(root/bName) as exported,pdfium.PdfDocument(root/cName) as browser,pdfium.PdfDocument(root/dName) as saved:
        assert len(expected)==len(exported)==len(browser)==len(saved),'Physical page count differs'
        for i in range(len(expected)):
            a,b,c,d=expected[i],exported[i],browser[i],saved[i]
            try:
                assert a.get_size()==b.get_size()==d.get_size() and all(abs(x-y)<=.15 for x,y in zip(a.get_size(),c.get_size())),'Page size differs'
                nativeBitmap,exportBitmap=d.render(scale=2),b.render(scale=2)
                try:
                    pixels=bytearray(exportBitmap.buffer)
                    if args.negative_control=='pixel' and not control:pixels[0]^=1;control=True
                    assert bytes(nativeBitmap.buffer)==bytes(pixels),f'{stage}/{i+1}: Native export pixels differ'
                finally:nativeBitmap.close();exportBitmap.close()
                noSpaces=not any(ch['text']==' ' for p in receipt[stage]['paragraphs'] for ch in p['characters'] if ch['page']==i+1)
                nativeGlyphs,browserGlyphs=glyphs(a,noSpaces),glyphs(c,noSpaces)
                assert [g['text'] for g in nativeGlyphs]==paintedByPage.get(i+1,[]),'Native semantic/PDF text differs'
                for start in lineStarts:
                    if start['page']!=i+1:continue
                    domLine=domLines[start['paragraph']][start['line']]
                    assert domLine['page']==i+1 and abs(domLine['x']-nativeGlyphs[start['glyph']]['x'])<=.15,'DOM physical column origin differs'
                endSpaces=[]
                if wrapped:
                    # Native Word encodes hanging line-end spaces in its PDF;
                    # Chromium may omit those unpainted trailing glyphs. The
                    # independent DOCX/editor comparison above retains every
                    # semantic space. Align only the ends of independently
                    # identified native lines; internal spaces remain exact.
                    starts=[s['glyph'] for s in lineStarts if s['page']==i+1]+[len(nativeGlyphs)]
                    alignedNative=[];alignedBrowser=[];cursor=0
                    if args.negative_control=='space' and not control:
                        target=next(j for j,g in enumerate(browserGlyphs) if g['text']==' ')
                        browserGlyphs.pop(target);control=True
                    for left,right in zip(starts,starts[1:]):
                        line=nativeGlyphs[left:right];trailing=0
                        while line and line[-1]['text']==' ':line=line[:-1];trailing+=1
                        actual=browserGlyphs[cursor:cursor+len(line)]
                        assert [g['text'] for g in line]==[g['text'] for g in actual],f'{stage}/{i+1}: Line text differs'
                        alignedNative.extend(line);alignedBrowser.extend(actual);cursor+=len(line)
                        present=0
                        while present<trailing and cursor<len(browserGlyphs) and browserGlyphs[cursor]['text']==' ':
                            cursor+=1;present+=1
                        endSpaces.append({'native':trailing,'browser':present})
                    assert cursor==len(browserGlyphs),'Extra browser PDF text'
                    nativeGlyphs,browserGlyphs=alignedNative,alignedBrowser
                assert [g['text'] for g in nativeGlyphs]==[g['text'] for g in browserGlyphs],f'{stage}/{i+1}: Text differs'
                pairs=[(n,v) for n,v in zip(nativeGlyphs,browserGlyphs) if not (n.get('generated') or v.get('generated'))]
                if args.negative_control=='glyph' and pairs and not control:pairs[0][1]['x']+=1;control=True
                errors={key:max((abs(n[key]-v[key]) for n,v in pairs),default=0) for key in ['x','y','size']}
                assert all(e<=.15 for e in errors.values()),f'{stage}/{i+1}: Glyph geometry differs: {errors}'
                assert all(n['color']==v['color'] for n,v in pairs),'Glyph color differs'
                nativeRules,browserRules=rules(a),rules(c)
                if args.negative_control=='separator' and browserRules and not control:browserRules[0]['bottom']+=1;control=True
                assert len(nativeRules)==len(browserRules),'Separator count differs'
                ruleErrors={key:max((abs(n[key]-v[key]) for n,v in zip(nativeRules,browserRules)),default=0) for key in ['x','top','bottom','width']}
                assert all(e<=.15 for e in ruleErrors.values()),f'{stage}/{i+1}: Separator geometry differs: {ruleErrors}'
                rows.append({'stage':stage,'page':i+1,'exactNativePixels':True,'glyphs':len(pairs),
                    'glyphMaxErrorPt':errors,'separators':len(nativeRules),'separatorMaxErrorPt':ruleErrors,'lineEndSpaces':endSpaces})
            finally:a.close();b.close();c.close();d.close()
assert not args.negative_control,'Control did not exercise its target'
report={'passed':True,'pdfKind':args.pdf_kind,'nativeReceiptHash':sha(receiptPath),'scriptHash':sha(Path(__file__)),
        'glyphHelperHash':sha(base/'scripts/word_pdf_glyphs.py'),'pages':rows,
        'scope':'Exact native export pixels; browser line text, painted glyphs and black separator rectangles within .15pt. Chromium may omit unpainted hanging line-end spaces; independent DOCX/editor text remains exact. Wider typography and PDF text-extraction behavior remain open.'}
(root/('pdf-download-report.json' if args.pdf_kind=='download' else 'pdf-report.json')).write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
print(json.dumps({'passed':True,'pages':len(rows),'glyphs':sum(r['glyphs'] for r in rows),'separators':sum(r['separators'] for r in rows)}))
