"""Compare measured native story pages with actual browser PDF/print files."""
import argparse,hashlib,json,sys
from pathlib import Path
base=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(base/'.local/pdf-tools'))
import pypdfium2 as pdfium
from word_pdf_glyphs import glyphs
from word_story_artifacts import check_build
parser=argparse.ArgumentParser()
parser.add_argument('--kind',choices=['download','print'],default='download')
parser.add_argument('--negative-control',choices=['glyph','page','text'])
args=parser.parse_args()
read=lambda p:json.loads(p.read_text(encoding='utf-8-sig'))
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
root=base/'.local/word-side-stories'
reference=read(base/'tests/fixtures/native-word-side-stories.json')
baseline=read(base/'docs/parity/baseline.json')
assert reference['executableHash']==next(a['sha256'] for a in baseline['applications'] if a['executable']=='WINWORD.EXE'),'Pinned native build changed'
assert reference['nativeReceiptHash']==sha(root/'native-discovery.json'),'Native discovery receipt changed'
assert reference['nativeScriptHash']==sha(base/'.local/probe-word-side-stories.ps1'),'Native discovery script changed'
rows=[]
for case_index,(name,sample) in enumerate(reference['cases'].items()):
    expected=root/f'{name}.pdf';actual=root/f'render/{name}/{args.kind}.pdf'
    browser=read(root/f'render/{name}/browser-report.json')
    check_build(browser)
    assert sha(base/'tests/fixtures'/sample['sourceFile'])==browser['sourceSha256']==sample['sourceSha256'],'Source fixture binding changed'
    assert sha(expected)==sample['nativePdfSha256'],'Native reference changed'
    assert sha(actual)==browser['hashes'][f'{args.kind}.pdf'],'Browser PDF binding changed'
    with pdfium.PdfDocument(expected) as a,pdfium.PdfDocument(actual) as b:
        pages=[]
        for index in range(min(len(a),len(b))):
            p,q=a[index],b[index]
            try:
                first,last=glyphs(p,False),glyphs(q,False)
                if case_index==index==0 and args.negative_control=='text':last[0]['text']='!'
                text=[g['text'] for g in first]==[g['text'] for g in last]
                pairs=[(g,h) for g,h in zip(first,last) if not(g.get('generated') or h.get('generated'))]
                if case_index==index==0 and args.negative_control=='glyph':pairs[0][1]['y']+=1
                differences={axis:max((abs(g[axis]-h[axis]) for g,h in pairs),default=0) for axis in ['x','y','size']}
                size=list(q.get_size())
                if case_index==index==0 and args.negative_control=='page':size[0]+=1
                paper=all(abs(u-v)<=.15 for u,v in zip(p.get_size(),size))
                color=all(g['color']==h['color'] for g,h in pairs)
                pages.append({'page':index+1,'glyphs':len(pairs),'text':text,'paper':paper,'color':color,'differences':differences,'passed':text and paper and color and all(v<=.15 for v in differences.values())})
            finally:p.close();q.close()
        rows.append({'name':name,'pages':pages,'passed':len(a)==len(b)==sample['native']['pages'] and all(p['passed'] for p in pages)})
report={'passed':all(r['passed'] for r in rows),'kind':args.kind,'negativeControl':args.negative_control,'rows':rows}
if not args.negative_control:(root/f'story-pages-{args.kind}-comparison.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
print(json.dumps(report,indent=2))
sys.exit(0 if report['passed'] else 1)
