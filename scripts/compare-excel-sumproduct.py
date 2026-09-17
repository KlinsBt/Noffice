"""Check actual browser XLSX caches and independent native export/re-edit PDF pixels."""
from pathlib import Path
import hashlib, json, sys, zipfile, xml.etree.ElementTree as ET
sys.path.insert(0, str(Path('.local/pdf-tools').resolve()))
import pypdfium2 as pdfium

root=Path('.local/excel-sumproduct')
sha=lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
native=read(root/'native-report.json'); browser=read(root/'browser-report.json')
reference=read(Path('tests/fixtures/native-excel-sumproduct.json'))
negative='--negative-control' in sys.argv
assert native['passed'] and browser['passed']
assert native['sourceSha256']==browser['sourceSha256']==reference['sourceSha256']==sha(Path('tests/fixtures/excel-sumproduct.xlsx'))
assert native['scriptSha256']==reference['scriptSha256']==sha(Path('scripts/verify-excel-sumproduct.ps1'))
assert native['browserReceiptSha256']==sha(root/'browser-report.json')
assert native['executableSha256']==reference['executableSha256']
for item in native['artifacts']: assert sha(root/item['path'])==item['sha256'],item['path']

pages=[]; caches=[]
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
for stage in native['stages']:
    name=stage['stage']; output=root/f'{name}.xlsx'
    assert sha(output)==browser['exports'][name]==stage['exportSha256']
    assert stage['cells']==next(s['cells'] for s in reference['stages'] if s['stage']==name)
    with zipfile.ZipFile(output) as z:
        workbook=ET.fromstring(z.read('xl/workbook.xml'))
        sheet=next(s for s in workbook.find('s:sheets',ns) if s.attrib['name']=='Calls')
        rid=sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        relations=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        target=next(r.attrib['Target'] for r in relations if r.attrib['Id']==rid)
        target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
        xml=ET.fromstring(z.read(target))
        cells={c.attrib['r']:c for c in xml.findall('.//s:sheetData/s:row/s:c',ns)}
        for item in stage['cells']:
            cell=cells[item['ref']]; expected=item['value']; raw=cell.findtext('s:v','',ns)
            kind=cell.attrib.get('t','n')
            expected_kind='e' if item['kind']=='error' else 'b' if isinstance(expected,bool) else 'str' if isinstance(expected,str) else 'n'
            actual=float(raw) if kind=='n' else raw=='1' if kind=='b' else raw
            if negative and name=='edit' and item['ref']=='A1': actual+=1
            passed=actual==expected and kind==expected_kind and cell.findtext('s:f','',ns)==item['formula'][1:]
            caches.append({'stage':name,'ref':item['ref'],'expected':expected,'actual':actual,'passed':passed})
    for prefix, a, b in [('export',f'expected-{name}.pdf',f'browser-{name}.pdf'),('reedit',f'reedit-{name}-0.pdf',f'reedit-{name}-1.pdf')]:
        left=pdfium.PdfDocument(root/a); right=pdfium.PdfDocument(root/b)
        assert len(left)==len(right),(name,prefix,'page count')
        for i in range(len(left)):
            x=left[i].render(scale=2).to_pil(); y=right[i].render(scale=2).to_pil()
            pages.append({'stage':name,'kind':prefix,'page':i+1,'exact':x.size==y.size and x.tobytes()==y.tobytes()})
        left.close();right.close()
report={'passed':all(c['passed'] for c in caches) and all(p['exact'] for p in pages),'negativeControl':negative,'scope':'Actual XLSX cached formulas/types/values and exact 144-DPI native-rendered export/re-edit pages; browser pixel/native UI parity remains open','caches':caches,'pages':pages,'nativeReceiptSha256':sha(root/'native-report.json'),'browserReceiptSha256':sha(root/'browser-report.json'),'referenceSha256':sha(Path('tests/fixtures/native-excel-sumproduct.json')),'scriptSha256':sha(Path(__file__))}
(root/('negative-control.json' if negative else 'pdf-report.json')).write_text(json.dumps(report,indent=2)+'\n')
print(json.dumps({'passed':report['passed'],'rawCaches':len(caches),'exactPages':sum(p['exact'] for p in pages),'pages':len(pages),'negativeControl':negative},indent=2))
sys.exit(0 if report['passed'] else 1)
