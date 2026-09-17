"""Check actual browser XLSX caches and independent native export/re-edit PDF pixels."""
from pathlib import Path
import hashlib, json, sys, zipfile, xml.etree.ElementTree as ET
sys.path.insert(0, str(Path('.local/pdf-tools').resolve()))
import pypdfium2 as pdfium

root=Path('.local/excel-arrays')
sha=lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
read=lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
native=read(root/'native-report.json'); browser=read(root/'browser-report.json')
reference=read(Path('tests/fixtures/native-excel-array-contexts.json'))
negative='--negative-control' in sys.argv
assert native['passed'] and browser['passed']
assert native['sourceSha256']==browser['sourceSha256']==reference['sourceSha256']==sha(Path('tests/fixtures/excel-array-contexts.xlsx'))
assert native['scriptSha256']==reference['editingReferenceScriptSha256']==sha(Path('scripts/verify-excel-array-contexts.ps1'))
assert native['browserReceiptSha256']==sha(root/'browser-report.json')
assert native['executableSha256']==reference['executableSha256']
for item in native['artifacts']: assert sha(root/item['path'])==item['sha256'],item['path']

pages=[]; caches=[]
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
for stage in native['stages']:
    name=stage['stage']; output=root/f'{name}.xlsx'
    assert sha(output)==browser['exports'][name]==stage['exportSha256']
    assert stage['cells']==next(s['cells'] for s in reference['editingStages'] if s['stage']==name)
    with zipfile.ZipFile(output) as z:
        workbook=ET.fromstring(z.read('xl/workbook.xml'))
        sheet=next(s for s in workbook.find('s:sheets',ns) if s.attrib['name']=='Calls')
        rid=sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        relations=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        target=next(r.attrib['Target'] for r in relations if r.attrib['Id']==rid)
        target=target.lstrip('/') if target.startswith('/') else 'xl/'+target
        xml=ET.fromstring(z.read(target))
        cells={c.attrib['r']:c for c in xml.findall('.//s:sheetData/s:row/s:c',ns)}
        with zipfile.ZipFile(root/f'expected-{name}.xlsx') as expected_zip:
            expected_xml=ET.fromstring(expected_zip.read(target))
            expected_cells={c.attrib['r']:c for c in expected_xml.findall('.//s:sheetData/s:row/s:c',ns)}
        for item in stage['cells']:
            cell=cells[item['ref']]; expected=item['value']; raw=cell.findtext('s:v','',ns)
            kind=cell.attrib.get('t','n')
            expected_kind='e' if item['kind']=='error' else 'b' if isinstance(expected,bool) else 'str' if isinstance(expected,str) else 'n'
            actual=float(raw) if kind=='n' else raw=='1' if kind=='b' else raw
            if negative and name=='edit' and item['ref']=='C1': actual+=1
            # Native COM hides saved compatibility prefixes such as _xlfn.IFNA.
            # Compare actual serialized formulas with independent native XML,
            # without deleting prefixes or weakening formula equality.
            expected_formula=expected_cells[item['ref']].findtext('s:f','',ns)
            actual_formula=cell.findtext('s:f','',ns)
            passed=actual==expected and kind==expected_kind and actual_formula==expected_formula
            formula=cell.find('s:f',ns)
            passed=passed and ((formula.attrib.get('t')=='array')==item['hasArray'])
            if item['hasArray']: passed=passed and formula.attrib.get('ref')==item['arrayAddress'].replace('$','')
            caches.append({'stage':name,'ref':item['ref'],'expected':expected,'actual':actual,'formula':actual_formula,'nativeFormula':expected_formula,'arrayRange':formula.attrib.get('ref'),'passed':passed})
    for prefix, a, b in [('export',f'expected-{name}.pdf',f'browser-{name}.pdf'),('reedit',f'reedit-{name}-0.pdf',f'reedit-{name}-1.pdf')]:
        left=pdfium.PdfDocument(root/a); right=pdfium.PdfDocument(root/b)
        assert len(left)==len(right),(name,prefix,'page count')
        for i in range(len(left)):
            x=left[i].render(scale=2).to_pil(); y=right[i].render(scale=2).to_pil()
            pages.append({'stage':name,'kind':prefix,'page':i+1,'exact':x.size==y.size and x.tobytes()==y.tobytes()})
        left.close();right.close()
report={'passed':all(c['passed'] for c in caches) and all(p['exact'] for p in pages),'negativeControl':negative,'scope':'Actual XLSX CSE identities and cached formulas/types/values, plus exact 144-DPI native-rendered export/re-edit pages; browser pixel/native UI parity remains open','caches':caches,'pages':pages,'nativeReceiptSha256':sha(root/'native-report.json'),'browserReceiptSha256':sha(root/'browser-report.json'),'referenceSha256':sha(Path('tests/fixtures/native-excel-array-contexts.json')),'scriptSha256':sha(Path(__file__))}
(root/('negative-control.json' if negative else 'pdf-report.json')).write_text(json.dumps(report,indent=2)+'\n')
print(json.dumps({'passed':report['passed'],'rawCaches':len(caches),'exactPages':sum(p['exact'] for p in pages),'pages':len(pages),'negativeControl':negative},indent=2))
sys.exit(0 if report['passed'] else 1)
