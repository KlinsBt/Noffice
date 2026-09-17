"""Read-only OOXML comparison; never evaluates formulas or external links."""
import argparse
import json
import posixpath
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree as ET

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def inspect(path):
    with zipfile.ZipFile(path) as z:
        def xml(name):
            return ET.fromstring(z.read(name)) if name in z.namelist() else None
        book = xml('xl/workbook.xml')
        rels = {e.get('Id'): posixpath.normpath(posixpath.join('xl', e.get('Target'))).lstrip('/')
                if not e.get('Target', '').startswith('/') else e.get('Target').lstrip('/')
                for e in xml('xl/_rels/workbook.xml.rels')}
        strings = xml('xl/sharedStrings.xml')
        strings = [''.join(si.itertext()) for si in strings] if strings is not None else []
        styles = xml('xl/styles.xml')
        groups = {e.tag.split('}')[-1]: list(e) for e in styles} if styles is not None else {}
        def record(e):
            return {'attributes': e.attrib, 'xml': ET.tostring(e, encoding='unicode')}
        sheets = []
        for entry in book.findall('s:sheets/s:sheet', NS):
            part = rels[entry.get('{'+NS['r']+'}id')]
            ws = xml(part)
            cells = {}
            for c in ws.findall('s:sheetData/s:row/s:c', NS):
                v = c.find('s:v', NS)
                value = v.text if v is not None else None
                kind = c.get('t', 'n')
                if kind == 's' and value is not None: value = strings[int(value)]
                if kind == 'inlineStr': value = ''.join(c.find('s:is', NS).itertext())
                formula = c.find('s:f', NS)
                cells[c.get('r')] = {'type': kind, 'value': value,
                    'formula': formula.text if formula is not None else None,
                    'formulaAttributes': formula.attrib if formula is not None else {},
                    'style': int(c.get('s', 0))}
            sheet = {'name': entry.get('name'), 'part': part, 'state': entry.get('state', 'visible'),
                     'dimension': ws.find('s:dimension', NS).get('ref') if ws.find('s:dimension', NS) is not None else None,
                     'cells': cells, 'cellCount': len(cells),
                     'valuedCells': sum(c['value'] is not None for c in cells.values()),
                     'formulaCount': sum(bool(c['formula'] is not None or c['formulaAttributes']) for c in cells.values()),
                     'styleUsage': dict(Counter(str(c['style']) for c in cells.values())),
                     'merges': [e.get('ref') for e in ws.findall('s:mergeCells/s:mergeCell', NS)],
                     'columns': [e.attrib for e in ws.findall('s:cols/s:col', NS)],
                     'rows': [e.attrib for e in ws.findall('s:sheetData/s:row', NS)],
                     'features': {name: [record(e) for e in ws.findall('s:'+name, NS)] for name in
                         ['sheetViews','sheetFormatPr','sheetProtection','conditionalFormatting','dataValidations','autoFilter','pageMargins','pageSetup','printOptions','headerFooter','drawing','legacyDrawing','rowBreaks','colBreaks','tableParts','hyperlinks']}}
            sheets.append(sheet)
        return {'path': str(path), 'bytes': Path(path).stat().st_size, 'parts': z.namelist(),
                'definedNames': [record(e) for e in book.findall('s:definedNames/s:definedName', NS)],
                'styles': {name: [record(e) for e in elements] for name,elements in groups.items()},
                'sheets': sheets}

def summarize(book):
    return {key: book[key] for key in ['bytes']} | {
        'parts': len(book['parts']), 'definedNames': len(book['definedNames']),
        'styles': {k: len(v) for k,v in book['styles'].items()},
        'sheets': [{k: s[k] for k in ['name','dimension','cellCount','valuedCells','formulaCount','state']} | {
            'merges': len(s['merges']), 'columns': s['columns'],
            'customHeightRows': sum('ht' in r for r in s['rows']),
            'hiddenRows': sum(r.get('hidden') == '1' for r in s['rows']),
            'features': {k:len(v) for k,v in s['features'].items() if v},
        } for s in book['sheets']]}

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('original'); parser.add_argument('exported'); parser.add_argument('--report', required=True)
    args = parser.parse_args()
    original, exported = inspect(args.original), inspect(args.exported)
    changes = []
    for old in original['sheets']:
        new = next((s for s in exported['sheets'] if s['name'] == old['name']), None)
        if not new: changes.append({'sheet':old['name'], 'missing':True}); continue
        diffs = []
        for ref in set(old['cells']) | set(new['cells']):
            a,b = old['cells'].get(ref,{}), new['cells'].get(ref,{})
            if (a.get('value'),a.get('formula')) != (b.get('value'),b.get('formula')):
                diffs.append({'ref':ref,'original':a,'exported':b})
        changes.append({'sheet':old['name'], 'valueOrFormulaDifferences':diffs,
            'lostMerges': sorted(set(old['merges'])-set(new['merges']))})
    report = {'original': original, 'exported': exported, 'changes':changes}
    dest=Path(args.report); dest.parent.mkdir(parents=True,exist_ok=True)
    dest.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'original':summarize(original),'exported':summarize(exported),
        'missingParts': sorted(set(original['parts'])-set(exported['parts'])),
        'changes':[{k:(len(v) if isinstance(v,list) else v) for k,v in c.items()} for c in changes]},ensure_ascii=False,indent=2))
