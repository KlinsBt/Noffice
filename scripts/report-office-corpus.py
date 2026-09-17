"""Independent ZIP/XML inventory and loss signals for downloaded Office fixtures.

Does not launch Office, evaluate formulas, render documents, or execute document content.
"""
import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

ROOT = Path('.local/office-corpus')
NS = {'word': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'excel': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'drawing': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'powerpoint': 'http://schemas.openxmlformats.org/presentationml/2006/main'}

def inventory(path, kind):
    with zipfile.ZipFile(path) as archive:
        entries = [e for e in archive.infolist() if not e.is_dir()]
        if sum(e.file_size for e in entries) > 100 * 1024 * 1024:
            raise ValueError('Archive inventory exceeds 100 MB')
        payloads = {e.filename: archive.read(e) for e in entries}
    metrics = Counter()
    texts = []
    xml_issues = []
    documents = {}
    for name, raw in payloads.items():
        if not name.endswith('.xml'):
            continue
        if b'<!DOCTYPE' in raw.upper() or b'<!ENTITY' in raw.upper():
            raise ValueError('DTD/entity declarations are not accepted by the inventory')
        doc = ET.fromstring(raw)
        documents[name] = doc
        for e in doc.iter():
            tag = e.tag.rsplit('}', 1)[-1]
            namespace = e.tag.split('}')[0].lstrip('{')
            root_namespace = doc.tag.split('}')[0].lstrip('{')
            spreadsheet_namespaces = (NS['excel'], 'http://purl.oclc.org/ooxml/spreadsheetml/main')
            if namespace in spreadsheet_namespaces and root_namespace in spreadsheet_namespaces and namespace != root_namespace:
                xml_issues.append(name + ': mixed Strict/Transitional SpreadsheetML')
            if kind == 'word' and namespace == NS['word']:
                labels = {'tbl':'tables', 'drawing':'drawings', 'pict':'legacyDrawings', 'footnoteReference':'footnoteReferences', 'endnoteReference':'endnoteReferences', 'commentReference':'commentReferences', 'ins':'trackedInsertions', 'del':'trackedDeletions', 'fldChar':'fieldMarkers', 'sdt':'contentControls', 'bookmarkStart':'bookmarks', 'sectPr':'sections', 'bidi':'bidiProperties'}
            elif kind == 'excel' and namespace in (NS['excel'], 'http://purl.oclc.org/ooxml/spreadsheetml/main'):
                labels = {'sheet':'sheets', 'mergeCell':'merges', 'f':'formulas', 'dataValidation':'validationRules', 'conditionalFormatting':'conditionalFormattingRanges', 'sheetProtection':'protectedSheets', 'definedName':'definedNames', 'pane':'panes', 'hyperlink':'hyperlinks'}
                if tag == 'sheet' and e.get('state', 'visible') != 'visible': metrics['hiddenSheets'] += 1
            elif kind == 'powerpoint' and namespace == NS['powerpoint']:
                labels = {'sldId':'slides', 'sp':'shapes', 'pic':'pictures', 'grpSp':'groups', 'graphicFrame':'graphicFrames', 'timing':'animationTimelines', 'transition':'transitions'}
            else:
                labels = {}
            if tag in labels: metrics[labels[tag]] += 1
            if kind == 'powerpoint' and namespace == NS['drawing'] and tag == 'tbl': metrics['tables'] += 1
        # Inspect user document/slide text, excluding style definitions and placeholder masters.
        relevant = kind == 'word' and (name == 'word/document.xml' or re.match(r'word/(header|footer|footnotes|endnotes|comments)', name))
        relevant = relevant or (kind == 'powerpoint' and re.match(r'ppt/(slides/slide\d+|notesSlides/notesSlide\d+)\.xml$', name))
        if relevant:
            textns = NS['word'] if kind == 'word' else NS['drawing']
            for paragraph in doc.iter('{' + textns + '}p'):
                texts.append(''.join(e.text or '' for e in paragraph.iter('{' + textns + '}t')))
    patterns = {'mediaParts': r'^(word|xl|ppt)/media/', 'chartParts': r'^(word|xl|ppt)/charts/chart\d+\.xml$', 'diagramParts': r'^(word|xl|ppt)/diagrams/.*\.xml$', 'embeddingParts': r'^(word|xl|ppt)/embeddings/', 'headerParts': r'^word/header\d+\.xml$', 'footerParts': r'^word/footer\d+\.xml$', 'commentParts': r'(^word/comments|^xl/comments|^ppt/comments)', 'pivotParts': r'^xl/pivot', 'tableParts': r'^xl/tables/table', 'masterParts': r'^ppt/slideMasters/slideMaster\d+\.xml$', 'layoutParts': r'^ppt/slideLayouts/slideLayout\d+\.xml$'}
    for key, pattern in patterns.items(): metrics[key] = sum(bool(re.search(pattern, name)) for name in payloads)
    return {'bytes': path.stat().st_size, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest(), 'parts': {name: hashlib.sha256(raw).hexdigest() for name, raw in payloads.items()}, 'features': dict(metrics), 'xmlIssues':sorted(set(xml_issues)), 'textTokens': dict(Counter(re.findall(r'\w+|[^\w\s]', '\n'.join(texts))))}

def compare(before, after):
    loss = {key: {'before': count, 'after': after['features'].get(key, 0)} for key, count in before['features'].items() if count > after['features'].get(key, 0)}
    token_loss = sum((Counter(before['textTokens']) - Counter(after['textTokens'])).values())
    missing = sorted(set(before['parts']) - set(after['parts']))
    return {'byteIdentical': before['sha256'] == after['sha256'], 'missingParts': missing, 'featureDecreases': loss, 'missingTextTokenOccurrences': token_loss, 'xmlIssues':after['xmlIssues'], 'changedRetainedParts': [name for name in before['parts'].keys() & after['parts'].keys() if before['parts'][name] != after['parts'][name]], 'reviewRequired': bool(loss or token_loss or missing or after['xmlIssues'])}

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--strict', action='store_true', help='Fail on incomplete workflows, missing edits, or measured preservation loss signals.')
    args = parser.parse_args()
    manifest = json.loads(Path('corpus/manifest.json').read_text(encoding='utf-8'))
    rows = []
    for item in manifest['files']:
        directory = ROOT / 'results' / item['id']
        row = {'id':item['id'], 'kind':item['kind'], 'filename':item['filename'], 'focus':item['focus']}
        try:
            source = inventory(ROOT / 'input' / item['kind'] / item['filename'], item['kind'])
            row['source'] = {k:v for k,v in source.items() if k != 'textTokens'}
            workflow = directory / 'workflow.json'
            row['workflow'] = json.loads(workflow.read_text(encoding='utf-8')) if workflow.exists() else {'status':'not-run'}
            for stage in ['unchanged', 'edited']:
                output = directory / (stage + Path(item['filename']).suffix)
                if output.exists(): row[stage] = compare(source, inventory(output, item['kind']))
        except Exception as error:
            row['error'] = str(error)
        rows.append(row)
    totals = {'downloaded':sum('source' in r for r in rows), 'workflowPassed':sum(r.get('workflow',{}).get('status') == 'workflow-passed' for r in rows), 'safelyRejected':sum(r.get('workflow',{}).get('status') == 'safely-rejected' for r in rows), 'unchangedByteIdentical':sum(r.get('unchanged',{}).get('byteIdentical',False) for r in rows), 'editedExports':sum('edited' in r for r in rows), 'editedReviewRequired':sum(r.get('edited',{}).get('reviewRequired',False) for r in rows)}
    desktop = []
    for path in sorted((ROOT / 'desktop').glob('report-*.json')):
        records = json.loads(path.read_text(encoding='utf-8-sig'))
        desktop.extend(records if isinstance(records, list) else [records])
    failures = []
    for row in rows:
        reasons = []
        if row.get('error'): reasons.append(row['error'])
        if row.get('workflow',{}).get('status') != 'workflow-passed': reasons.append('No successful import/edit workflow: ' + row.get('workflow',{}).get('status', 'not-run'))
        if not row.get('unchanged',{}).get('byteIdentical'): reasons.append('No byte-identical unchanged export')
        if 'edited' not in row: reasons.append(row.get('workflow',{}).get('editUnavailable', 'No edited export'))
        elif row['edited']['reviewRequired']: reasons.append('Edited export has preservation loss signals')
        if reasons: failures.append({'id':row['id'], 'reasons':reasons})
    report = {'revision':manifest['revision'], 'totals':totals, 'gateFailures':failures, 'desktopValidation':desktop or 'not-run', 'scope':'Package inventories and text/feature loss signals; separate native Office evidence is attached when available. No complete formula-semantic or desktop parity certification.', 'files':rows}
    ROOT.mkdir(parents=True, exist_ok=True)
    (ROOT / 'report.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf-8')
    lines = ['# Downloaded Office corpus results', '', report['scope'], '', f"Totals: {json.dumps(totals)}", '', '| File | Focus | Workflow | Unchanged identical | Edited review signals |', '| --- | --- | --- | --- | --- |']
    for row in rows:
        edited = row.get('edited',{})
        signals = ', '.join(edited.get('featureDecreases',{}))
        if edited.get('missingTextTokenOccurrences'): signals += f"; {edited['missingTextTokenOccurrences']} missing text tokens"
        if edited.get('missingParts'): signals += f"; {len(edited['missingParts'])} missing parts"
        if edited.get('xmlIssues'): signals += '; ' + ', '.join(edited['xmlIssues'])
        lines.append(f"| {row['filename']} | {row['focus']} | {row.get('workflow',{}).get('status', 'error')} | {row.get('unchanged',{}).get('byteIdentical',False)} | {signals or ('No counted decreases' if edited else 'Not tested')} |")
    lines += ['', '## Outstanding strict-gate requirements', '']
    lines += [f"- `{failure['id']}`: {'; '.join(failure['reasons'])}" for failure in failures] or ['None in the measured package gate.']
    lines += ['', 'Zero counted decreases does not prove fidelity. Package names may legitimately change during reconstruction; decreases are review signals, not an automatic visual verdict. Protected initial selections are recorded as unedited. Windows Office screen/print/playback validation remains open.', '']
    (ROOT / 'REPORT.md').write_text('\n'.join(lines),encoding='utf-8')
    print(json.dumps(totals))
    print(f"Reports: {ROOT / 'REPORT.md'} and report.json")
    if args.strict and failures:
        for failure in failures: print(f"{failure['id']}: {'; '.join(failure['reasons'])}")
        sys.exit(1)

if __name__ == '__main__': main()
