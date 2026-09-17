"""Native exported pixels plus browser PDF line contents, columns and page geometry."""
from pathlib import Path
import argparse, ctypes, hashlib, json, sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET
base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
import pypdfium2_raw as raw

def physical_lines(text, columns, no_authored_spaces):
    """PDFium inserts CR/LF inside some Chromium multicolumn words. Discard only
    those flagged generated characters; group all actual glyphs by their origins.
    Authored spaces and every actual character still participate in exact checks.
    """
    result, pending_spaces = [], ''
    for index in range(text.count_chars()):
        char = chr(raw.FPDFText_GetUnicode(text, index))
        if char in '\r\n' and raw.FPDFText_IsGenerated(text, index) == 1:
            continue
        if char == ' ' and raw.FPDFText_IsGenerated(text, index) == 1:
            pending_spaces += char
            continue
        x, y = ctypes.c_double(), ctypes.c_double()
        assert raw.FPDFText_GetCharOrigin(text, index, x, y)
        column = max((i for i, origin in enumerate(columns) if x.value >= origin - .15), default=0)
        column_change = bool(result and column != result[-1]['column'])
        # PDFium inserts a synthetic space across adjacent column origins on a
        # shared baseline. Discard it only when independently captured semantic
        # text contains no authored spaces; actual PDF characters remain exact.
        if pending_spaces and result and not (column_change and no_authored_spaces):
            result[-1]['text'] += pending_spaces
        pending_spaces = ''
        if not result or column_change or abs(result[-1]['y'] - y.value) > .01 or x.value < result[-1]['lastX'] - .01:
            result.append({'text': '', 'x': x.value, 'y': y.value, 'lastX': x.value, 'column': column})
        result[-1]['text'] += char
        result[-1]['lastX'] = x.value
    if pending_spaces and result: result[-1]['text'] += pending_spaces
    return result

def native_columns(paragraphs, page):
    origins = set()
    for paragraph in paragraphs:
        chars = paragraph['characters']
        for i, c in enumerate(chars):
            if c['page'] != page or c['text'] in '\v\f\x0e': continue
            if i == 0 or chars[i-1]['text'] in '\v\f\x0e' or chars[i-1]['page'] != page or abs(c['y']-chars[i-1]['y']) > .01 or c['x'] < chars[i-1]['x'] - .01:
                origins.add(round(c['x'], 2))
    return sorted(origins)

parser = argparse.ArgumentParser()
parser.add_argument('--case', required=True)
parser.add_argument('--family', default='section', choices=['section', 'inline', 'forced', 'ui'])
parser.add_argument('--negative-control', choices=['pixel', 'text', 'space', 'column'])
parser.add_argument('--transition', action='store_true')
args = parser.parse_args()
root = base / f'.local/word-{args.family}-acceptance' / args.case
if args.transition:
    assert args.family == 'section' and args.case.startswith(('oddPage-', 'evenPage-'))
    root = root.with_name('transition-' + args.case)
prefix = 'word-forced-keep' if args.family == 'forced' else f'word-{args.family}-flow'
references = json.loads((base / f'tests/fixtures/native-{prefix}.json').read_text(encoding='utf8'))
assert args.case in references['cases'], 'Unknown flow fixture'
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
read = lambda path: json.loads(path.read_text(encoding='utf-8-sig'))
receipt = read(root / 'native-report.json')
assert receipt['passed'] and receipt['scriptHash'] == sha(base / 'scripts/verify-word-section-flow.ps1')
assert receipt['sourceHash'] == sha(base / f'tests/fixtures/{prefix}-{args.case}.docx')
assert receipt['referenceHash'] == sha(base / f'tests/fixtures/native-{prefix}.json')
for binding in receipt['hashes']:
    assert sha(root / binding['path']) == binding['sha256'], 'Stale ' + binding['path']
def section_types(path):
    namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    with ZipFile(path) as package:
        document = ET.fromstring(package.read('word/document.xml'))
    result = []
    for section in document.iter(f'{{{namespace}}}sectPr'):
        kind = section.find(f'{{{namespace}}}type')
        result.append(kind.get(f'{{{namespace}}}val', 'nextPage') if kind is not None else 'nextPage')
    return result

# These workflows edit body text, not section start semantics. Equal rendered
# pages alone must not hide replacement of a retained nextColumn/continuous type.
types = section_types(base / f'tests/fixtures/{prefix}-{args.case}.docx')
for name in ['browser.docx', 'browser-edited.docx', 'expected-edited.docx']:
    assert section_types(root / name) == types, 'Retained section types differ: ' + name
stages = [('source', 'source-native.pdf', 'browser-native.pdf', 'browser-print.pdf'),
          ('edited', 'expected-edited.pdf', 'browser-edited.pdf', 'browser-edited-print.pdf')]
pixels, printed, control = [], [], False
for stage, expected, actual, print_file in stages:
    with pdfium.PdfDocument(root / expected) as native, pdfium.PdfDocument(root / actual) as export, pdfium.PdfDocument(root / print_file) as browser:
        page_count = max(c['page'] for p in receipt[stage]['paragraphs'] for c in [*p['characters'], p['caret']])
        assert len(native) == len(export) == len(browser) == page_count, f'{stage}: page count differs'
        for index in range(len(native)):
            a, b, c = native[index], export[index], browser[index]
            x, y = a.render(scale=2), b.render(scale=2)
            try:
                left, right = bytes(x.buffer), bytearray(y.buffer)
                if args.negative_control == 'pixel' and not control:
                    right[0] ^= 1
                    control = True
                assert a.get_size() == b.get_size() and left == bytes(right), f'Native pixels differ: {stage}/{index+1}'
                pixels.append({'stage': stage, 'page': index+1, 'exact': True})
            finally:
                x.close(); y.close()
            at, ct = a.get_textpage(), c.get_textpage()
            try:
                paragraphs = receipt[stage]['paragraphs']
                columns = native_columns(paragraphs, index+1)
                no_spaces = not any(c['text'] == ' ' for p in paragraphs for c in p['characters'] if c['page'] == index+1)
                before, after = physical_lines(at, columns, no_spaces), physical_lines(ct, columns, no_spaces)
                if args.negative_control and not control:
                    if args.negative_control == 'text': after[0]['text'] = after[0]['text'][1:]
                    if args.negative_control == 'space': after[0]['text'] = after[0]['text'][:1] + ' ' + after[0]['text'][1:]
                    if args.negative_control == 'column': after[0]['x'] += 1
                    control = True
                assert [line['text'] for line in before] == [line['text'] for line in after], f'Printed line contents differ: {stage}/{index+1}'
                assert all(abs(u['x']-v['x']) <= .15 for u, v in zip(before, after)), f'Printed columns differ: {stage}/{index+1}'
                assert all(abs(u-v) <= 1 for u,v in zip(a.get_size(), c.get_size())), f'Printed page geometry differs: {stage}/{index+1}'
                printed.append({'stage': stage, 'page': index+1, 'passed': True, 'nativeLines': before, 'browserLines': after})
            finally:
                at.close(); ct.close(); a.close(); b.close(); c.close()
report = {'passed': True, 'nativeReceiptHash': sha(root/'native-report.json'), 'scriptHash': sha(Path(__file__)),
          'retainedSectionTypes': types,
          'scope': 'Exact native export pixels; browser exact physical line strings, column starts within .15pt, media boxes within 1pt. Absolute browser glyph baselines remain an open gate.',
          'nativePixels': pixels, 'browserPrint': printed}
(root/'pdf-report.json').write_text(json.dumps(report, indent=2)+'\n', encoding='utf8')
print(json.dumps({'passed': True, 'nativeExactPages': len(pixels), 'browserPrintPages': len(printed)}))
