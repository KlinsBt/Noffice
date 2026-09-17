"""Extract native soft lines using glyph baselines, not mixed-size COM tops."""
from pathlib import Path
import ctypes
import hashlib
import json
import re
import sys

base = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
import pypdfium2.raw as raw

sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
read = lambda p: json.loads(p.read_text(encoding='utf-8-sig'))
root = base / '.local/word-mixed-wrap'
receipt_name = 'native-report.json' if '--compare' in sys.argv else 'reference-report.json'
reference = read(root / receipt_name)
probe = read(root / 'boxes/native-report.json')
assert reference['sourceSha256'] == probe['sourceSha256'] == sha(base / 'tests/fixtures/word-mixed-wrap.docx')
assert reference['scriptSha256'] == sha(base / 'scripts/verify-word-mixed-wrap.ps1')
assert probe['scriptSha256'] == sha(base / 'scripts/probe-word-paragraph-boxes.ps1')
assert reference['executableSha256'] == probe['executableSha256']
baseline = read(base / 'docs/parity/baseline.json')
assert reference['executableSha256'] == next(a['sha256'] for a in baseline['applications'] if a['executable'] == 'WINWORD.EXE')
for artifact in reference['artifacts']:
    if artifact['path'].startswith(('source', 'expected-')):
        assert artifact['sha256'] == sha(root / artifact['path']), 'Stale ' + artifact['path']

def matrices(text):
    result = []
    for i in range(text.count_chars()):
        matrix = raw.FS_MATRIX()
        if raw.FPDFText_GetMatrix(text.raw, i, ctypes.byref(matrix)):
            result.append((text.get_text_range(i, 1), *[getattr(matrix, k) for k in ('a','b','c','d','f')],text.get_charbox(i)[1],text.get_charbox(i)[3]))
    return result

stages = {}
for stage in probe['stages']:
    name = stage['stage']
    for suffix, key in [('.docx','docxSha256'),('.pdf','pdfSha256'),('-boxes.pdf','boxesPdfSha256')]:
        assert stage[key] == sha(root / 'boxes' / (name + suffix))
    path = root / ('source-native.pdf' if name == 'source' else f'expected-{name}.pdf')
    with pdfium.PdfDocument(path) as original, pdfium.PdfDocument(root / 'boxes' / f'{name}-boxes.pdf') as marked:
        assert len(original) == len(marked) == 1
        p, q = original[0], marked[0]
        text, marked_text = p.get_textpage(), q.get_textpage()
        assert matrices(text) == matrices(marked_text), 'Shading changed vertical glyph layout'
        boxes = {2:[], 3:[]}
        for obj in q.get_objects():
            color = [ctypes.c_uint() for _ in range(4)]
            if raw.FPDFPageObj_GetFillColor(obj.raw, *[ctypes.byref(c) for c in color]):
                rgb = [c.value for c in color[:3]]
                index = 2 if rgb == [0,255,255] else 3 if rgb == [255,0,255] else None
                if index:
                    bounds = obj.get_bounds()
                    boxes[index].append({'top':q.get_height()-bounds[3], 'bottom':q.get_height()-bounds[1]})
        paragraphs = []
        for index, prefix in [(2,'alpha'),(3,'beta')]:
            lines = []
            for match in re.finditer(prefix+r'\d+', text.get_text_range()):
                matrix = raw.FS_MATRIX()
                assert raw.FPDFText_GetMatrix(text.raw, match.start(), ctypes.byref(matrix))
                baseline = p.get_height()-matrix.f
                if not lines or abs(lines[-1]['baseline']-baseline)>0.15:
                    lines.append({'baseline':baseline,'words':[]})
                lines[-1]['words'].append(match.group())
            assert sum(len(l['words']) for l in lines)==20
            top = min(b['top'] for b in boxes[index])
            height = max(b['bottom'] for b in boxes[index])-top
            ordered_boxes = sorted(boxes[index], key=lambda b:b['top'])
            assert len(ordered_boxes) == len(lines)
            for line, box in zip(lines, ordered_boxes):
                line['baseline'] -= top
                line['boxTop'] = box['top']-top
                line['boxHeight'] = box['bottom']-box['top']
            paragraphs.append({'paragraph':index,'height':height,'lines':lines,'backgroundFragments':len(boxes[index])})
        stages[name] = paragraphs

report = {'sourceSha256':reference['sourceSha256'],'executableSha256':reference['executableSha256'],'fonts':reference['fonts'],'stages':stages,'scope':'Independent native glyph line identities and paragraph-background heights; shaded/unshaded vertical glyph coordinates identical; horizontal painter quantization differs and is excluded','scriptSha256':sha(Path(__file__)),'nativeReceiptPath':receipt_name,'nativeReceiptSha256':sha(root/receipt_name),'probeSha256':sha(root/'boxes/native-report.json')}
out = base/'tests/fixtures/native-word-mixed-wrap.json' if '--record' in sys.argv else root/'native-lines.json'
out.write_text(json.dumps(report,indent=2)+'\n')
print(json.dumps({name:[{'height':p['height'],'lines':len(p['lines'])} for p in ps] for name,ps in stages.items()},indent=2))
