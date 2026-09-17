"""Bind actual browser files, compare native export pixels and mixed browser print pages."""
from pathlib import Path
import hashlib
import json
import sys
base = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium
root = base / '.local/word-mixed'
digest = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
native = json.loads((root / 'native-report.json').read_text(encoding='utf-8-sig'))
assert native['passed']
assert native['scriptHash'] == digest(base / 'scripts/verify-word-mixed-sections.ps1')
for binding in native['hashes']:
    assert digest(root / binding['path']) == binding['sha256'], f"Stale receipt: {binding['path']}"
comparisons = []
for operation, count in [('split', 2), ('join', 1)]:
    with pdfium.PdfDocument(root / f'native-{operation}.pdf') as a, pdfium.PdfDocument(root / f'browser-native-{operation}.pdf') as b:
        assert len(a) == len(b) == count
        for i in range(count):
            left, right = a[i], b[i]
            x, y = left.render(scale=2), right.render(scale=2)
            try:
                exact = left.get_size() == right.get_size() and bytes(x.buffer) == bytes(y.buffer)
                comparisons.append({'operation': operation, 'page': i+1, 'exactPixels': exact})
                x.to_pil().save(root / f'native-{operation}-{i+1}.png')
                y.to_pil().save(root / f'browser-native-{operation}-{i+1}.png')
            finally:
                x.close(); y.close(); left.close(); right.close()
printed = []
expected = native['observed'][0]['expected']
with pdfium.PdfDocument(root / 'split-print.pdf') as pdf:
    assert len(pdf) == 2, 'Mixed browser print must contain two pages'
    for i in range(2):
        p = pdf[i]
        actual = p.get_size()
        wanted = [expected[i]['width']/20, expected[i]['height']/20]
        text = p.get_textpage().get_text_range().replace('\r', '').replace('\n', '')
        wanted_text = expected[i]['text'].replace('\r', '').replace('\f', '')
        # Chromium quantizes physical page rectangles; do not claim exact glyph placement.
        passed = all(abs(x-y) <= 1 for x,y in zip(actual,wanted)) and text == wanted_text
        printed.append({'page': i+1, 'actualPoints': actual, 'expectedPoints': wanted,
                        'tolerancePoints': 1, 'text': text, 'expectedText': wanted_text, 'passed': passed})
        p.close()
report = {'nativeReceiptHash': digest(root / 'native-report.json'), 'nativePixels': comparisons,
          'browserPrint': printed, 'passed': all(c['exactPixels'] for c in comparisons) and all(p['passed'] for p in printed)}
(root / 'pdf-report.json').write_text(json.dumps(report, indent=2)+'\n', encoding='utf-8')
print(json.dumps(report, indent=2))
sys.exit(0 if report['passed'] else 1)
