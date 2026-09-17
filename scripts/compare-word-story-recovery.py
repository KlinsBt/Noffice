"""Verify actual recovery exports against the independently authored option state."""
import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '.local/pdf-tools'))
import pypdfium2 as pdfium
from word_story_artifacts import check_build

root = ROOT / '.local/word-story-options'
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
read = lambda path: json.loads(path.read_text(encoding='utf-8-sig'))
receipt = read(root / 'recovery-native-report.json')
assert receipt['scriptHash'] == sha(ROOT / 'scripts/verify-word-story-recovery.ps1')
assert receipt['optionReceiptHash'] == sha(root / 'native-export-report.json')
options = read(root / 'native-export-report.json')
assert options['scriptHash'] == sha(ROOT / 'scripts/verify-word-story-options.ps1')
baseline = read(ROOT / 'docs/parity/baseline.json')
assert receipt['executableHash'] == options['executableHash'] == next(a['sha256'] for a in baseline['applications'] if a['executable'] == 'WINWORD.EXE')
for item in receipt['hashes'] + options['hashes']:
    assert sha(root / item['path']) == item['sha256'], f"Changed artifact {item['path']}"
expected = next(row for row in options['rows'] if row['name'] == 'first-one-off' and row['stage'] == 'edited')
assert {row['name'] for row in receipt['rows']} == {'late-abort', 'other-tab'}
results = []
for row in receipt['rows']:
    folder = root / row['folder']
    browser = read(folder / 'browser-report.json')
    check_build(browser)
    assert not browser['errors']
    assert browser['exportHash'] == sha(folder / 'unsaved.docx')
    assert browser['sourceHash'] == sha(ROOT / 'tests/fixtures/word-section-stories/first-even.docx')
    pages = []
    with pdfium.PdfDocument(root / expected['folder'] / 'expected-edited.pdf') as a, pdfium.PdfDocument(folder / 'native-unsaved.pdf') as b:
        assert len(a) == len(b) == expected['expected']['pages'] == row['actual']['pages']
        for index in range(len(a)):
            p, q = a[index], b[index]
            x, y = p.render(scale=2), q.render(scale=2)
            try:
                pages.append(p.get_size() == q.get_size() and x.to_pil().tobytes() == y.to_pil().tobytes())
            finally:
                x.close(); y.close(); p.close(); q.close()
    results.append({'name': row['name'], 'properties': row['actual'] == expected['expected'], 'pages': pages,
                    'passed': row['actual'] == expected['expected'] and bool(pages) and all(pages)})
report = {'passed': all(row['passed'] for row in results), 'rows': results,
          'nativeReceiptHash': sha(root / 'recovery-native-report.json')}
(root / 'recovery-comparison.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
print(json.dumps(report, indent=2))
sys.exit(0 if report['passed'] else 1)
