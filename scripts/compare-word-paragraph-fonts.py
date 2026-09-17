"""Bind actual DOCX font semantics and exact native renderings to independent edits."""
from pathlib import Path
import argparse
import hashlib
import json
import os
import sys

base = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(base / '.local/pdf-tools'))
import pypdfium2 as pdfium

args = argparse.ArgumentParser()
args.add_argument('--negative-control', action='store_true')
options = args.parse_args()
root = base / '.local/word-paragraph-fonts'
hash_file = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
fixture = json.loads((base / 'tests/fixtures/native-word-paragraph-fonts.json').read_text(encoding='utf-8'))
rows = []
for case in fixture['cases']:
    directory = root / case['name']
    receipt = json.loads((directory / 'native-report.json').read_text(encoding='utf-8-sig'))
    assert receipt['scriptHash'] == hash_file(base / 'scripts/verify-word-paragraph-fonts.ps1'), 'Stale native script'
    assert receipt['executableHash'] == fixture['executableHash'], 'Native baseline mismatch'
    for font in receipt['fontHashes']:
        assert hash_file(Path(os.environ['WINDIR']) / 'Fonts' / font['name']) == font['sha256'], 'Native font changed'
    for entry in receipt['hashes']:
        assert hash_file(directory / entry['path']) == entry['sha256'], 'Stale actual file'
    assert receipt['expected'] == case['after'], 'Native source-dependent behavior differs from discovery'
    typing = next((c for c in fixture['typingCases'] if c['name'] == case['name']), None)
    if typing:
        assert receipt['typedExpected'] == typing['after'], 'Native typing differs from discovery'
    for stage in ['edited', 'restored'] + (['typed'] if typing else []):
        if options.negative_control and not rows:
            receipt['actual'][stage]['paragraphs'][0]['mark']['size'] += 1
        semantic = receipt['actual'][stage] == (receipt['typedExpected'] if stage == 'typed' else receipt['expected'])
        with pdfium.PdfDocument(directory / ('native-typed.pdf' if stage == 'typed' else 'native.pdf')) as left, pdfium.PdfDocument(directory / f'{stage}-native.pdf') as right:
            exact = len(left) == len(right)
            for i in range(min(len(left), len(right))):
                a, b = left[i], right[i]
                x, y = a.render(scale=2), b.render(scale=2)
                try:
                    exact &= a.get_size() == b.get_size() and bytes(x.buffer) == bytes(y.buffer)
                finally:
                    x.close(); y.close(); a.close(); b.close()
            rows.append({'case': case['name'], 'stage': stage, 'fontSemantics': semantic,
                         'exactNativePixels': bool(exact), 'pages': len(left),
                         'receiptHash': hash_file(directory / 'native-report.json')})
report = {'passed': all(r['fontSemantics'] and r['exactNativePixels'] for r in rows), 'rows': rows}
name = 'negative-control.json' if options.negative_control else 'comparison.json'
(root / name).write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
print(json.dumps(report))
sys.exit(0 if report['passed'] else 1)
