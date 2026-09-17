"""Publish the owned Word-saved test baseline, removing personal core metadata.
Keep XML namespace declarations intact: xsi:type values can reference prefixes.
"""
from pathlib import Path
import hashlib, re, zipfile, sys
base = Path(__file__).resolve().parent.parent
folder = 'word-font-cascade' if '--cascade' in sys.argv else 'word-line-metrics'
source = base / '.local' / folder / 'word-saved-source.docx'
target = base / 'tests/fixtures' / (folder + '.docx')
target.parent.mkdir(exist_ok=True)
with zipfile.ZipFile(source) as original, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as output:
    for part in original.infolist():
        data = original.read(part.filename)
        if part.filename == 'docProps/core.xml':
            for tag in [b'dc:creator', b'cp:lastModifiedBy']:
                data = re.sub(b'(<'+tag+b'>)[^<]*(</'+tag+b'>)', rb'\g<1>Noffice\g<2>', data)
        output.writestr(part, data)
print(hashlib.sha256(target.read_bytes()).hexdigest())
