from pathlib import Path
from zipfile import ZipFile,ZIP_DEFLATED
import re
src=Path('.local/word-font-sizes/source.docx')
with ZipFile(src) as z: parts={i.filename:z.read(i) for i in z.infolist()}
core=parts['docProps/core.xml'].decode('utf-8')
for tag in ['dc:creator','cp:lastModifiedBy']:
    core=re.sub(r'(<'+tag+r'>).*?(</'+tag+r'>)',r'\1Noffice test fixture\2',core)
parts['docProps/core.xml']=core.encode('utf-8')
with ZipFile('tests/fixtures/word-font-sizes.docx','w',ZIP_DEFLATED) as z:
    for name,data in parts.items(): z.writestr(name,data)
