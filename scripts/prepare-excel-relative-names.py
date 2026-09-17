"""Sanitize only Noffice's authored relative-name fixture; preserve XML prefixes."""
from pathlib import Path
import re
import zipfile

path = Path(__file__).resolve().parent.parent / '.local/excel-relative-names/source.xlsx'
with zipfile.ZipFile(path) as archive:
    entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
    for info, data in entries:
        if info.filename in ('docProps/core.xml', 'docProps/app.xml'):
            text = data.decode('utf-8')
            text = re.sub(r'(<(?:dc:creator|cp:lastModifiedBy|Company|Manager)>).*?(</[^>]+>)', r'\1Noffice\2', text)
            data = text.encode('utf-8')
        if info.filename == 'xl/workbook.xml':
            data = re.sub(rb'<mc:AlternateContent>.*?</mc:AlternateContent>', b'', data)
        archive.writestr(info, data)
