"""Deliberately author the MIT calibration fixture, independently of browser output."""
from pathlib import Path
import zipfile

base = Path(__file__).resolve().parents[1]
font = lambda size: f'<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="{size * 2}"/>'
cases = [(size, line) for size in [10, 15, 20, 25, 30, 35, 40] for line in [240, 360]]
cases += [(size, line) for line in [480, 1000] for size in [10, 15, 20, 25, 30, 35, 40]]
rows = []
for size, line in cases:
    rule = 'atLeast' if line == 1000 else 'auto'
    rows.append(f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="{rule}"/><w:shd w:fill="00FFFF"/><w:rPr>{font(size)}</w:rPr></w:pPr><w:r><w:rPr>{font(size)}</w:rPr><w:t>probe{size}m{line} Hg</w:t></w:r></w:p>')
target = base / 'tests/fixtures/word-glyph-fonts.docx'
with zipfile.ZipFile(base / 'tests/fixtures/word-mixed-wrap.docx') as source, zipfile.ZipFile(target, 'w', zipfile.ZIP_DEFLATED) as out:
    for name in source.namelist():
        data = source.read(name)
        if name == 'word/document.xml':
            data = ('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + ''.join(rows) + '<w:sectPr><w:pgSz w:w="12000" w:h="30000"/><w:pgMar w:top="900" w:bottom="900" w:left="900" w:right="900"/></w:sectPr></w:body></w:document>').encode()
        out.writestr(name, data)
print(f'Authored {len(rows)} independent calibration cases. Native recapture and review are required: {target}')
