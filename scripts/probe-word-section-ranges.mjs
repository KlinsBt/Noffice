import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { Document, Packer, Paragraph } from 'docx';
import JSZip from 'jszip';

const root = '.local/word-ranges';
await fs.mkdir(root, { recursive: true });
const zip = await JSZip.loadAsync(
  await Packer.toBuffer(new Document({ sections: [{ children: [new Paragraph('Fixture')] }] })),
);
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
zip.file(
  'word/document.xml',
  `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W}"><w:body>
<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:pPr><w:r><w:t>AlphaBeta</w:t></w:r></w:p>
<w:p><w:r><w:t>Gamma</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`,
);
const buffer = await zip.generateAsync({ type: 'nodebuffer' });
await fs.writeFile(`${root}/source.docx`, buffer);
await fs.writeFile(
  `${root}/source.json`,
  JSON.stringify(
    {
      sourceHash: createHash('sha256').update(buffer).digest('hex'),
      cases: [
        'Insert a paragraph mark at character 5 within the section-ending paragraph.',
        'Delete the first section break character.',
      ],
    },
    null,
    2,
  ),
);
console.log('Authored section-boundary fixture in .local/word-ranges.');
