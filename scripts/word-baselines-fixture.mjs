import fs from 'node:fs/promises';
import JSZip from 'jszip';

const zip = await JSZip.loadAsync(await fs.readFile('tests/fixtures/word-mixed-lines.docx'));
const part = await zip.file('word/document.xml').async('string');
const font = (size) => `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="${size * 2}"/>`;
const run = (text, size) =>
  `<w:r><w:rPr>${font(size)}</w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
const cases = [];
for (let repeat = 0; repeat < 3; repeat++)
  for (const size of [10, 20, 30])
    cases.push({ size, mark: 10, rule: 'auto', line: 240, mixed: false });
for (const size of [10, 20, 30])
  cases.push({ size, mark: 10, rule: 'auto', line: 360, mixed: false });
for (const size of [10, 20, 30])
  cases.push({ size, mark: 40, rule: 'auto', line: 360, mixed: true });
for (const size of [10, 20, 30])
  cases.push({ size, mark: 10, rule: 'atLeast', line: 1000, mixed: true });
const paragraphs = cases
  .map(
    (c) =>
      `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${c.line}" w:lineRule="${c.rule}"/><w:widowControl w:val="0"/><w:rPr>${font(c.mark)}</w:rPr></w:pPr>${run('Sg ', c.mixed ? 8 : c.size)}${run('Large', c.size)}</w:p>`,
  )
  .join('');
zip.file(
  'word/document.xml',
  part.replace(
    /(<w:body>)[\s\S]*(<w:sectPr[ >][\s\S]*?<\/w:sectPr>)(<\/w:body>)/,
    `$1${paragraphs}$2$3`,
  ),
);
await fs.mkdir('.local/word-baselines', { recursive: true });
await fs.writeFile(
  '.local/word-baselines/authored.docx',
  await zip.generateAsync({ type: 'nodebuffer' }),
);
await fs.writeFile('.local/word-baselines/cases.json', JSON.stringify(cases, null, 2));
