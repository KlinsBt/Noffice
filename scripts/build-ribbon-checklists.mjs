import ExcelJS from 'exceljs';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
const root = '.local/ribbon-catalog',
  output = process.env.NOFFICE_RIBBON_OUTPUT || 'docs/ribbon';
const receipt = JSON.parse(await readFile(`${root}/receipt.json`, 'utf8'));
if (receipt.revision !== 'b230a0df45036b2d5e8b49ebf368b3f90ada8d63')
  throw Error('Unexpected catalog revision');
await mkdir(output, { recursive: true });
const titles = {
  TabHome: 'Home / Start',
  TabInsert: 'Insert / Einfügen',
  TabWordDesign: 'Design / Entwurf',
  TabPageLayoutWord: 'Layout',
  TabReferences: 'References / Verweise',
  TabMailings: 'Mailings / Sendungen',
  TabReviewWord: 'Review / Überprüfen',
  TabView: 'View / Ansicht',
};
const totals = [];
for (const application of ['word', 'excel', 'powerpoint']) {
  const data = await readFile(`${root}/${application}controls.xlsx`);
  const source = receipt.files.find((file) => file.path.endsWith(`/${application}controls.xlsx`));
  if (createHash('sha256').update(data).digest('hex') !== source.sha256)
    throw Error(`Source changed: ${application}`);
  const book = new ExcelJS.Workbook();
  await book.xlsx.load(data);
  const sheet = book.worksheets[0],
    tabs = new Map();
  let count = 0;
  const filename = `${output}/${application.toUpperCase()}_COMMANDS.md`;
  let prior = '';
  try {
    prior = await readFile(filename, 'utf8');
  } catch {}
  const checked = new Set(
    [...prior.matchAll(/^- \[x\].*<!-- source-row:(\d+) -->/gm)].map((m) => Number(m[1])),
  );
  sheet.eachRow((row, index) => {
    if (index === 1) return;
    const id = row.getCell(1).text;
    if (!id) return;
    const type = row.getCell(2).text,
      tab = row.getCell(4).text || 'Unassigned / tab containers';
    if (!tabs.has(tab)) tabs.set(tab, []);
    const group = row.getCell(5).text,
      parents = [6, 7, 8].map((i) => row.getCell(i).text).filter(Boolean);
    tabs
      .get(tab)
      .push(
        `- [${checked.has(index) ? 'x' : ' '}] \`${id}\` — ${type}${group ? `; group \`${group}\`` : ''}${parents.length ? `; menu ${parents.map((p) => '`' + p + '`').join(' → ')}` : ''}. <!-- source-row:${index} -->`,
      );
    count++;
  });
  const lines = [
    `# ${application[0].toUpperCase() + application.slice(1)} command acceptance checklist`,
    '',
    `Source: [Microsoft Office 2016 control catalog](${source.url}), revision \`${receipt.revision}\`; SHA-256 \`${source.sha256}\`. Microsoft catalog data is covered by [its MIT license](LICENSE.microsoft.txt).`,
    '',
    `Contains **${count} catalog placements** across **${tabs.size} tab/context categories**. Repeated controls occur in multiple menus. Containers, galleries and context-menu entries are included; a catalog placement is not necessarily a separate feature. This fixes the acceptance baseline to Office 2016; newer Office features require additional inventory.`,
    '',
    'Check a row only after its complete stated command behavior and applicable save/export, undo and error paths pass. Add evidence to TASKS.md or a linked scope report. Existing partial implementations do not automatically certify an entire Microsoft command, gallery or dialog. Unticked rows include partial implementations. This file is an acceptance inventory, not an app menu.',
    '',
  ];
  for (const [tab, rows] of tabs)
    lines.push(`## ${titles[tab] || tab}`, '', `Catalog tab: \`${tab}\`.`, '', ...rows, '');
  const text = lines.join('\n');
  if (process.argv.includes('--check')) {
    if (text !== prior) throw Error(`Checklist drift: ${filename}`);
  } else await writeFile(filename, text);
  totals.push({ application, placements: count, categories: tabs.size });
}
if (!process.argv.includes('--check')) {
  await writeFile(`${output}/source.json`, JSON.stringify(receipt, null, 2) + '\n');
  await writeFile(`${output}/LICENSE.microsoft.txt`, await readFile(`${root}/LICENSE`));
}
console.log(JSON.stringify(totals));
