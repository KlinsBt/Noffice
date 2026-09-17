import { mkdir, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
const revision = 'b230a0df45036b2d5e8b49ebf368b3f90ada8d63';
const files = [
  ['LICENSE', '4b1ad51b2f0efc36f38aa3349a9f30fbd9217547'],
  ['Office 2016/wordcontrols.xlsx', 'e26e752cc7804707d28ef82536de603868993304'],
  ['Office 2016/excelcontrols.xlsx', '2fba513e4a0c38ccb8b994afdd3d0ec132f5af48'],
  ['Office 2016/powerpointcontrols.xlsx', 'dd79ca6aa711e6ea68c7fcb3f28c3b7bde76c58b'],
];
const folder = '.local/ribbon-catalog';
await mkdir(folder, { recursive: true });
const receipts = [];
for (const [path, sha] of files) {
  const url = `https://raw.githubusercontent.com/OfficeDev/office-fluent-ui-command-identifiers/${revision}/${path.split('/').map(encodeURIComponent).join('/')}`;
  const response = await fetch(url);
  if (!response.ok) throw Error(`${response.status} ${path}`);
  const data = Buffer.from(await response.arrayBuffer());
  if (data.length > 2_000_000) throw Error('Catalog exceeds expected download limit.');
  const actual = createHash('sha1').update(`blob ${data.length}\0`).update(data).digest('hex');
  if (actual !== sha) throw Error(`Git blob mismatch: ${path}`);
  await writeFile(`${folder}/${path.split('/').at(-1)}`, data);
  receipts.push({
    path,
    url,
    gitBlob: sha,
    sha256: createHash('sha256').update(data).digest('hex'),
    bytes: data.length,
  });
}
await writeFile(
  `${folder}/receipt.json`,
  JSON.stringify({ revision, files: receipts }, null, 2) + '\n',
);
console.log(`Verified ${files.length} Microsoft command catalog files at ${revision}`);
