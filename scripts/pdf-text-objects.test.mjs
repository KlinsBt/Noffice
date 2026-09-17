import test from 'node:test';
import assert from 'node:assert/strict';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { encodedPdfTextObjects } from './pdf-text-objects.mjs';

async function ordinary(keepSpace) {
  const document = await PDFDocument.create();
  const font = await document.embedFont(StandardFonts.Helvetica);
  const page = document.addPage([300, 200]);
  page.drawText('A', { font, size: 10, x: 72, y: 100 });
  if (keepSpace) page.drawText(' ', { font, size: 10, x: 84, y: 100 });
  page.drawText('B', { font, size: 10, x: 96, y: 100 });
  return document.save();
}

test('encoded spaces remain observable independently of reconstructed word gaps', async () => {
  const complete = (await encodedPdfTextObjects(await ordinary(true)))[0];
  const missing = (await encodedPdfTextObjects(await ordinary(false)))[0];
  assert.equal(complete.map(object => object.text).join(''), 'A B');
  assert.equal(missing.map(object => object.text).join(''), 'AB');
  const geometry = ({ text, x, y, size }) => ({ text, x, y, size });
  assert.deepEqual(complete.filter(object => object.text !== ' ').map(geometry), missing.map(geometry));
});

test('literal delimiters, backslashes and printable Latin-1 decode without text inference', async () => {
  const document = await PDFDocument.create();
  const font = await document.embedFont(StandardFonts.Helvetica);
  const text = '(A) \\ \u00a1\u00e1\u00df\u00ff';
  document.addPage().drawText(text, { font, size: 12 });
  const pages = await encodedPdfTextObjects(await document.save());
  assert.equal(pages[0].map(object => object.text).join(''), text);
});

async function mapped(content, mapping = '<0001> <00660069>\n<0002> <0020>') {
  const document = await PDFDocument.create(), context = document.context;
  const cmap = context.register(context.flateStream(
    '1 begincodespacerange\n<0000> <ffff>\nendcodespacerange\n' +
    '2 beginbfchar\n' + mapping + '\nendbfchar\n'));
  const font = context.register(context.obj({ Type: 'Font', Subtype: 'Type0', Encoding: 'Identity-H', ToUnicode: cmap }));
  const page = document.addPage();
  page.node.set(context.obj('Resources'), context.obj({ Font: { F1: font } }));
  page.node.set(context.obj('Contents'), context.register(context.flateStream(content)));
  return document.save();
}

test('CID text preserves spaces and complete multi-character Unicode clusters', async () => {
  const pages = await encodedPdfTextObjects(await mapped('BT /F1 10 Tf 1 0 0 1 72 100 Tm [<0001> -24 <0002> <0001>] TJ ET'));
  assert.equal(pages[0][0].text, 'fi fi');
});

test('unmapped, duplicate and truncated character codes fail instead of disappearing', async () => {
  const content = value => `BT /F1 10 Tf 1 0 0 1 72 100 Tm <${value}> Tj ET`;
  await assert.rejects(encodedPdfTextObjects(await mapped(content('0003'))), /Unmapped/);
  await assert.rejects(encodedPdfTextObjects(await mapped(content('00'))), /Incomplete/);
  await assert.rejects(encodedPdfTextObjects(await mapped(content('0001'), '<0001> <0061>\n<0001> <0062>')), /Duplicate/);
});

test('unsupported indirect, transformed or malformed text cannot produce evidence', async () => {
  for (const content of ['/X1 Do', '1 0 0 1 0 0 cm', 'BT /F1 10 Tf 1 0 0 1 72 100 Tm <0001> Tj',
    'BT /F1 10 Tf 0 1 -1 0 72 100 Tm <0001> Tj ET'])
    await assert.rejects(encodedPdfTextObjects(await mapped(content)), /Unsupported|Unterminated/);
});
