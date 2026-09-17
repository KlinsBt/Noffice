import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { wordEditSession } from './docx-edit-session';
import { WORD_NS, wordXml, descendants, val } from './docx-import';

const settings = (id: string) =>
  `<w:settings xmlns:w="${WORD_NS}"><w:rsids><w:rsidRoot w:val="00000001"/><w:rsid w:val="${id}"/></w:rsids></w:settings>`;
const settingsPackage = () => new JSZip().file('word/_rels/document.xml.rels',
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>');
it('allocates one monotonic save ID per export and changes settings only when claimed', async () => {
  const zip = settingsPackage();
  zip.file('word/settings.xml', settings('00000010'));
  zip.file('word/header1.xml', `<w:hdr xmlns:w="${WORD_NS}"><w:p w:rsidR="0000FFFF"/></w:hdr>`);
  const before = await zip.file('word/settings.xml')!.async('string');
  const claim = await wordEditSession(zip);
  expect(await zip.file('word/settings.xml')!.async('string')).toBe(before);
  const id = claim()!;
  expect(id).toMatch(/^[\dA-F]{8}$/);
  expect(Number.parseInt(id, 16)).toBeGreaterThan(0xffff);
  expect(claim()).toBe(id);
  expect(
    descendants(wordXml(await zip.file('word/settings.xml')!.async('string')), 'rsid').map((n) =>
      val(n),
    ),
  ).toEqual(['00000010', id]);
});
it.each(['FFFFFFFF', 'invalid'])(
  'preserves unsupported/exhausted session metadata %s',
  async (id) => {
    const zip = settingsPackage();
    zip.file('word/settings.xml', settings(id));
    const claim = await wordEditSession(zip);
    expect(claim()).toBeUndefined();
    expect(await zip.file('word/settings.xml')!.async('string')).toBe(settings(id));
  },
);
it('does not create missing native metadata parts', async () => {
  const zip = new JSZip();
  expect((await wordEditSession(zip))()).toBeUndefined();
  expect(Object.keys(zip.files)).toEqual([]);
});

it('keeps reopened editing sessions distinct without prematurely exhausting IDs', async () => {
  const zip = settingsPackage();
  zip.file('word/settings.xml', settings('FFFFFFFD'));
  const claim = await wordEditSession(zip);
  const first = claim('first'),
    second = claim('second');
  expect(new Set([first, second])).toEqual(new Set(['FFFFFFFE', 'FFFFFFFF']));
  expect(claim('first')).toBe(first);
  expect(claim('third')).toBeUndefined();
});

it('preserves opaque non-Word XML without trying to interpret its content', async () => {
  const zip = settingsPackage();
  zip.file('word/settings.xml', settings('00000010'));
  zip.file('customXml/opaque.xml', '<uninterpreted');
  const claim = await wordEditSession(zip);
  expect(claim()).toMatch(/^[\dA-F]{8}$/);
  expect(await zip.file('customXml/opaque.xml')!.async('string')).toBe('<uninterpreted');
});

it('reproduces persisted typing-session IDs across reload exports of the same original', async () => {
  const original = settingsPackage();
  original.file('word/settings.xml', settings('00000010'));
  const bytes = await original.generateAsync({ type: 'uint8array' });
  const exports = [];
  for (let i = 0; i < 2; i++) {
    const zip = await JSZip.loadAsync(bytes);
    const claim = await wordEditSession(zip);
    const first = claim('0123456789abcdef0123456789abcdef');
    const second = claim('fedcba9876543210fedcba9876543210');
    expect(first).not.toBe(second);
    exports.push(await zip.file('word/settings.xml')!.async('string'));
  }
  expect(exports[0]).toBe(exports[1]);
});
