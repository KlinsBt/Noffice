import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { initializeWordSourceSession } from './docx-initial-session';
import { readWordSettings } from './docx-settings';
import { readWordCompatibility } from './word-compatibility';
import { WORD_NS, descendants, val, wordXml } from './word-xml';

const initial = '0123456789abcdef0123456789abcdef';
const later = 'fedcba9876543210fedcba9876543210';
const source = () => new JSZip()
  .file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="settings" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="customSettings.xml"/></Relationships>')
  .file('word/customSettings.xml', `<w:settings xmlns:w="${WORD_NS}"><w:compat/></w:settings>`)
  .file('customXml/opaque.xml', '<opaque untouched');
const body = () => wordXml(`<w:document xmlns:w="${WORD_NS}"><w:body><w:p><w:r><w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p></w:body></w:document>`);

it('shares the first session across source and initial edits, then persists distinct later sessions', async () => {
  const outputs = [];
  for (let repeat = 0; repeat < 2; repeat++) {
    const zip = source(), document = body();
    zip.file('word/document.xml', new XMLSerializer().serializeToString(document));
    const claim = (await initializeWordSourceSession(zip, document, initial))!;
    expect(claim).toBeTypeOf('function');
    expect(claim()).toBe(claim(initial));
    const second = claim(later);
    expect(second).not.toBe(claim(initial));
    expect(claim(later)).toBe(second);
    expect(descendants(document, 'p').map(p => val(p, 'rsidRDefault'))).toEqual([claim(initial), claim(initial)]);
    expect(document.documentElement.textContent).toBe('AB');
    expect(await zip.file('customXml/opaque.xml')!.async('string')).toBe('<opaque untouched');
    const settings = await zip.file('word/customSettings.xml')!.async('string');
    const parsed = wordXml(settings);
    expect(descendants(parsed, 'rsid').map(node => val(node))).toEqual([claim(initial), second]);
    expect(descendants(parsed, 'compat')).toHaveLength(1);
    outputs.push(settings);
    expect(await initializeWordSourceSession(zip, document, initial)).toBeUndefined();
    expect(await zip.file('word/customSettings.xml')!.async('string')).toBe(settings);
  }
  expect(outputs[0]).toBe(outputs[1]);
});

it.each(['missing-origin', 'existing-register', 'unregistered-id', 'opaque-word-xml', 'nested-body'])('preserves unsupported first-session metadata: %s', async (kind) => {
  const zip = source(), document = body();
  if (kind === 'existing-register') zip.file('word/customSettings.xml', `<w:settings xmlns:w="${WORD_NS}"><w:rsids/></w:settings>`);
  if (kind === 'unregistered-id') zip.file('word/header1.xml', `<w:hdr xmlns:w="${WORD_NS}" w:rsidR="12345678"/>`);
  if (kind === 'opaque-word-xml') zip.file('customXml/opaque.xml', `<broken xmlns:w="${WORD_NS}"`);
  if (kind === 'nested-body') {
    const p = descendants(document, 'p')[0], nested = document.createElementNS(WORD_NS, 'w:sdt');
    p.parentElement!.insertBefore(nested, p); nested.append(p);
  }
  const xml = new XMLSerializer().serializeToString(document), bytes = await zip.generateAsync({ type: 'uint8array' });
  expect(await initializeWordSourceSession(zip, document, kind === 'missing-origin' ? undefined : initial)).toBeUndefined();
  expect(new XMLSerializer().serializeToString(document)).toBe(xml);
  expect(await zip.generateAsync({ type: 'uint8array' })).toEqual(bytes);
});

it('does not assign unrelated initial sessions a shared document root', async () => {
  const first = (await initializeWordSourceSession(source(), body(), initial))!;
  const second = (await initializeWordSourceSession(source(), body(), later))!;
  expect(first()).not.toBe(second());
  expect(first()).toMatch(/^00[\dA-F]{6}$/);
});

const withoutSettings = (orphan: boolean) => {
  const zip = source();
  zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
  zip.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  if (!orphan) zip.remove('word/customSettings.xml');
  return zip;
};

it.each([false, true])('initializes a missing active settings part without activating an orphan: %s', async orphan => {
  const zip = withoutSettings(orphan), document = body();
  const retained = await zip.file('word/customSettings.xml')?.async('string');
  const claim = (await initializeWordSourceSession(zip, document, initial))!;
  expect(claim).toBeTypeOf('function');
  expect(descendants(document, 'p').map(p => val(p, 'rsidRDefault'))).toEqual([claim(), claim(initial)]);
  const next = claim(later);
  expect(next).not.toBe(claim());
  const settings = (await readWordSettings(zip))!;
  expect(settings.path).toBe('word/nofficeSettings.xml');
  expect(descendants(settings.document, 'rsidRoot').map(node => val(node))).toEqual([claim()]);
  expect(descendants(settings.document, 'rsid').map(node => val(node))).toEqual([claim(), next]);
  expect(readWordCompatibility(settings)).toMatchObject({ mode: 12, modeOrigin: 'default' });
  expect(await zip.file('word/customSettings.xml')?.async('string')).toBe(retained);
  expect(await zip.file('customXml/opaque.xml')!.async('string')).toBe('<opaque untouched');
  expect(document.documentElement.textContent).toBe('AB');
});

it.each(['missing', 'invalid'])('rejects %s package metadata without changing body or optional parts', async kind => {
  const zip = withoutSettings(true), document = body();
  if (kind === 'missing') zip.remove('[Content_Types].xml');
  else zip.file('[Content_Types].xml', '<wrong/>');
  const before = await zip.generateAsync({ type: 'uint8array' }), xml = new XMLSerializer().serializeToString(document);
  await expect(initializeWordSourceSession(zip, document, initial)).rejects.toThrow('package metadata');
  expect(new XMLSerializer().serializeToString(document)).toBe(xml);
  expect(await zip.generateAsync({ type: 'uint8array' })).toEqual(before);
});
