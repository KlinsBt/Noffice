import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFile } from 'node:fs/promises';
import { readWordSettings } from './docx-settings';
import { registerMissingWordSessions } from './docx-missing-session-register';
import { readWordCompatibility } from './word-compatibility';
import { WORD_NS, descendants, val, wordXml } from './word-xml';
const source = async (name = 'orphan-settings') => JSZip.loadAsync(await readFile(`tests/fixtures/word-compatibility-${name}.docx`));
const doc = (attributes: string) => wordXml(`<w:document xmlns:w="${WORD_NS}"><w:body><w:p ${attributes}><w:r><w:t>Edited</w:t></w:r></w:p></w:body></w:document>`);

it.each(['orphan-settings', 'settings-missing'])('registers retained IDs without activating or changing unused settings: %s', async (name) => {
  const zip = await source(name), orphan = await zip.file('word/settings.xml')?.async('string');
  const document = doc('w:rsidR="00000001" w:rsidRDefault="0000abcd"');
  expect(await registerMissingWordSessions(zip, [document])).toBe(true);
  expect(await zip.file('word/settings.xml')?.async('string')).toBe(orphan);
  const settings = (await readWordSettings(zip))!;
  expect(settings.path).toBe('word/nofficeSettings.xml');
  expect(descendants(settings.document, 'rsid').map((node) => val(node))).toEqual(['00000001','0000ABCD']);
  expect(readWordCompatibility(settings)).toMatchObject({ mode: 12, modeOrigin: 'default' });
  const before = await zip.generateAsync({type:'uint8array'});
  expect(await registerMissingWordSessions(zip, [document])).toBe(false);
  expect(await zip.generateAsync({type:'uint8array'})).toEqual(before);
});

it('reserves case-insensitive part names, content-type identities and relationship IDs', async () => {
  const zip = await source();
  zip.file('word/NOFFICESETTINGS.xml','preserved unrelated bytes');
  const types = wordXml(await zip.file('[Content_Types].xml')!.async('string'));
  const node = types.createElementNS(types.documentElement.namespaceURI!, 'Override');
  node.setAttribute('PartName','/word/nofficeSettings1.xml');node.setAttribute('ContentType','application/xml');types.documentElement.append(node);
  zip.file('[Content_Types].xml',new XMLSerializer().serializeToString(types));
  const rels = wordXml(await zip.file('word/_rels/document.xml.rels')!.async('string'));
  const rel = rels.createElementNS(rels.documentElement.namespaceURI!,'Relationship');
  rel.setAttribute('Id','rIdNofficeSettings');rel.setAttribute('Type','urn:opaque');rel.setAttribute('Target','opaque.xml');rels.documentElement.append(rel);
  zip.file('word/_rels/document.xml.rels',new XMLSerializer().serializeToString(rels));
  expect(await registerMissingWordSessions(zip,[doc('w:rsidR="00000001"')])).toBe(true);
  expect((await readWordSettings(zip))!.path).toBe('word/nofficeSettings2.xml');
  expect(await zip.file('word/NOFFICESETTINGS.xml')!.async('string')).toBe('preserved unrelated bytes');
  const result = wordXml(await zip.file('word/_rels/document.xml.rels')!.async('string'));
  expect([...result.documentElement.children].some((node)=>node.getAttribute('Id')==='rIdNofficeSettings1')).toBe(true);
});

it.each(['', 'w:rsidR="invalid"'])('leaves absent or invalid source identities unchanged: %s', async attributes => {
  const zip=await source(),before=await zip.generateAsync({type:'uint8array'});
  expect(await registerMissingWordSessions(zip,[doc(attributes)])).toBe(false);
  expect(await zip.generateAsync({type:'uint8array'})).toEqual(before);
});

it('does not change existing active settings or create a duplicate register', async()=>{
  const zip=await source('mode-14'),before=await zip.generateAsync({type:'uint8array'});
  expect(await registerMissingWordSessions(zip,[doc('w:rsidR="00000001"')])).toBe(false);
  expect(await zip.generateAsync({type:'uint8array'})).toEqual(before);
});

it('rejects broken package metadata before writing any optional session part',async()=>{
  const zip=await source();zip.file('[Content_Types].xml','<wrong/>');const before=await zip.generateAsync({type:'uint8array'});
  await expect(registerMissingWordSessions(zip,[doc('w:rsidR="00000001"')])).rejects.toThrow('package metadata');
  expect(await zip.generateAsync({type:'uint8array'})).toEqual(before);
});
