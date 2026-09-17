import { describe, expect, it } from 'vitest';
import { assertXmlComplexity, contentFingerprint } from './office-preservation';
import { newFile, type WordContent } from './model';
import { nativeBackup, unchangedOfficeOriginal } from './formats';
import JSZip from 'jszip';

describe('original Office content identity and recursive import bounds', () => {
  it('returns untouched package bytes without layout but never discards text or page edits', async () => {
    const file = newFile('word');
    if (file.content.kind !== 'word') throw Error();
    const originalContent = structuredClone(file.content);
    const bytes = await new JSZip().file('word/document.xml', '<document/>').generateAsync({ type: 'arraybuffer' });
    file.original = { name: 'source.docx', data: bytes, contentFingerprint: await contentFingerprint(file.content) };
    const original = await unchangedOfficeOriginal(file);
    expect(original).toBeInstanceOf(Blob);
    const returned = await new Promise<ArrayBuffer>((resolve, reject) => {
      const reader = new FileReader();reader.onload = () => resolve(reader.result as ArrayBuffer);reader.onerror = () => reject(reader.error);reader.readAsArrayBuffer(original!);
    });
    expect(new Uint8Array(returned)).toEqual(new Uint8Array(bytes));
    file.content.html = '<p>Edited content must be exported.</p>';
    expect(await unchangedOfficeOriginal(file)).toBeUndefined();
    file.content = { ...originalContent, orientation: 'landscape' };
    expect(await unchangedOfficeOriginal(file)).toBeUndefined();
    expect(new Uint8Array(file.original.data)).toEqual(new Uint8Array(bytes));
  });
  it('retains ZIP and missing-provenance rejection on the renderer-independent original path', async () => {
    const file = newFile('word');
    file.original = { name: 'source.docx', data: new Uint8Array(60).buffer, contentFingerprint: await contentFingerprint(file.content) };
    await expect(unchangedOfficeOriginal(file)).rejects.toThrow('ZIP');
    delete file.original.contentFingerprint;
    expect(await unchangedOfficeOriginal(file)).toBeUndefined();
  });
  it('retains content identity across JSON key ordering and optional undefined fields', async () => {
    const a: WordContent = {
      kind: 'word',
      html: '<p>A</p>',
      paper: 'a4',
      margin: 'normal',
      orientation: undefined,
    };
    const b: WordContent = { margin: 'normal', paper: 'a4', html: '<p>A</p>', kind: 'word' };
    expect(await contentFingerprint(a)).toBe(await contentFingerprint(b));
    expect(await contentFingerprint({ ...a, html: '<p>Edited</p>' })).not.toBe(
      await contentFingerprint(a),
    );
    expect(await contentFingerprint({ ...a, orientation: 'landscape' })).not.toBe(
      await contentFingerprint(a),
    );
  });
  it('includes the imported content fingerprint in native backups', async () => {
    const file = newFile('word');
    file.original = {
      name: 'source.docx',
      data: new Uint8Array([1, 2]).buffer,
      contentFingerprint: await contentFingerprint(file.content),
    };
    expect(JSON.parse(nativeBackup(file)).original.contentFingerprint).toBe(
      file.original.contentFingerprint,
    );
  });
  it('rejects extreme XML nesting before recursive document conversion', () => {
    expect(() =>
      assertXmlComplexity('<root>' + '<table>'.repeat(256) + '</table>'.repeat(256) + '</root>'),
    ).toThrow('256 levels');
    expect(() =>
      assertXmlComplexity(
        '<root>' + '<table>'.repeat(254) + '<cell/>' + '</table>'.repeat(254) + '</root>',
      ),
    ).not.toThrow();
  });
  it('does not count tag-looking text in comments, CDATA, processing instructions or attributes', () => {
    expect(() =>
      assertXmlComplexity(
        '<?xml version="1.0"?><root><!-- <x><y> --><![CDATA[<x><y>]]><cell value="&lt; >"/></root>',
        2,
      ),
    ).not.toThrow();
    expect(() => assertXmlComplexity('<!DOCTYPE r [<!ENTITY x "text">]><r/>')).toThrow('DTD');
  });
});
