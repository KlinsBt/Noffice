import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { paintPresentationFixture } from '../tests/fixtures/pptx-paints';
import { importFile, exportOffice } from './formats';
import { type DeckContent } from './model';
import { readPresentation } from './pptx-import';
async function bytes(blob: Blob) {
  return new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = reject;
    reader.readAsArrayBuffer(blob);
  });
}
async function fixture() {
  const input = await paintPresentationFixture();
  const file = await importFile(
    Object.assign(new File([input], 'Paints.pptx'), { arrayBuffer: async () => input }),
  );
  return { input, file, content: file.content as DeckContent };
}
it('imports text fills, transparency, no-fill shapes and theme colors on inherited backgrounds', async () => {
  const { content } = await fixture();
  expect(content.slides[0].background).toBe('#bbddee');
  expect(content.slides[0].elements.map((e) => [e.fill, e.fillOpacity])).toEqual([
    ['#ddeeaa', 0.5],
    ['transparent', 0],
    ['#345678', 1],
  ]);
});
it('retains source theme/background parts while editing a text fill and transparency', async () => {
  const { input, file, content } = await fixture();
  content.slides[0].elements[0].fill = '#223344';
  content.slides[0].elements[0].fillOpacity = 0.75;
  const output = await bytes(await exportOffice(file)),
    before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const read = await readPresentation(output);
  expect(read.slides[0].background).toBe('#bbddee');
  expect(read.slides[0].elements[0]).toMatchObject({ fill: '#223344', fillOpacity: 0.75 });
  expect(read.slides[0].elements[2].fill).toBe('#345678');
});
it('exports text-box fills and transparency in newly authored decks too', async () => {
  const { file, content } = await fixture();
  delete file.original;
  const read = await readPresentation(await bytes(await exportOffice(file)));
  expect(read.slides[0].elements[0]).toMatchObject({
    fill: content.slides[0].elements[0].fill,
    fillOpacity: 0.5,
  });
  expect(read.slides[0].background).toBe('#bbddee');
});
