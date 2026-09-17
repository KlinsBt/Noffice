import { expect, it } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import { readPresentation } from './pptx-import';
import { contentSchema } from './model';
import { importFile, exportOffice } from './formats';
import { slideTextParagraphs } from './slide-text';
import { hydratePresentationText } from './presentation-migration';
import { contentFingerprint } from './office-preservation';
import { newFile, type DeckContent } from './model';
import JSZip from 'jszip';
import native from '../tests/fixtures/native-powerpoint-inheritance.json';

const input = () => {
  const bytes = fs.readFileSync('tests/fixtures/powerpoint-inheritance.pptx');
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(native.sourceSha256);
  return Uint8Array.from(bytes).buffer;
};
it('resolves two native layouts independently and records placeholder and group source identities', async () => {
  const deck = await readPresentation(input());
  expect(deck.slides).toHaveLength(2);
  expect(deck.slides[0].sourceMasterPath).not.toBe(deck.slides[1].sourceMasterPath);
  for (const [index, slide] of deck.slides.entries()) {
    const title = slide.elements.find((e) => e.text === `Master ${index + 1} title`)!;
    const expected = native.slides[index].shapes[0];
    expect(title.sourcePlaceholder).toMatchObject({
      index: '0',
      type: 'ctrTitle',
      geometry: 'layout',
      layoutShapeId: '2',
      masterShapeId: '2',
    });
    expect(title.x).toBeCloseTo((expected.x * 960) / 720, 4);
    expect(title.fontSize).toBeCloseTo((expected.characters![0].size * 960) / 720, 4);
    expect(title.fontFamily).toBe(expected.characters![0].font);
    expect(title.italic).toBe(true);
    expect(title.color).toBe('#000000');
    expect(title.align).toBe('center');
    expect(slide.elements.filter((e) => e.sourceGroupIds?.length === 2)).toHaveLength(2);
  }
  expect(contentSchema.parse(JSON.parse(JSON.stringify(deck)))).toEqual(deck);
});
it('retains and resolves every native rich-placeholder character instead of copying the first run', async () => {
  const deck = await readPresentation(input());
  for (const [index, slide] of deck.slides.entries()) {
    const rich = slide.elements.find((e) => e.text === 'Plain BOLD tail')!;
    const expected = native.slides[index].shapes[1].characters!;
    expect(rich.sourceText?.runs).toHaveLength(3);
    const characters = rich.sourceText!.runs.flatMap((r) =>
      [...r.text].map((text) => ({
        text,
        size: (r.fontSize * 720) / 960,
        font: r.fontFamily,
        bold: r.bold,
        italic: r.italic,
        underline: r.underline,
      })),
    );
    expect(characters).toEqual(expected.map(({ rgb: _, ...char }) => char));
  }
});

async function imported() {
  const data = input();
  return importFile(
    Object.assign(new File([data], 'Inheritance.pptx'), { arrayBuffer: async () => data }),
  );
}
async function exported(file: Awaited<ReturnType<typeof imported>>) {
  const blob = await exportOffice(file);
  return new Promise<ArrayBuffer>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(blob);
  });
}
it('replacing an entire rich run keeps the replaced run style in the view and actual PPTX', async () => {
  const file = await imported();
  const rich = (file.content as DeckContent).slides[0].elements.find(
    (e) => e.text === 'Plain BOLD tail',
  )!;
  rich.text = 'Plain BRAVE tail';
  const view = slideTextParagraphs(rich)[0];
  expect(view.find((r) => r.text === 'BRAVE')).toMatchObject({
    bold: true,
    fontSize: (28 * 960) / 720,
  });
  const output = await exported(file);
  const read = await readPresentation(output);
  expect(read.slides[0].elements.find((e) => e.text === rich.text)!.sourceText!.runs).toEqual(view);
  const before = await JSZip.loadAsync(input()),
    after = await JSZip.loadAsync(output);
  const group = (xml: string) =>
    new DOMParser().parseFromString(xml, 'application/xml').getElementsByTagNameNS('*', 'grpSp')[0];
  expect(
    group(await before.file('ppt/slides/slide1.xml')!.async('string')).isEqualNode(
      group(await after.file('ppt/slides/slide1.xml')!.async('string')),
    ),
  ).toBe(true);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml') {
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
    }
});
it('projects whole-object overrides and restored original text without mutating source runs', async () => {
  const deck = await readPresentation(input());
  const rich = deck.slides[0].elements.find((e) => e.text === 'Plain BOLD tail')!;
  const saved = JSON.stringify(rich.sourceText);
  expect(
    slideTextParagraphs({ ...rich, text: 'Plain BRAVE tail', color: '#ff0000' })[0].every(
      (r) => r.color === '#ff0000',
    ),
  ).toBe(true);
  expect(slideTextParagraphs(rich)[0]).toEqual(rich.sourceText!.runs);
  expect(JSON.stringify(rich.sourceText)).toBe(saved);
});
it('migrates legacy fonts and rich runs without losing distinct text/style edits, identities or original bytes', async () => {
  const data = input();
  const content = await readPresentation(data, true);
  const file = newFile('powerpoint', 'Legacy', content);
  file.original = {
    name: 'Legacy.pptx',
    data,
    contentFingerprint: await contentFingerprint(content),
  };
  const migrated = await hydratePresentationText(file);
  expect(migrated.revision).toBe(file.revision);
  expect(migrated.original!.data).toBe(data);
  expect(migrated.original!.contentFingerprint).toBe(await contentFingerprint(migrated.content));
  expect((migrated.content as DeckContent).slides[0].elements[0].fontSize).toBeCloseTo(
    (38 * 960) / 720,
  );
  expect(await exported(migrated)).toEqual(data);
  expect(await hydratePresentationText(migrated)).toBe(migrated);
  content.slides[0].elements[0].text = 'Edited legacy title';
  content.slides[0].elements[0].fontSize = 64;
  const edited = await hydratePresentationText(file);
  expect((edited.content as DeckContent).slides[0].elements[0]).toMatchObject({
    text: 'Edited legacy title',
    fontSize: 64,
    id: content.slides[0].elements[0].id,
  });
  const read = await readPresentation(await exported(edited));
  expect(read.slides[0].elements[0]).toMatchObject({ text: 'Edited legacy title', fontSize: 64 });
});
it('rejects unimplemented group rotation edits and malformed paragraph levels without mutating the original', async () => {
  const file = await imported();
  const original = Uint8Array.from(new Uint8Array(file.original!.data));
  (file.content as DeckContent).slides[0].elements.find((e) => e.sourceGroupIds?.length)!.rotation =
    30;
  await expect(exported(file)).rejects.toThrow('unsupported grouped');
  expect(new Uint8Array(file.original!.data)).toEqual(original);
  const zip = await JSZip.loadAsync(input());
  zip.file(
    'ppt/slides/slide1.xml',
    (await zip.file('ppt/slides/slide1.xml')!.async('string')).replace(
      '<a:p>',
      '<a:p><a:pPr lvl="99"/>',
    ),
  );
  await expect(readPresentation(await zip.generateAsync({ type: 'arraybuffer' }))).rejects.toThrow(
    'paragraph level',
  );
});
