import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { flippedImageFixture } from '../tests/fixtures/pptx-retained';

it('imports and edits image flips without changing media or unrelated package parts', async () => {
  const input = await flippedImageFixture(),
    file = await importFile(inputFile(input));
  const image = (file.content as DeckContent).slides[0].elements.find((e) => e.type === 'image')!;
  expect(image).toMatchObject({ flipH: true, flipV: false });
  image.flipH = false;
  image.flipV = true;
  image.rotation = 30;
  const output = await bytes(await exportOffice(file)),
    before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const read = await readPresentation(output);
  expect(read.slides[0].elements.find((e) => e.type === 'image')).toMatchObject({
    flipH: false,
    flipV: true,
    rotation: 30,
  });
  delete image.flipH;
  delete image.flipV;
  image.rotation = 60;
  const legacy = await readPresentation(await bytes(await exportOffice(file)));
  expect(legacy.slides[0].elements.find((e) => e.type === 'image')).toMatchObject({
    flipH: true,
    flipV: false,
    rotation: 60,
  });
});
it('writes image flips for new decks and added pictures in retained decks', async () => {
  const { file, content } = await fixture();
  content.slides[0].elements.push(
    textElement('', {
      type: 'image',
      src: `data:image/png;base64,${pictureData}`,
      flipH: true,
      flipV: true,
    }),
  );
  for (const retained of [true, false]) {
    if (!retained) delete file.original;
    const read = await readPresentation(await bytes(await exportOffice(file)));
    expect(read.slides[0].elements.find((e) => e.type === 'image')).toMatchObject({
      flipH: true,
      flipV: true,
    });
  }
});
import { retainedPresentationFixture, pictureData } from '../tests/fixtures/pptx-retained';
import { importFile, exportOffice } from './formats';
import { elements, parseXML } from './xlsx-import';
import { contentSchema, textElement, type DeckContent } from './model';
import { readPresentation } from './pptx-import';
const inputFile = (data: ArrayBuffer) =>
  Object.assign(new File([data], 'Source.pptx'), { arrayBuffer: async () => data });
async function bytes(blob: Blob) {
  return new Promise<ArrayBuffer>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.onerror = reject;
    reader.readAsArrayBuffer(blob);
  });
}
async function fixture() {
  const input = await retainedPresentationFixture();
  const file = await importFile(inputFile(input));
  return { input, file, content: file.content as DeckContent };
}
it('keeps source rotation when an older saved model has no rotation metadata', async () => {
  const zip = await JSZip.loadAsync(await retainedPresentationFixture());
  const doc = parseXML(await zip.file('ppt/slides/slide1.xml')!.async('string'));
  elements(doc, 'xfrm')[0].setAttribute('rot', '2700000');
  zip.file('ppt/slides/slide1.xml', new XMLSerializer().serializeToString(doc));
  const file = await importFile(inputFile(await zip.generateAsync({ type: 'arraybuffer' })));
  const content = file.content as DeckContent;
  delete content.slides[0].elements[0].rotation;
  content.slides[0].elements[0].text = 'Edited older import';
  const output = await JSZip.loadAsync(await bytes(await exportOffice(file)));
  expect(
    elements(
      parseXML(await output.file('ppt/slides/slide1.xml')!.async('string')),
      'xfrm',
    )[0].getAttribute('rot'),
  ).toBe('2700000');
});
it.each(['slide', 'layout', 'master'])(
  'resolves placeholder geometry from the %s without overlapping inherited title positions',
  async (source) => {
    const zip = await JSZip.loadAsync(await retainedPresentationFixture());
    const p = 'http://schemas.openxmlformats.org/presentationml/2006/main';
    const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const transform =
      '<a:xfrm><a:off x="1828800" y="2743200"/><a:ext cx="3657600" cy="914400"/></a:xfrm>';
    const shape = (where: string) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="50" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="${where === 'slide' ? 'ctrTitle' : 'title'}" idx="${where === 'master' ? '42' : '7'}"/></p:nvPr></p:nvSpPr><p:spPr>${where === source ? transform : ''}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Inherited title</a:t></a:r></a:p></p:txBody></p:sp>`;
    for (const [where, path, root] of [
      ['slide', 'ppt/slides/slide1.xml', 'sld'],
      ['layout', 'ppt/slideLayouts/slideLayout1.xml', 'sldLayout'],
      ['master', 'ppt/slideMasters/slideMaster1.xml', 'sldMaster'],
    ])
      zip.file(
        path,
        `<p:${root} xmlns:p="${p}" xmlns:a="${a}"><p:cSld><p:spTree>${shape(where)}</p:spTree></p:cSld></p:${root}>`,
      );
    const model = await readPresentation(await zip.generateAsync({ type: 'arraybuffer' }));
    const title = model.slides[0].elements[0];
    expect(title.x).toBe(192);
    expect(title.y).toBe(216);
    expect(title.w).toBe(384);
    expect(title.h).toBe(72);
  },
);
it('retains custom slide proportions through native backup validation and new-deck export', async () => {
  const { file, content } = await fixture();
  content.slides[0].elements[0].rotation = 135;
  delete file.original;
  file.content = contentSchema.parse(JSON.parse(JSON.stringify(content)));
  const output = await bytes(await exportOffice(file));
  const read = await readPresentation(output);
  expect(read.aspectRatio).toBeCloseTo(4 / 3, 6);
  expect(read.slides[0].elements[0].rotation).toBe(135);
  expect(read.slides[0].elements[0].y).toBeCloseTo(content.slides[0].elements[0].y, 3);
  expect(read.slides[0].elements[0].h).toBeCloseTo(content.slides[0].elements[0].h, 3);
});
it('edits existing rich text and geometry while retaining run properties, notes, chart payloads and original aspect ratio', async () => {
  const { input, file, content } = await fixture();
  const text = content.slides[0].elements[0];
  expect(content.aspectRatio).toBeCloseTo(4 / 3, 8);
  text.text = 'Original revised title';
  text.x = 150;
  text.w = 650;
  text.rotation = 45;
  const output = await bytes(await exportOffice(file)),
    after = await JSZip.loadAsync(output),
    before = await JSZip.loadAsync(input);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const xml = parseXML(await after.file('ppt/slides/slide1.xml')!.async('string')),
    shape = elements(xml, 'sp')[0];
  expect(
    elements(shape, 't')
      .map((t) => t.textContent)
      .join(''),
  ).toBe('Original revised title');
  expect(elements(shape, 'rPr').map((r) => elements(r, 'srgbClr')[0]?.getAttribute('val'))).toEqual(
    ['B00020', '0033AA'],
  );
  expect(elements(shape, 'rPr')[0].getAttribute('b')).toBe('1');
  expect(elements(shape, 'rPr')[1].getAttribute('i')).toBe('1');
  const reimport = (await importFile(inputFile(output))).content as DeckContent;
  expect(reimport.slides[0].elements[0].x).toBeCloseTo(150, 4);
  expect(reimport.slides[0].elements[0].w).toBeCloseTo(650, 4);
  expect(reimport.slides[0].elements[0].rotation).toBe(45);
});
it('exports chosen font, emphasis, color, alignment, shape fill, background and existing notes', async () => {
  const { input, file, content } = await fixture(),
    slide = content.slides[0];
  Object.assign(slide.elements[0], {
    fontFamily: 'Inter',
    fontSize: 32,
    bold: false,
    italic: false,
    underline: true,
    color: '#008855',
    align: 'center',
  });
  slide.elements[1].fill = '#112233';
  slide.background = '#EEEEDD';
  slide.notes = 'Revised notes\nSecond paragraph';
  const output = await bytes(await exportOffice(file)),
    zip = await JSZip.loadAsync(output),
    original = await JSZip.loadAsync(input);
  for (const path of Object.keys(original.files))
    if (
      !original.files[path].dir &&
      !['ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml'].includes(path)
    )
      expect(await zip.file(path)!.async('uint8array'), path).toEqual(
        await original.file(path)!.async('uint8array'),
      );
  const xml = parseXML(await zip.file('ppt/slides/slide1.xml')!.async('string'));
  expect(elements(xml, 'rPr').every((r) => r.getAttribute('u') === 'sng')).toBe(true);
  expect(elements(xml, 'pPr')[0].getAttribute('algn')).toBe('ctr');
  const read = (await importFile(inputFile(output))).content as DeckContent;
  expect(read.slides[0].notes).toBe('Revised notes\nSecond paragraph');
  expect(read.slides[0].elements[0].fontFamily).toBe('Inter');
  expect(read.slides[0].elements[0].fontSize).toBeCloseTo(32, 1);
  expect(read.slides[0].elements[1].fill).toBe('#112233');
  expect(read.slides[0].background).toBe('#eeeedd');
});
it('rejects unsupported imported structure edits instead of rebuilding a reduced presentation', async () => {
  const { file, content } = await fixture();
  const removed = content.slides.pop()!;
  await expect(exportOffice(file)).rejects.toThrow('original slides');
  content.slides.push(removed);
  content.slides[0].elements.pop();
  await expect(exportOffice(file)).rejects.toThrow('Removing');
});
it('reorders imported slides by their existing relationships and edits the correct source slide', async () => {
  const { input, file, content } = await fixture();
  content.slides.reverse();
  content.slides[1].elements[0].text = 'Edited after reorder';
  const output = await bytes(await exportOffice(file)),
    after = await JSZip.loadAsync(output),
    before = await JSZip.loadAsync(input);
  for (const path of Object.keys(before.files))
    if (
      !before.files[path].dir &&
      !['ppt/presentation.xml', 'ppt/slides/slide1.xml'].includes(path)
    )
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const imported = await readPresentation(output);
  expect(imported.slides.map((s) => s.sourcePath)).toEqual([
    'ppt/slides/slide2.xml',
    'ppt/slides/slide1.xml',
  ]);
  expect(imported.slides[0].notes).toBe('Chart notes');
  expect(imported.slides[1].elements[0].text).toBe('Edited after reorder');
});
it('refuses duplicate slide identities and section-dependent reordering', async () => {
  const { input, file, content } = await fixture();
  content.slides[1].sourcePath = content.slides[0].sourcePath;
  await expect(exportOffice(file)).rejects.toThrow('original slides');
  const zip = await JSZip.loadAsync(input),
    doc = parseXML(await zip.file('ppt/presentation.xml')!.async('string'));
  const section = doc.createElementNS(
    'http://schemas.microsoft.com/office/powerpoint/2010/main',
    'p14:sectionLst',
  );
  doc.documentElement.append(section);
  zip.file('ppt/presentation.xml', new XMLSerializer().serializeToString(doc));
  const modified = await importFile(inputFile(await zip.generateAsync({ type: 'arraybuffer' })));
  (modified.content as DeckContent).slides.reverse();
  await expect(exportOffice(modified)).rejects.toThrow('section-aware');
});
it('edits across run boundaries, appends paragraphs and preserves unrelated source runs', async () => {
  const { file, content } = await fixture();
  content.slides[0].elements[0].text = 'New text\nAdded paragraph';
  const output = await bytes(await exportOffice(file));
  const read = (await importFile(inputFile(output))).content as DeckContent;
  expect(read.slides[0].elements[0].text).toBe('New text\nAdded paragraph');
});
it('adds a local image and first notes with a registered notes master while retaining chart content', async () => {
  const input = await retainedPresentationFixture(false),
    file = await importFile(inputFile(input));
  const content = file.content as DeckContent;
  expect(content.slides[0].notes).toBe('');
  content.slides[0].notes = 'New speaker notes';
  content.slides[0].elements.push(
    textElement('', {
      type: 'image',
      src: 'data:image/png;base64,' + pictureData,
      x: 10,
      y: 20,
      w: 30,
      h: 40,
    }),
  );
  const output = await bytes(await exportOffice(file)),
    zip = await JSZip.loadAsync(output),
    before = await JSZip.loadAsync(input);
  for (const path of Object.keys(before.files))
    if (
      !before.files[path].dir &&
      ![
        'ppt/slides/slide1.xml',
        'ppt/slides/_rels/slide1.xml.rels',
        'ppt/presentation.xml',
        'ppt/_rels/presentation.xml.rels',
        '[Content_Types].xml',
      ].includes(path)
    )
      expect(await zip.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const read = (await importFile(inputFile(output))).content as DeckContent;
  expect(read.slides[0].notes).toBe('New speaker notes');
  expect(
    Object.keys(zip.files).filter((p) => /^ppt\/notesMasters\/[^/]+\.xml$/.test(p)),
  ).toHaveLength(1);
  expect(await zip.file('ppt/notesSlides/_rels/notesSlide1.xml.rels')!.async('string')).toContain(
    '/notesMaster',
  );
  const presentation = parseXML(await zip.file('ppt/presentation.xml')!.async('string'));
  expect(elements(presentation, 'notesMasterId')).toHaveLength(1);
  expect(await zip.file('ppt/theme/noffice-notes-theme1.xml')!.async('uint8array')).toEqual(
    await before.file('ppt/theme/theme1.xml')!.async('uint8array'),
  );
  expect(await zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels')!.async('string')).toContain(
    'noffice-notes-theme1.xml',
  );
  expect(elements(presentation, 'sldId').map((e) => e.getAttribute('id'))).toEqual(
    elements(parseXML(await before.file('ppt/presentation.xml')!.async('string')), 'sldId').map(
      (e) => e.getAttribute('id'),
    ),
  );
  expect(read.slides[0].elements.at(-1)?.src).toBe('data:image/png;base64,' + pictureData);
});

it('reuses one separate notes master/theme for first notes on two slides', async () => {
  const input = await retainedPresentationFixture(false);
  const file = await importFile(inputFile(input));
  (file.content as DeckContent).slides.forEach((s, i) => {
    s.notes = `Notes ${i + 1}`;
  });
  const output = await bytes(await exportOffice(file));
  const zip = await JSZip.loadAsync(output);
  expect(
    Object.keys(zip.files).filter((p) => /^ppt\/notesMasters\/[^/]+\.xml$/.test(p)),
  ).toHaveLength(1);
  expect(
    Object.keys(zip.files).filter((p) => /^ppt\/theme\/noffice-notes-theme\d+\.xml$/.test(p)),
  ).toHaveLength(1);
  expect((await readPresentation(output)).slides.map((s) => s.notes)).toEqual([
    'Notes 1',
    'Notes 2',
  ]);
});

it('rejects first-note export when a retained theme relationship is broken', async () => {
  const zip = await JSZip.loadAsync(await retainedPresentationFixture(false));
  zip.remove('ppt/theme/theme1.xml');
  const input = await zip.generateAsync({ type: 'arraybuffer' });
  const file = await importFile(inputFile(input));
  (file.content as DeckContent).slides[0].notes = 'Needs a valid theme';
  await expect(exportOffice(file)).rejects.toThrow('missing theme');
  expect(file.original!.data).toEqual(input);
});
