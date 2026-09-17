import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { presentationPaints } from './drawing-colors';
import { outlineDashes, outlineDashArray } from './slide-outline';
import { paintPresentationFixture } from '../tests/fixtures/pptx-paints';
import { importFile, exportOffice } from './formats';
import { readPresentation } from './pptx-import';
import type { DeckContent } from './model';
const xml = (text: string) => new DOMParser().parseFromString(text, 'application/xml');
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
    Object.assign(new File([input], 'Outlines.pptx'), { arrayBuffer: async () => input }),
  );
  return { input, file, content: file.content as DeckContent };
}
it('resolves direct overrides over theme line styles and keeps fill and outline independent', () => {
  const theme = xml(
    '<theme><clrScheme><accent1><srgbClr val="123456"/></accent1></clrScheme><lnStyleLst><ln w="25400"><solidFill><schemeClr val="phClr"><alpha val="40000"/></schemeClr></solidFill><prstDash val="dash"/></ln></lnStyleLst></theme>',
  );
  const source = xml(
    '<slide><sp><spPr><solidFill><srgbClr val="FFFFFF"/></solidFill><ln w="50800"/></spPr><style><lnRef idx="1"><schemeClr val="accent1"/></lnRef></style></sp></slide>',
  );
  const paints = presentationPaints(theme, null, null, source);
  expect(paints.outline([source.querySelector('sp')], 12192000)).toEqual({
    color: '#123456',
    opacity: 0.4,
    width: 4,
    dash: 'dash',
  });
  const noFill = xml('<sp><spPr><ln><noFill/></ln></spPr></sp>');
  expect(
    paints.outline([noFill.documentElement, source.querySelector('sp')], 12192000),
  ).toMatchObject({ color: 'transparent', opacity: 0 });
  for (const line of [
    '<ln w="NaN"><solidFill><srgbClr val="123456"/></solidFill></ln>',
    '<ln><gradFill/></ln>',
    '<ln><custDash/></ln>',
  ])
    expect(
      paints.outline([xml(`<sp><spPr>${line}</spPr></sp>`).documentElement], 12192000),
    ).toBeUndefined();
});
it('imports a no-fill rectangle with its visible red outline', async () => {
  const { content } = await fixture();
  expect(content.slides[0].elements[1]).toMatchObject({
    fill: 'transparent',
    outline: { color: '#ff0000', width: 2, opacity: 1, dash: 'solid' },
  });
});
it('patches only edited line properties and preserves unrelated package payloads', async () => {
  const { file, content, input } = await fixture();
  content.slides[0].elements[1].outline = {
    color: '#224466',
    width: 6,
    opacity: 0.6,
    dash: 'dashDot',
  };
  const output = await bytes(await exportOffice(file));
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  expect((await readPresentation(output)).slides[0].elements[1].outline).toEqual(
    content.slides[0].elements[1].outline,
  );
  content.slides[0].elements[1].outline!.color = 'transparent';
  expect(
    (await readPresentation(await bytes(await exportOffice(file)))).slides[0].elements[1].outline
      ?.color,
  ).toBe('transparent');
});
it.each(outlineDashes)('exports new text and shape outlines with preset %s', async (dash) => {
  const { file, content } = await fixture();
  delete file.original;
  for (const el of content.slides[0].elements)
    el.outline = { color: '#224466', width: 3, opacity: 0.6, dash };
  const read = await readPresentation(await bytes(await exportOffice(file)));
  for (const el of read.slides[0].elements)
    expect(el.outline).toEqual({ color: '#224466', width: 3, opacity: 0.6, dash });
  expect(outlineDashArray(read.slides[0].elements[0].outline!)).toEqual(
    dash === 'solid' ? undefined : expect.any(String),
  );
});
