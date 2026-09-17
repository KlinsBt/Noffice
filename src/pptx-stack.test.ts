import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { stackPresentationFixture } from '../tests/fixtures/pptx-stack';
import { exportOffice, importFile } from './formats';
import { type DeckContent, textElement, contentSchema } from './model';
import { readPresentation } from './pptx-import';
import { elements, parseXML } from './xlsx-import';
import { stackSlide, stackKey } from './slide-stack';
import { pptxStackNodes } from './pptx-stack';

async function fixture(source?: ArrayBuffer) {
  const input = source || (await stackPresentationFixture());
  const file = await importFile(
    Object.assign(new File([input], 'Layers.pptx'), { arrayBuffer: async () => input }),
  );
  return { input, file, content: file.content as DeckContent };
}
async function output(blob: Blob) {
  return new Promise<ArrayBuffer>((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as ArrayBuffer);
    r.onerror = reject;
    r.readAsArrayBuffer(blob);
  });
}
it('imports interleaved images and text in source order, retaining opaque chart layers', async () => {
  const { content } = await fixture();
  expect(content.slides[0].elements.map((e) => e.type)).toEqual(['text', 'image', 'text']);
  expect(content.slides[0].stackOrder).toHaveLength(4);
  expect(contentSchema.parse(content)).toEqual(content);
});
it('patches the reordered source shapes by identity and preserves every other package payload', async () => {
  const { input, file, content } = await fixture();
  const before = await JSZip.loadAsync(input);
  const first = content.slides[0].elements[0];
  content.slides[0] = stackSlide(content.slides[0], [first.id], 'front');
  content.slides[0].elements.at(-1)!.text = 'Bottom text edited';
  const bytes = await output(await exportOffice(file));
  const after = await JSZip.loadAsync(bytes);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  const doc = parseXML(await after.file('ppt/slides/slide1.xml')!.async('string'));
  expect(
    [...pptxStackNodes(elements(doc, 'spTree')[0]).values()].map((e) =>
      elements(e, 'cNvPr')[0]?.getAttribute('name'),
    ),
  ).toEqual(['Middle image', 'Retained chart', 'Top text', 'Bottom text']);
  const imported = await readPresentation(bytes);
  expect(imported.slides[0].elements.at(-1)?.text).toBe('Bottom text edited');
  expect(imported.slides[0].elements[1].text).toBe('Top text');
});
it('places a new picture between retained layers without moving or rebuilding the source objects', async () => {
  const { input, file, content } = await fixture();
  const image = textElement('', { type: 'image', src: content.slides[0].elements[1].src });
  content.slides[0].elements.push(image);
  content.slides[0] = stackSlide(content.slides[0], [image.id], 'backward');
  const expected = content.slides[0].elements.map((e) => e.type);
  const after = await JSZip.loadAsync(await output(await exportOffice(file)));
  const before = await JSZip.loadAsync(input);
  const original = parseXML(await before.file('ppt/slides/slide1.xml')!.async('string'));
  const edited = parseXML(await after.file('ppt/slides/slide1.xml')!.async('string'));
  const originals = pptxStackNodes(elements(original, 'spTree')[0]);
  const patched = pptxStackNodes(elements(edited, 'spTree')[0]);
  for (const [key, node] of originals) expect(patched.get(key)?.outerHTML).toBe(node.outerHTML);
  expect(
    (
      await readPresentation(await after.generateAsync({ type: 'arraybuffer' }))
    ).slides[0].elements.map((e) => e.type),
  ).toEqual(expected);
});
it('rejects missing, repeated and forged source layers without rebuilding the package', async () => {
  for (const corruption of ['missing', 'duplicate', 'forged']) {
    const { file, content } = await fixture();
    const slide = content.slides[0];
    if (corruption === 'missing') slide.stackOrder!.splice(2, 1);
    if (corruption === 'duplicate') slide.stackOrder!.push(slide.stackOrder![0]);
    if (corruption === 'forged') slide.stackOrder![2] = 'source:999';
    await expect(exportOffice(file)).rejects.toThrow('stacking order');
  }
});
it('treats groups as one retained layer and refuses moving individual group children', async () => {
  const zip = await JSZip.loadAsync(await stackPresentationFixture());
  const doc = parseXML(await zip.file('ppt/slides/slide1.xml')!.async('string'));
  const tree = elements(doc, 'spTree')[0];
  const group = doc.createElementNS(tree.namespaceURI, 'p:grpSp');
  const nv = doc.createElementNS(tree.namespaceURI, 'p:nvGrpSpPr');
  const props = doc.createElementNS(tree.namespaceURI, 'p:cNvPr');
  props.setAttribute('id', '99');
  nv.appendChild(props);
  group.appendChild(nv);
  const shape = elements(tree, 'sp')[0];
  tree.insertBefore(group, shape);
  group.appendChild(shape);
  zip.file('ppt/slides/slide1.xml', new XMLSerializer().serializeToString(doc));
  const { content, file } = await fixture(await zip.generateAsync({ type: 'arraybuffer' }));
  const slide = content.slides[0];
  expect(slide.elements[0].sourceStackKey).toBe('source:99');
  expect(() => stackSlide(slide, [slide.elements[0].id], 'front')).toThrow('group');
  content.slides[0] = stackSlide(slide, [slide.elements[1].id], 'back');
  expect(content.slides[0].stackOrder![0]).toBe(stackKey(slide.elements[1]));
  const read = await readPresentation(await output(await exportOffice(file)));
  expect(read.slides[0].elements[0].type).toBe('image');
});
