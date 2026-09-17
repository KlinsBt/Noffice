import { it, expect } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { importFile, exportOffice } from './formats';
import { readPresentation } from './pptx-import';
import { groupSizeField } from './pptx-group-edit';
import { normalizeGroupBounds } from './pptx-group-bounds';
import { hydratePresentationText } from './presentation-migration';
import { contentFingerprint } from './office-preservation';
import type { DeckContent } from './model';
import native from '../tests/fixtures/native-powerpoint-child-bounds.json';
import ui from '../tests/fixtures/native-powerpoint-child-ui.json';
import angles from '../tests/fixtures/native-powerpoint-child-angles.json';

const bytes = () =>
  Uint8Array.from(fs.readFileSync('tests/fixtures/powerpoint-child-bounds.pptx')).buffer;
async function imported(data = bytes()) {
  return importFile(
    Object.assign(new File([data], 'Children.pptx'), { arrayBuffer: async () => data }),
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
const xml = (data: string) => new DOMParser().parseFromString(data, 'application/xml');
function checkGroups(doc: Document, expected: (typeof native.transforms)[number]) {
  const groups = [...doc.getElementsByTagNameNS('*', 'grpSp')];
  expect(groups).toHaveLength(expected.length);
  for (const [i, group] of groups.entries()) {
    expect(group.getElementsByTagNameNS('*', 'cNvPr')[0].getAttribute('id')).toBe(expected[i].id);
    const frame = group.getElementsByTagNameNS('*', 'xfrm')[0];
    for (const [key, value] of Object.entries(expected[i])) {
      if (key === 'id') continue;
      const [tag, attr] = key.split('.');
      expect(
        Math.abs(
          Number(frame.getElementsByTagNameNS('*', tag)[0].getAttribute(attr)) - Number(value),
        ),
        key,
      ).toBeLessThanOrEqual(12.7);
    }
  }
}
it.each(angles.rows)(
  'matches native $angle degree quadrant boundaries in four group contexts without changing leaf transforms',
  async (row) => {
    const original = fs.readFileSync('tests/fixtures/powerpoint-groups.pptx');
    expect(createHash('sha256').update(original).digest('hex')).toBe(angles.sourceSha256);
    const zip = await JSZip.loadAsync(original);
    for (let i = 0; i < 4; i++) {
      const path = `ppt/slides/slide${i + 1}.xml`,
        doc = xml(await zip.file(path)!.async('string'));
      const shapes = [...doc.getElementsByTagNameNS('*', 'sp')];
      const leaf = shapes.find(
        (s) => s.getElementsByTagNameNS('*', 'cNvPr')[0].getAttribute('id') === '2',
      )!;
      leaf
        .getElementsByTagNameNS('*', 'xfrm')[0]
        .setAttribute('rot', String(Math.round(row.angle * 60000)));
      zip.file(path, new XMLSerializer().serializeToString(doc));
    }
    const file = await imported(await zip.generateAsync({ type: 'arraybuffer' }));
    for (const slide of (file.content as DeckContent).slides)
      slide.elements.find((e) => e.sourceShapeId === '3')!.x += 16;
    const output = await JSZip.loadAsync(await exported(file));
    for (let i = 0; i < 4; i++)
      checkGroups(
        xml(await output.file(`ppt/slides/slide${i + 1}.xml`)!.async('string')),
        row.transforms[i],
      );
  },
);
it('matches independent native sequential size setters, all eight group frames and source preservation', async () => {
  const file = await imported(),
    deck = file.content as DeckContent,
    before = structuredClone(deck);
  for (const [i, slide] of deck.slides.entries()) {
    const leaf = slide.elements.find((e) => e.sourceShapeId === '3')!;
    for (const key of ['x', 'y', 'w', 'h'] as const) {
      const value = leaf[key] + (key === 'x' || key === 'w' ? 16 : 12);
      Object.assign(
        leaf,
        key === 'w' || key === 'h' ? groupSizeField(leaf, key, value) : { [key]: value },
      );
      const expected = native.steps
        .find(
          (s) =>
            s.slide === i + 1 &&
            s.property === { x: 'Left', y: 'Top', w: 'Width', h: 'Height' }[key],
        )!
        .shapes.find((s) => s.id === 3)!;
      for (const k of ['x', 'y', 'w', 'h'] as const)
        expect(Math.abs(leaf[k] * 0.75 - expected[k])).toBeLessThanOrEqual(0.001);
    }
  }
  const data = await exported(file),
    zip = await JSZip.loadAsync(data),
    original = await JSZip.loadAsync(bytes());
  for (let i = 0; i < 4; i++)
    checkGroups(
      xml(await zip.file(`ppt/slides/slide${i + 1}.xml`)!.async('string')),
      native.transforms[i],
    );
  for (const [path, entry] of Object.entries(original.files))
    if (!entry.dir && !/^ppt\/slides\/slide[1-4]\.xml$/.test(path))
      expect(await zip.file(path)!.async('uint8array'), path).toEqual(
        await entry.async('uint8array'),
      );
  const reopened = await readPresentation(data);
  for (const [i, slide] of reopened.slides.entries())
    for (const leaf of slide.elements) {
      const expected = deck.slides[i].elements.find((e) => e.sourceShapeId === leaf.sourceShapeId)!;
      for (const k of ['x', 'y', 'w', 'h'] as const) expect(leaf[k]).toBeCloseTo(expected[k], 3);
    }
  expect(file.original!.data).toEqual(bytes());
  expect(before.slides[2].elements.find((e) => e.sourceShapeId === '3')!.w).toBe(120);
});
it('matches native Size-pane corner retention using the actual committed physical sizes', async () => {
  const deck = await readPresentation(bytes());
  let leaf = deck.slides[2].elements.find((e) => e.sourceShapeId === '3')!;
  for (const [i, key] of (['x', 'y', 'w', 'h'] as const).entries()) {
    const expected = ui.stages[i + 1].frame,
      value = (expected[key] * 4) / 3;
    leaf = {
      ...leaf,
      ...(key === 'w' || key === 'h' ? groupSizeField(leaf, key, value) : { [key]: value }),
    };
    for (const k of ['x', 'y', 'w', 'h'] as const)
      expect(Math.abs(leaf[k] * 0.75 - expected[k])).toBeLessThanOrEqual(0.001);
  }
});
it('preserves unchanged exports and source-aware legacy frame migration after sibling edits', async () => {
  const file = await imported();
  expect(await exported(file)).toEqual(bytes());
  const deck = file.content as DeckContent,
    legacy = structuredClone(deck);
  for (const slide of legacy.slides)
    for (const leaf of slide.elements)
      if (leaf.sourceGroupTransform) delete leaf.sourceGroupTransform.dimensions;
  file.original!.contentFingerprint = await contentFingerprint(legacy);
  legacy.slides[0].elements[1].x += 16;
  const migrated = await hydratePresentationText({ ...file, content: legacy });
  const restored = migrated.content as DeckContent;
  expect(restored.slides[0].elements[1].x).toBeCloseTo(deck.slides[0].elements[1].x + 16, 8);
  expect(restored.slides[0].elements[1].sourceGroupTransform!.dimensions).toEqual([
    9144000, 5143500,
  ]);
  expect(migrated.revision).toBe(file.revision);
  expect(migrated.original!.data).toEqual(file.original!.data);
});
it('rejects quadrant rounding amplified by coarse coordinates', () => {
  const doc = xml(
    '<spTree><grpSp><grpSpPr><xfrm><off x="0" y="0"/><ext cx="9144000" cy="5143500"/><chOff x="0" y="0"/><chExt cx="10" cy="10"/></xfrm></grpSpPr><sp><spPr><xfrm rot="2700000"><off x="1" y="1"/><ext cx="3" cy="2"/></xfrm><prstGeom prst="rect"/></spPr></sp></grpSp></spTree>',
  );
  expect(() => normalizeGroupBounds([doc.getElementsByTagName('sp')[0]], 9144000, 5143500)).toThrow(
    'coarse coordinates',
  );
});
