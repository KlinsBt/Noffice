import { it, expect } from 'vitest';
import fs from 'node:fs';
import JSZip from 'jszip';
import { createHash } from 'node:crypto';
import { importFile, exportOffice } from './formats';
import { readPresentation } from './pptx-import';
import { hydratePresentationText, needsPresentationText } from './presentation-migration';
import { contentFingerprint } from './office-preservation';
import { coarseGroupGrid, quantizeGroupEdit } from './pptx-group-edit';
import type { DeckContent } from './model';
import native from '../tests/fixtures/native-powerpoint-coarse-ui.json';

const input = () => {
  const bytes = fs.readFileSync('tests/fixtures/powerpoint-coarse.pptx');
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(native.sourceSha256);
  return Uint8Array.from(bytes).buffer;
};
const imported = async () => {
  const data = input();
  return importFile(
    Object.assign(new File([data], 'Coarse.pptx'), { arrayBuffer: async () => data }),
  );
};
const exported = async (file: Awaited<ReturnType<typeof imported>>) => {
  const blob = await exportOffice(file);
  return new Promise<ArrayBuffer>((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result as ArrayBuffer);
    reader.readAsArrayBuffer(blob);
  });
};
it('matches native UI position commits and keeps the source grid stable after repeated edits', async () => {
  const book = await readPresentation(input());
  let item = book.slides[0].elements[0];
  expect(coarseGroupGrid(item)).toEqual({ x: 4, y: 3 });
  for (const [key, target, stage] of [
    ['x', 121.6, 'left'],
    ['y', 141.25, 'top'],
  ] as const) {
    const before = structuredClone(item);
    item = quantizeGroupEdit(item, { ...item, [key]: (target * 4) / 3 });
    expect(before.sourceGroupTransform).toEqual(item.sourceGroupTransform);
    const expected = native.stages.find((s) => s.stage === stage)!.frame;
    for (const k of ['x', 'y', 'w', 'h'] as const)
      expect(item[k]).toBeCloseTo((expected[k] * 4) / 3, 8);
  }
  const wider = quantizeGroupEdit(item, { ...item, w: item.w + 2.2 });
  expect(wider.w).toBe(164);
  expect(quantizeGroupEdit(wider, { ...wider, x: wider.x + 1 }).x).toBe(wider.x);
});
it('preserves native no-ops and rejects unmeasured coarse transforms and unsafe extents', async () => {
  const book = await readPresentation(input()),
    item = book.slides[0].elements[0];
  expect(quantizeGroupEdit(item, { ...item, x: item.x + 1, y: item.y + 0.5 })).toEqual(item);
  for (const i of [2, 3]) {
    const rotated = book.slides[i].elements[0];
    expect(() => quantizeGroupEdit(rotated, { ...rotated, x: rotated.x + 1 })).toThrow(
      'rotated or reflected',
    );
  }
  expect(() => quantizeGroupEdit(item, { ...item, w: 0 })).toThrow('cannot be represented');
  expect(() => quantizeGroupEdit(item, { ...item, x: Infinity })).toThrow('cannot be represented');
});
it('exports the snapped visible frame and preserves unrelated package payloads', async () => {
  const file = await imported(),
    book = file.content as DeckContent;
  const source = structuredClone(book);
  let item = book.slides[0].elements[0];
  item = quantizeGroupEdit(item, { ...item, x: (121.6 * 4) / 3 });
  item = quantizeGroupEdit(item, { ...item, y: (141.25 * 4) / 3 });
  book.slides[0].elements[0] = item;
  const bytes = await exported(file),
    reopened = await readPresentation(bytes);
  for (const k of ['x', 'y', 'w', 'h'] as const)
    expect(reopened.slides[0].elements[0][k]).toBeCloseTo(item[k], 8);
  for (let s = 0; s < 4; s++)
    for (let e = 0; e < 3; e++)
      if (s !== 0 || e !== 0)
        for (const k of ['x', 'y', 'w', 'h'] as const)
          expect(reopened.slides[s].elements[e][k]).toBeCloseTo(source.slides[s].elements[e][k], 8);
  const before = await JSZip.loadAsync(input()),
    after = await JSZip.loadAsync(bytes);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  expect(file.original!.data).toEqual(input());
});
it.each([false, true])(
  'hydrates legacy dimensions without losing edits or changing source/revision (edited=%s)',
  async (edited) => {
    const file = await imported(),
      book = file.content as DeckContent;
    for (const slide of book.slides)
      for (const item of slide.elements) delete item.sourceGroupTransform!.dimensions;
    file.original!.contentFingerprint = await contentFingerprint(book);
    if (edited) book.slides[0].elements[0].x += 4;
    expect(needsPresentationText(file)).toBe(true);
    const migrated = await hydratePresentationText(file),
      next = migrated.content as DeckContent;
    expect(next.slides[0].elements[0].x).toBe(edited ? 164 : 160);
    expect(next.slides[0].elements[0].sourceGroupTransform!.dimensions).toEqual([9144000, 5143500]);
    expect(migrated.revision).toBe(file.revision);
    expect(migrated.original!.data).toEqual(file.original!.data);
    expect(needsPresentationText(migrated)).toBe(false);
    if (!edited) expect(await exported(migrated)).toEqual(input());
  },
);
