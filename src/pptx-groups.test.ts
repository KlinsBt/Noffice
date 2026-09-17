import { it, expect } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { readPresentation } from './pptx-import';
import { importFile, exportOffice } from './formats';
import { contentSchema, type DeckContent } from './model';
import native from '../tests/fixtures/native-powerpoint-groups.json';
import edits from '../tests/fixtures/native-powerpoint-group-edits.json';
import bounds from '../tests/fixtures/native-powerpoint-group-bounds.json';
import { localGroupGeometry } from './pptx-group-transform';
import { normalizeGroupBounds } from './pptx-group-bounds';
import { hydratePresentationText } from './presentation-migration';
import { contentFingerprint } from './office-preservation';

const input = () => {
  const data = fs.readFileSync('tests/fixtures/powerpoint-groups.pptx');
  expect(createHash('sha256').update(data).digest('hex')).toBe(native.sourceSha256);
  return Uint8Array.from(data).buffer;
};
async function imported() {
  const data = input();
  return importFile(
    Object.assign(new File([data], 'Groups.pptx'), { arrayBuffer: async () => data }),
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
it('projects nested scaled, translated, rotated and flipped leaf frames to native effective coordinates', async () => {
  const deck = await readPresentation(input());
  for (const [i, slide] of deck.slides.entries()) {
    for (const expected of native.effective[i]) {
      const element = slide.elements.find((e) => e.sourceShapeId === String(expected.id))!;
      for (const key of ['x', 'y', 'w', 'h'] as const)
        expect(element[key], `slide ${i + 1} ${expected.name} ${key}`).toBeCloseTo(
          (expected[key] * 4) / 3,
          3,
        );
      expect(element.rotation || 0).toBeCloseTo(expected.rotation, 3);
    }
  }
  expect(contentSchema.parse(JSON.parse(JSON.stringify(deck)))).toEqual(deck);
});
it('migrates saved group-local coordinates, preserving distinct edits and unchanged original export', async () => {
  const file = await imported();
  file.content = await readPresentation(input(), false, true);
  file.original!.contentFingerprint = await contentFingerprint(file.content);
  const migrated = await hydratePresentationText(file);
  expect((migrated.content as DeckContent).groupModelVersion).toBe(1);
  expect((migrated.content as DeckContent).slides[0].elements[0].x).toBeCloseTo(160);
  expect(migrated.revision).toBe(file.revision);
  expect(await exported(migrated)).toEqual(input());
  const item = (file.content as DeckContent).slides[0].elements[0];
  item.x = 176;
  const edited = await hydratePresentationText(file);
  expect((edited.content as DeckContent).slides[0].elements[0].x).toBe(176);
  expect((await readPresentation(await exported(edited))).slides[0].elements[0].x).toBeCloseTo(176);
});
it('rejects changes to grouped child rotation without touching source bytes, and refuses singular transforms', async () => {
  const file = await imported(),
    original = input();
  (file.content as DeckContent).slides[2].elements[0].rotation = 45;
  await expect(exported(file)).rejects.toThrow('unsupported grouped');
  expect(file.original!.data).toEqual(original);
  const zip = await JSZip.loadAsync(original);
  const path = 'ppt/slides/slide1.xml';
  zip.file(
    path,
    (await zip.file(path)!.async('string')).replace(
      'cx="2540000" cy="1143000"',
      'cx="0" cy="1143000"',
    ),
  );
  await expect(readPresentation(await zip.generateAsync({ type: 'arraybuffer' }))).rejects.toThrow(
    'extents',
  );
});
it('reports unrepresentable skew instead of silently claiming an editable frame', async () => {
  const zip = await JSZip.loadAsync(input());
  const path = 'ppt/slides/slide1.xml';
  // Rotate the inner group beneath its nonuniformly scaled parent.
  const doc = new DOMParser().parseFromString(
    await zip.file(path)!.async('string'),
    'application/xml',
  );
  const groups = doc.getElementsByTagNameNS('*', 'grpSpPr');
  groups[2].getElementsByTagNameNS('*', 'xfrm')[0].setAttribute('rot', '1800000');
  zip.file(path, new XMLSerializer().serializeToString(doc));
  const data = await zip.generateAsync({ type: 'arraybuffer' });
  const file = await importFile(
    Object.assign(new File([data], 'Skew.pptx'), { arrayBuffer: async () => data }),
  );
  expect((file.content as DeckContent).slides[0].elements[0].sourceGroupTransform?.projected).toBe(
    false,
  );
  expect(file.warnings.some((w) => w.includes('skewed transforms'))).toBe(true);
  expect(await exported(file)).toEqual(data);
});
it('does not clamp tiny child-coordinate units before applying the parent scale', async () => {
  const zip = await JSZip.loadAsync(input());
  const path = 'ppt/slides/slide1.xml';
  const doc = new DOMParser().parseFromString(
    await zip.file(path)!.async('string'),
    'application/xml',
  );
  const inner = doc.getElementsByTagNameNS('*', 'grpSp')[1];
  // A change of coordinate units leaves the native-measured effective geometry invariant.
  const own = inner.getElementsByTagNameNS('*', 'xfrm')[0];
  for (const tag of ['chOff', 'chExt']) {
    const node = own.getElementsByTagNameNS('*', tag)[0];
    for (const attr of Array.from(node.attributes))
      node.setAttribute(attr.name, String(Number(attr.value) / 12700));
  }
  for (const shape of Array.from(inner.getElementsByTagNameNS('*', 'sp'))) {
    const xfrm = shape.getElementsByTagNameNS('*', 'xfrm')[0];
    for (const node of Array.from(xfrm.children))
      for (const attr of Array.from(node.attributes))
        node.setAttribute(attr.name, String(Number(attr.value) / 12700));
  }
  zip.file(path, new XMLSerializer().serializeToString(doc));
  const data = await zip.generateAsync({ type: 'arraybuffer' });
  const deck = await readPresentation(data);
  for (const [i, el] of deck.slides[0].elements.entries())
    for (const key of ['x', 'y', 'w', 'h'] as const)
      expect(el[key]).toBeCloseTo((native.effective[0][i][key] * 4) / 3, 4);
  expect(contentSchema.safeParse(deck).success).toBe(true);
  const file = await importFile(
    Object.assign(new File([data], 'Coarse.pptx'), { arrayBuffer: async () => data }),
  );
  (file.content as DeckContent).slides[0].elements[0].x += 1;
  await expect(exported(file)).rejects.toThrow('coarse coordinates');
  expect(file.original!.data).toEqual(data);
});
it('inverts nested axis-aligned group transforms for child moves and resizes without modifying other parts', async () => {
  const file = await imported();
  const element = (file.content as DeckContent).slides[0].elements[0];
  element.x += 16;
  element.y += 12;
  element.w += 16;
  element.h += 12;
  const output = await exported(file);
  const actual = (await readPresentation(output)).slides[0].elements[0];
  for (const key of ['x', 'y', 'w', 'h'] as const) expect(actual[key]).toBeCloseTo(element[key], 4);
  const before = await JSZip.loadAsync(input()),
    after = await JSZip.loadAsync(output);
  const withoutEditedTransform = (xml: string) => {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const leaf = doc.getElementsByTagNameNS('*', 'sp')[0];
    leaf.getElementsByTagNameNS('*', 'xfrm')[0].remove();
    for (let group = leaf.parentElement; group?.localName === 'grpSp'; group = group.parentElement)
      Array.from(group.children)
        .find((el) => el.localName === 'grpSpPr')!
        .getElementsByTagNameNS('*', 'xfrm')[0]
        .remove();
    return doc.documentElement;
  };
  expect(
    withoutEditedTransform(await before.file('ppt/slides/slide1.xml')!.async('string')).isEqualNode(
      withoutEditedTransform(await after.file('ppt/slides/slide1.xml')!.async('string')),
    ),
  ).toBe(true);
  for (const path of Object.keys(before.files))
    if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
});

it('exports the independently measured rotated/reflected child frames and preserves every other package part', async () => {
  const file = await imported();
  expect(edits.sourceSha256).toBe(native.sourceSha256);
  for (const i of [2, 3]) {
    const el = (file.content as DeckContent).slides[i].elements[0];
    const expected = edits.effective[i][0];
    for (const key of ['x', 'y', 'w', 'h'] as const) el[key] = (expected[key] * 4) / 3;
  }
  const output = await exported(file),
    deck = await readPresentation(output);
  for (const [i, slide] of deck.slides.entries())
    for (const expected of edits.effective[i]) {
      const el = slide.elements.find((e) => e.sourceShapeId === String(expected.id))!;
      for (const key of ['x', 'y', 'w', 'h'] as const)
        expect(el[key]).toBeCloseTo((expected[key] * 4) / 3, 3);
      expect(el.rotation || 0).toBeCloseTo(expected.rotation, 3);
    }
  const before = await JSZip.loadAsync(input()),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files)) {
    if (before.files[path].dir) continue;
    if (['ppt/slides/slide3.xml', 'ppt/slides/slide4.xml'].includes(path)) {
      const strip = (xml: string) => {
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const leaf = doc.getElementsByTagNameNS('*', 'sp')[0];
        leaf.getElementsByTagNameNS('*', 'xfrm')[0].remove();
        for (
          let group = leaf.parentElement;
          group?.localName === 'grpSp';
          group = group.parentElement
        )
          Array.from(group.children)
            .find((el) => el.localName === 'grpSpPr')!
            .getElementsByTagNameNS('*', 'xfrm')[0]
            .remove();
        return doc.documentElement;
      };
      expect(
        strip(await before.file(path)!.async('string')).isEqualNode(
          strip(await after.file(path)!.async('string')),
        ),
      ).toBe(true);
    } else
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  }
  expect(file.original!.data).toEqual(input());
});
it.each(['x', 'y', 'w', 'h'] as const)(
  'preserves coupled local coordinates when editing only world %s',
  async (key) => {
    for (const i of [2, 3]) {
      const file = await imported(),
        el = (file.content as DeckContent).slides[i].elements[0];
      el[key] += 8;
      const actual = (await readPresentation(await exported(file))).slides[i].elements[0];
      for (const field of ['x', 'y', 'w', 'h'] as const)
        expect(actual[field]).toBeCloseTo(el[field], 3);
    }
  },
);
it('rejects singular or forged group transforms and unsupported text/child rotation edits', async () => {
  const deck = await readPresentation(input()),
    before = deck.slides[2].elements[0],
    after = { ...before, x: before.x + 16 };
  const broken = {
    ...before,
    sourceGroupTransform: {
      ...before.sourceGroupTransform!,
      matrix: [1, 1, 1, 1, 0, 0] as [number, number, number, number, number, number],
    },
  };
  expect(() => localGroupGeometry(broken, after, 9144000, 5143500)).toThrow();
  expect(() => localGroupGeometry({ ...before, type: 'text' }, after, 9144000, 5143500)).toThrow();
  expect(() => localGroupGeometry(before, { ...after, rotation: 60 }, 9144000, 5143500)).toThrow();
});

it('normalizes all eight nested group frames to independent native bounds without shifting other leaves', async () => {
  const file = await imported();
  expect(bounds.sourceSha256).toBe(native.sourceSha256);
  for (const [i, slide] of (file.content as DeckContent).slides.entries()) {
    for (const key of ['x', 'y', 'w', 'h'] as const)
      slide.elements[0][key] = (bounds.effective[i][0][key] * 4) / 3;
  }
  const output = await exported(file),
    zip = await JSZip.loadAsync(output);
  const result = await readPresentation(output);
  for (let i = 0; i < 4; i++) {
    const doc = new DOMParser().parseFromString(
      await zip.file(`ppt/slides/slide${i + 1}.xml`)!.async('string'),
      'application/xml',
    );
    const groups = Array.from(doc.getElementsByTagNameNS('*', 'grpSp'));
    expect(groups).toHaveLength(bounds.transforms[i].length);
    for (const [j, group] of groups.entries()) {
      const expected = bounds.transforms[i][j];
      expect(group.getElementsByTagNameNS('*', 'cNvPr')[0].getAttribute('id')).toBe(expected.id);
      const transform = group.getElementsByTagNameNS('*', 'xfrm')[0];
      for (const [key, value] of Object.entries(expected)) {
        if (key === 'id') continue;
        const [tag, attr] = key.split('.');
        expect(
          Math.abs(
            Number(transform.getElementsByTagNameNS('*', tag)[0].getAttribute(attr)) -
              Number(value),
          ),
          `slide ${i + 1} group ${j} ${key}`,
        ).toBeLessThanOrEqual(12.7);
      }
    }
    for (const [j, expected] of bounds.effective[i].entries())
      for (const key of ['x', 'y', 'w', 'h'] as const)
        expect(result.slides[i].elements[j][key]).toBeCloseTo((expected[key] * 4) / 3, 3);
  }
  expect(file.original!.data).toEqual(input());
});
it('normalizes shared ancestors only after both children are edited', async () => {
  const file = await imported(),
    slide = (file.content as DeckContent).slides[2];
  slide.elements[0].x += 16;
  slide.elements[1].y += 20;
  const actual = (await readPresentation(await exported(file))).slides[2];
  for (const [i, expected] of slide.elements.entries())
    for (const key of ['x', 'y', 'w', 'h'] as const)
      expect(actual.elements[i][key]).toBeCloseTo(expected[key], 3);
});

it('rejects group-bound rounding amplified by coarse ancestor scales', () => {
  const doc = new DOMParser().parseFromString(
    `<spTree><grpSp><grpSpPr><xfrm><off x="0" y="0"/><ext cx="9144000" cy="5143500"/><chOff x="0" y="0"/><chExt cx="10" cy="10"/></xfrm></grpSpPr><grpSp><grpSpPr><xfrm><off x="0" y="0"/><ext cx="5" cy="5"/><chOff x="0" y="0"/><chExt cx="3" cy="3"/></xfrm></grpSpPr><sp><spPr><xfrm><off x="1" y="1"/><ext cx="2" cy="2"/></xfrm></spPr></sp></grpSp></grpSp></spTree>`,
    'application/xml',
  );
  expect(() => normalizeGroupBounds([doc.getElementsByTagName('sp')[0]], 9144000, 5143500)).toThrow(
    'coarse coordinates',
  );
});
it.each(['rotated', 'opaque'] as const)(
  'retains unmeasured %s direct-child group bounds',
  async (kind) => {
    const zip = await JSZip.loadAsync(input());
    const doc = new DOMParser().parseFromString(
      await zip.file('ppt/slides/slide1.xml')!.async('string'),
      'application/xml',
    );
    const leaf = doc.getElementsByTagNameNS('*', 'sp')[0];
    const sibling = doc.getElementsByTagNameNS('*', 'sp')[1];
    if (kind === 'rotated') {
      sibling.getElementsByTagNameNS('*', 'xfrm')[0].setAttribute('rot', '1800000');
      sibling.getElementsByTagNameNS('*', 'prstGeom')[0].setAttribute('prst', 'triangle');
    } else
      sibling.parentElement!.appendChild(
        doc.createElementNS(sibling.namespaceURI, 'p:graphicFrame'),
      );
    const original = new XMLSerializer().serializeToString(doc);
    normalizeGroupBounds([leaf], 9144000, 5143500);
    expect(new XMLSerializer().serializeToString(doc)).toBe(original);
  },
);
