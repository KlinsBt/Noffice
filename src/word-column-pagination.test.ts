import { expect, it, vi } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { getSchema } from '@tiptap/core';
import { DOMParser as EditorDOMParser } from '@tiptap/pm/model';
import native from '../tests/fixtures/native-word-section-flow.json';
import oversized from '../tests/fixtures/native-word-keep-overflow.json';
import inline from '../tests/fixtures/native-word-inline-flow.json';
import forcedKeep from '../tests/fixtures/native-word-forced-keep.json';
import uiFlow from '../tests/fixtures/native-word-ui-flow.json';
import unequalFlow from '../tests/fixtures/native-word-unequal-flow.json';
import { paragraphGraphemes, type WordLine } from './word-line-measurements';
import { readDocx } from './docx-import';
import { wordExtensions } from './word-extensions';
import { resolveWordSections } from './word-section-layout';
import { mapWordSectionRanges } from './word-section-ranges';
import { sectionSurfaceInput } from './word-section-surfaces';
import { planWordFragments, type ParagraphFlow } from './word-fragment-plan';
import { reflowWordColumns, type WordLineProvider } from './word-column-reflow';

vi.mock('mammoth', async (original) => {
  const actual = await original<typeof import('mammoth')>();
  return {
    ...actual,
    convertToHtml: (input: { arrayBuffer: ArrayBuffer }, options: unknown) =>
      actual.convertToHtml(
        { buffer: Buffer.from(input.arrayBuffer) },
        options as Parameters<typeof actual.convertToHtml>[1],
      ),
  };
});

const cases = [
  ...Object.entries(unequalFlow.cases)
    .filter(([, reference]) => reference.configuration.flow !== 'auto')
    .map(([name, reference]) => ({ name, reference, prefix: 'word-unequal-flow' })),
  ...Object.entries(uiFlow.cases).map(([name, reference]) => ({
    name,
    reference,
    prefix: 'word-ui-flow',
  })),
  ...Object.entries(forcedKeep.cases).map(([name, reference]) => ({
    name,
    reference,
    prefix: 'word-forced-keep',
  })),
  ...Object.entries(inline.cases).map(([name, reference]) => ({
    name,
    reference,
    prefix: 'word-inline-flow',
  })),
  ...Object.entries(native.cases).map(([name, reference]) => ({
    name,
    reference,
    prefix: 'word-section-flow',
  })),
  ...Object.entries(oversized.cases).map(([name, reference]) => ({
    name,
    reference,
    prefix: 'word-keep-overflow',
  })),
];
for (const { name, reference, prefix } of cases)
  it(`retains and allocates the independently captured section flow: ${name}`, async () => {
    const source = readFileSync(`tests/fixtures/${prefix}-${name}.docx`);
    expect(createHash('sha256').update(source).digest('hex')).toBe(reference.sourceSha256);
    const { content } = await readDocx(Uint8Array.from(source).buffer);
    const body = new DOMParser().parseFromString(content.html, 'text/html').body;
    const doc = EditorDOMParser.fromSchema(getSchema(wordExtensions())).parse(body);
    const sections = resolveWordSections(content);
    const input = sectionSurfaceInput(
      doc,
      mapWordSectionRanges(doc, content.docxStructure!),
      sections,
    );
    if (name === 'unequal-columns') {
      expect(input).not.toBeNull();
      expect(sections[0].columns?.widths).toEqual([1480, 2960]);
    }
    expect(input).not.toBeNull();
    let metrics: ParagraphFlow[] = [];
    const providers = new Map<number, WordLineProvider>();
    doc.forEach((p) => {
      let from = 0;
      // This is the allocation oracle, separate from browser line measurement.
      // Uniform Arial 10pt Single is 11.5pt in the pinned font reference.
      const height =
        (prefix === 'word-ui-flow' && p.attrs.paragraphLineHeight === '1') || name.endsWith('-auto')
          ? 11.5 / 0.75
          : name.endsWith('-atLeast')
            ? 30 / 0.75
            : name.endsWith('-minimum')
              ? 18 / 0.75
              : 20;
      const text = paragraphGraphemes(p)!.text;
      let rows = [...text.matchAll(/[^\n\f\u000e]*(?:[\n\f\u000e]|$)/g)].map((m) => m[0]);
      if (text && !/[\n\f\u000e]$/.test(text)) rows.pop();
      if (input!.blocks[metrics.length].pageBreakParagraph && rows.at(-1) === '') rows.pop();
      if (prefix === 'word-unequal-flow' && name.endsWith('-wrapped')) {
        // Isolate allocation from the browser's font measurement. These native
        // wordNN fixtures fit two tokens at 80pt and four at 142pt. The provider
        // must receive each column's width and resume at the new text offset.
        rows = [...text.matchAll(/(?:\S+\s*){1,2}/g)].map((match) => match[0]);
        providers.set(metrics.length, (from, width) => {
          const count = width * 0.75 < 100 ? 2 : 4;
          const row = new RegExp(`^(?:\\S+\\s*){1,${count}}`).exec(text.slice(from))?.[0];
          return row ? { from, to: from + row.length, height } : null;
        });
      }
      const lines = rows.map((row, index) => {
        const to = from + row.length;
        const line: WordLine = { from, to, top: index * height, height };
        if (row.endsWith('\f')) line.breakAfter = 'page';
        if (row.endsWith('\u000e')) line.breakAfter = 'column';
        from = to;
        return line;
      });
      metrics.push({
        lines,
        height: lines.length * height,
        before: 0,
        after: 0,
        keepLines: p.attrs.keepLines === true,
        widowControl: p.attrs.widowControl !== false,
      });
    });
    const reflow = providers.size ? reflowWordColumns(input!, metrics, providers) : null;
    if (providers.size) {
      expect(reflow).not.toBeNull();
      metrics = reflow!.flows;
    }
    const plan = reflow?.plan ?? planWordFragments(input!, metrics)!;
    expect(plan).not.toBeNull();
    if (prefix === 'word-unequal-flow' && 'pdfPages' in reference.native) {
      const rules = plan.separators || [];
      const paths = reference.native.pdfPages.flatMap((page, index) =>
        ('paths' in page ? page.paths : []).map((path) => ({ page: index, ...path })),
      );
      expect(rules).toHaveLength(paths.length);
      for (let i = 0; i < paths.length; i++) {
        const rule = rules[i],
          path = paths[i],
          page = plan.pages[rule.page];
        expect(rule.page).toBe(path.page);
        expect(
          Math.abs((rule.left - page.left) * 0.75 - (path.left + path.right) / 2),
        ).toBeLessThanOrEqual(0.15);
        expect(Math.abs((rule.top - page.top) * 0.75 - path.top)).toBeLessThanOrEqual(0.15);
        expect(
          Math.abs((rule.top + rule.height - page.top) * 0.75 - path.bottom),
        ).toBeLessThanOrEqual(0.15);
      }
    }
    if ('pdfPages' in reference.native)
      expect(
        plan.pages.map((page) => ({
          width: page.width * 0.75,
          height: page.height * 0.75,
          blank: !!page.blank,
        })),
      ).toEqual(
        reference.native.pdfPages.map((page) => ({
          width: page.width,
          height: page.height,
          blank: !page.text,
        })),
      );
    for (const [block, p] of reference.native.paragraphs.entries()) {
      for (const [offset, character] of p.characters.entries()) {
        const matches = plan.fragments.filter(
          (f) => f.block === block && offset >= f.from && offset < f.to,
        );
        expect(matches, `${block}:${offset}`).toHaveLength(1);
        const f = matches[0],
          page = plan.pages[f.page];
        expect(f.page + 1, `${block}:${offset}`).toBe(character.page);
        const line = metrics[block].lines.find((l) => offset >= l.from && offset < l.to)!;
        const firstLine = metrics[block].lines.find((l) => l.from === f.from)!;
        const top = (f.top - page.top + line.top - firstLine.top) * 0.75;
        if ('paragraphBoxes' in reference) {
          // Word's COM glyph top differs from the actual paragraph box for
          // some automatic lines. The independently shaded PDF supplies the
          // comparable layout coordinate, with unchanged glyph matrices.
          expect(reference.paragraphBoxes.glyphsUnchanged).toBe(true);
          const boxes = reference.paragraphBoxes.boxes.filter(
            (b) =>
              b.paragraph === block &&
              b.page === character.page &&
              character.x >= b.left &&
              character.x <= b.right,
          );
          const box = boxes.sort(
            (a, b) => Math.abs(a.top - character.y) - Math.abs(b.top - character.y),
          )[0];
          expect(box).toBeDefined();
          expect(Math.abs(top - box.top)).toBeLessThanOrEqual(0.15);
        } else expect(top).toBeCloseTo(character.y, 2);
        if (offset === line.from) {
          const x = (f.left - page.left) * 0.75;
          // New unequal-column captures expose COM's printer-coordinate
          // quantization (127.9 for a geometric 128pt column). Their actual
          // native PDF glyph origins are checked separately at the established
          // 0.15pt bound. Retain the earlier fixtures' exact geometry checks.
          if (prefix === 'word-unequal-flow')
            expect(Math.abs(x - character.x)).toBeLessThanOrEqual(0.15);
          else expect(x).toBeCloseTo(character.x, 2);
        }
      }
    }
    expect(plan.pages).toHaveLength(
      Math.max(
        ...reference.native.paragraphs.flatMap((p) => [
          ...p.characters.map((c) => c.page),
          ...('caret' in p ? [(p.caret as { page: number }).page] : []),
        ]),
      ),
    );
    if (prefix === 'word-inline-flow') {
      const caret = (
        reference.native
          .paragraphs[0] as (typeof inline.cases)['terminal-page']['native']['paragraphs'][0]
      ).caret;
      const last = plan.fragments.at(-1)!;
      expect(last.page + 1).toBe(caret.page);
      expect((last.top - plan.pages[last.page].top) * 0.75).toBeCloseTo(caret.y, 2);
      if (last.from === last.to)
        expect((last.left - plan.pages[last.page].left) * 0.75).toBeCloseTo(caret.x, 2);
    }
  });
