import { expect, it } from 'vitest';
import JSZip from 'jszip';
import { readFileSync } from 'node:fs';
import { getSchema } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { newFile } from './model';
import { readDocxStructure } from './docx-sections';
import { wordXml, descendants } from './word-xml';
import { readWordStorySources } from './word-story-source';
import { resolveWordSections } from './word-section-layout';
import { mapWordSectionRanges } from './word-section-ranges';
import { sectionSurfaceInput } from './word-section-surfaces';
import { planWordFragments } from './word-fragment-plan';
import { wordStoryPageBounds, type WordStoryMeasurements } from './word-story-layout';
import reference from '../tests/fixtures/native-word-side-stories.json';

for (const [name, sample] of Object.entries(reference.cases))
  it(`allocates native ${name} body lines with the selected page's story heights`, async () => {
    const zip = await JSZip.loadAsync(readFileSync(`tests/fixtures/${sample.sourceFile}`));
    const structure = readDocxStructure(
      wordXml(await zip.file('word/document.xml')!.async('string')),
    );
    const source = await readWordStorySources(zip, structure);
    // The independent native fixture fixes every story line at exact12pt.
    // Browser measurement is separately required before enabling runtime use.
    const measurements: WordStoryMeasurements = {
      version: 1,
      evenAndOddHeaders: source.evenAndOddHeaders,
      parts: source.parts.map(({ document, ...part }) => ({
        ...part,
        height: (descendants(document, 'p').length + descendants(document, 'br').length) * 16,
      })),
    };
    const file = newFile('word');
    if (file.content.kind !== 'word') throw Error('Expected Word');
    file.content.docxStructure = structure;
    file.content.margin = 'narrow';
    file.content.orientation = 'landscape';
    const sections = resolveWordSections(file.content);
    const text = Array.from({ length: 40 }, (_, i) => `body${String(i + 1).padStart(2, '0')}`);
    const doc = getSchema(wordExtensions()).nodeFromJSON({
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          attrs: { sourceParagraph: '0:0' },
          content: text.flatMap((text, i) => [
            { type: 'text', text },
            ...(i < 39 ? [{ type: 'hardBreak' }] : []),
          ]),
        },
      ],
    });
    const input = sectionSurfaceInput(
      doc,
      mapWordSectionRanges(doc, structure),
      sections,
      measurements,
    )!;
    expect(input).not.toBeNull();
    const plan = planWordFragments(input, [
      {
        height: 640,
        before: 0,
        after: 0,
        keepLines: false,
        widowControl: false,
        lines: text.map((_, i) => ({
          from: i * 7,
          to: i * 7 + (i === 39 ? 6 : 7),
          top: i * 16,
          height: 16,
        })),
      },
    ])!;
    expect(plan).not.toBeNull();
    expect(plan.pages).toHaveLength(sample.native.pages);
    expect(plan.fragments.map((f) => Math.ceil((f.to - f.from) / 7))).toEqual(
      sample.paintPages.map((p) => p.lines.filter((l) => l.text.startsWith('body')).length),
    );
    for (const [index, page] of plan.pages.entries()) {
      const fragment = plan.fragments.find((f) => f.page === index)!;
      const firstLine = sample.paintPages[index].lines.find((l) => l.text.startsWith('body'))!;
      expect(Math.abs((fragment.top - page.top) * 0.75 + 9.6 - firstLine.y)).toBeLessThan(0.15);
      expect(page.stories).toHaveLength(2);
    }
    const bad = {
      ...measurements,
      parts: measurements.parts.map((part) => ({ ...part, height: Number.NaN })),
    };
    expect(wordStoryPageBounds(sections[0], bad, 1, true)).toBeNull();
    expect(wordStoryPageBounds(sections[0], measurements, 0, true)).toBeNull();
    // Equal widths do not imply equal leader phase: linked stories on pages
    // with shifted left/right margins must select their own measured context.
    const section = sections[0], left = section.margins.left! / 15;
    const width = (section.width! - section.margins.left! - section.margins.right!) / 15;
    const contextual = { ...measurements, parts: measurements.parts.flatMap((part) => [
      { ...part, width, pageLeft: left, renderKey: 'A/' + part.path },
      { ...part, width, pageLeft: left + 10, renderKey: 'B/' + part.path },
    ]) };
    expect(wordStoryPageBounds(section, contextual, 1, true)!.stories.every((story) => story.renderKey!.startsWith('A/'))).toBe(true);
    const shifted = { ...section, margins: { ...section.margins,
      left: section.margins.left! + 150, right: section.margins.right! - 150 } };
    expect(wordStoryPageBounds(shifted, contextual, 1, true)!.stories.every((story) => story.renderKey!.startsWith('B/'))).toBe(true);
  });
