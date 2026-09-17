import { expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { getSchema } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { newFile } from './model';
import { readDocxStructure } from './docx-sections';
import { wordXml, descendants, WORD_NS, val } from './word-xml';
import { readWordStorySources } from './word-story-source';
import { resolveWordSections } from './word-section-layout';
import { mapWordSectionRanges } from './word-section-ranges';
import { sectionSurfaceInput } from './word-section-surfaces';
import { planWordFragments } from './word-fragment-plan';
import type { WordStoryMeasurements } from './word-story-layout';
import reference from '../tests/fixtures/native-word-section-stories.json';

for (const [name, sample] of Object.entries(reference.cases))
  it(`allocates the native ${name} story transitions, widths and blank pages`, async () => {
    const bytes = readFileSync(`tests/fixtures/${sample.sourceFile}`);
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(sample.sourceSha256);
    const zip = await JSZip.loadAsync(bytes);
    const main = wordXml(await zip.file('word/document.xml')!.async('string'));
    const structure = readDocxStructure(main),
      sources = await readWordStorySources(zip, structure);
    const file = newFile('word');
    if (file.content.kind !== 'word') throw Error('Expected Word');
    file.content.docxStructure = structure;
    file.content.margin = 'narrow';
    file.content.orientation =
      val(descendants(main, 'pgSz').at(-1), 'orient') === 'landscape' ? 'landscape' : 'portrait';
    const sections = resolveWordSections(file.content);
    const parts = new Map<string, WordStoryMeasurements['parts'][number]>();
    for (const [index, section] of sections.entries()) {
      const width = (section.width! - section.margins.left! - section.margins.right!) / 15;
      for (const kind of ['headers', 'footers'] as const)
        for (const slot of ['default', 'first', 'even'] as const) {
          if (
            (slot === 'first' && !section.differentFirstPage) ||
            (slot === 'even' && !sources.evenAndOddHeaders)
          )
            continue;
          const ref = section[kind][slot];
          if (!ref) continue;
          const source = sources.parts.find((p) => p.relationshipIds.includes(ref.relationshipId))!;
          const renderKey = JSON.stringify([source.path, width]);
          const firstPage = sample.paintPages.find((p) =>
            p.lines.some((l) => l.text.startsWith(`B${index + 1}-`)),
          )!;
          // Native source line geometry supplies the measured input. Browser
          // wrapping/measurement is verified separately against these same files.
          const lines =
            name === 'linked-width' && kind === 'headers'
              ? firstPage.lines.filter((l) => !l.text.startsWith('B') && !l.text.includes('footer'))
                  .length
              : 1;
          parts.set(renderKey, {
            path: source.path,
            kind: source.kind,
            relationshipIds: source.relationshipIds,
            width,
            renderKey,
            height: lines * 16,
          });
        }
    }
    const measurements: WordStoryMeasurements = {
      version: 1,
      evenAndOddHeaders: sources.evenAndOddHeaders,
      parts: [...parts.values()],
    };
    const lines = descendants(main, 'p').map((p) =>
      [...p.getElementsByTagNameNS(WORD_NS, '*')]
        .filter((e) => ['t', 'br'].includes(e.localName))
        .map((e) => (e.localName === 'br' ? '\n' : e.textContent))
        .join('')
        .split('\n'),
    );
    const doc = getSchema(wordExtensions()).nodeFromJSON({
      type: 'doc',
      content: lines.map((text, index) => ({
        type: 'paragraph',
        attrs: {
          sourceParagraph: `0:${index}`,
          paragraphLineHeight: '12pt',
          paragraphLineRule: 'exact',
          spaceBefore: '0pt',
          spaceAfter: '0pt',
          widowControl: false,
        },
        content: text.flatMap((text, i, all) => [
          { type: 'text', text },
          ...(i < all.length - 1 ? [{ type: 'hardBreak' }] : []),
        ]),
      })),
    });
    const input = sectionSurfaceInput(
      doc,
      mapWordSectionRanges(doc, structure),
      sections,
      measurements,
    )!;
    expect(input).not.toBeNull();
    const flows = lines.map((text) => {
      let offset = 0;
      return {
        height: text.length * 16,
        before: 0,
        after: 0,
        keepLines: false,
        widowControl: false,
        lines: text.map((text, i, all) => {
          const from = offset;
          offset += text.length + (i < all.length - 1 ? 1 : 0);
          return { from, to: offset, top: i * 16, height: 16 };
        }),
      };
    });
    const plan = planWordFragments(input, flows)!;
    expect(plan).not.toBeNull();
    expect(plan.pages).toHaveLength(sample.native.pages);
    for (const [index, page] of plan.pages.entries()) {
      const painted = sample.paintPages[index];
      const fragments = plan.fragments.filter((f) => f.page === index);
      const texts = fragments.flatMap((f) =>
        flows[f.block].lines.flatMap((line, i) =>
          line.from >= f.from && line.from < f.to ? [lines[f.block][i]] : [],
        ),
      );
      expect(texts).toEqual(painted.lines.filter((l) => l.text.startsWith('B')).map((l) => l.text));
      expect((page.stories || []).length).toBe(texts.length ? 2 : 0);
      expect(!!page.blank).toBe(!texts.length);
      for (const fragment of fragments) {
        const at = flows[fragment.block].lines.findIndex((l) => l.from === fragment.from);
        const native = painted.lines.find((l) => l.text === lines[fragment.block][at])!;
        expect(Math.abs((fragment.top - page.top) * 0.75 + 9.6 - native.y)).toBeLessThan(0.15);
        expect(Math.abs((fragment.left - page.left) * 0.75 - native.x)).toBeLessThan(0.15);
      }
      const actual = (page.stories || [])
        .map((s) =>
          descendants(sources.parts.find((p) => p.path === s.path)!.document, 't')
            .map((t) => t.textContent)
            .join(''),
        )
        .join(' ');
      const expected = painted.lines
        .filter((l) => !l.text.startsWith('B'))
        .map((l) => l.text)
        .join(' ');
      expect(actual.replace(/\s+/g, ' ').trim()).toBe(expected.replace(/\s+/g, ' ').trim());
    }
  });
