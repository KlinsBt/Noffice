import { expect, it } from 'vitest';
import { getSchema } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { authoredWordSection, authoredWordSectionMap } from './word-authored-section';
import { sectionSurfaceInput } from './word-section-surfaces';
import { singleSectionPage } from './word-section-layout';
import type { WordContent } from './model';

const schema = getSchema(wordExtensions());
const content: WordContent = {
  kind: 'word',
  html: '<p>Created here</p>',
  paper: 'a4',
  margin: 'normal',
};

it('maps authored page settings and live paragraphs without adding source identities', () => {
  const before = JSON.stringify(content);
  const doc = schema.node('doc', null, [
    schema.node('paragraph', null, schema.text('Created here')),
  ]);
  const section = authoredWordSection(content)!;
  expect(singleSectionPage(content)).toEqual(section);
  expect(section).toMatchObject({
    width: 11906,
    height: 16838,
    margins: { left: 1440, top: 1440 },
  });
  const input = sectionSurfaceInput(doc, authoredWordSectionMap(doc, section), [section]);
  expect(input?.blocks).toEqual([{ from: 0, to: 14, section: 0, width: (11906 - 2880) / 15 }]);
  expect(doc.firstChild!.attrs.sourceParagraph).toBeNull();
  expect(JSON.stringify(content)).toBe(before);
  expect(
    authoredWordSection({ ...content, paper: 'letter', orientation: 'landscape', margin: 'wide' }),
  ).toMatchObject({ width: 15840, height: 12240, margins: { left: 2160, bottom: 2160 } });
});

it('does not treat incomplete imports or retained boundaries as an authored section', () => {
  expect(
    authoredWordSection({ ...content, html: '<p data-source-paragraph="0:0">Old import</p>' }),
  ).toBeNull();
  expect(
    authoredWordSection({
      ...content,
      docxStructure: { version: 1, sourcePath: 'word/document.xml', sections: [] },
    }),
  ).toBeNull();
  const section = authoredWordSection(content)!;
  const doc = schema.node('doc', null, [
    schema.node('paragraph', { sourceParagraph: '0:0' }, schema.text('Old import')),
  ]);
  expect(authoredWordSectionMap(doc, section)).toBeUndefined();
});
