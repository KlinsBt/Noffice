import { expect, it } from 'vitest';
import { WORD_NS, wordXml, writeWordText } from './word-xml';

it.each([
  ['plain text', false], ['two  spaces', false], ['', false],
  [' leading', true], ['trailing ', true], [' both ', true], ['\ttext', true],
  ['text\n', true], ['\rtext', true], ['\u00a0text\u00a0', false],
] as const)('writes native text-edge preservation for %j', (value, preserve) => {
  const doc = wordXml(`<w:t xmlns:w="${WORD_NS}" xml:space="preserve">old</w:t>`);
  writeWordText(doc.documentElement, value);
  expect(doc.documentElement.textContent).toBe(value);
  expect(doc.documentElement.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'space'))
    .toBe(preserve ? 'preserve' : null);
});
