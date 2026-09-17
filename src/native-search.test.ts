import { expect, it } from 'vitest';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import native from '../tests/fixtures/native-excel-search.json';
import { textMatches, replaceMatches } from './text-search';

it('binds the native Excel oracle to its exact authored input', () => {
  const bytes = fs.readFileSync('tests/fixtures/search-oracle-cases.json');
  expect(createHash('sha256').update(bytes).digest('hex')).toBe(native.inputSha256);
  const inputs = JSON.parse(bytes.toString());
  expect(native.application).toBe('Microsoft Excel');
  expect(native.cases).toHaveLength(inputs.length);
  native.cases.forEach((c, i) =>
    expect(c).toMatchObject({
      ...inputs[i],
      matchCase: !!inputs[i].matchCase,
      wholeCell: !!inputs[i].wholeCell,
    }),
  );
});
it.each(native.cases)('matches desktop Excel replacement: $text / $query', (c) => {
  const options = { wildcards: true, matchCase: c.matchCase, wholeCell: c.wholeCell };
  const matches = textMatches(c.text, c.query, { ...options, forReplacement: true });
  expect(replaceMatches(c.text, matches, c.replacement, 32767)).toBe(c.result);
  expect(textMatches(c.text, c.query, options).length > 0).toBe(c.found);
});
