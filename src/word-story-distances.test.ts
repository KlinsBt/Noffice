import { expect, it } from 'vitest';
import { wordStoryDistancePoints } from './word-story-options';
import reference from '../tests/fixtures/native-word-story-distances.json';
import baseline from '../docs/parity/baseline.json';

it('uses distance observations from the installed baseline Word build', () => {
  expect(reference.executableHash).toBe(
    baseline.applications.find((app) => app.executable === 'WINWORD.EXE')?.sha256,
  );
});

for (const sample of reference.cases)
  it(`matches native ${sample.kind} acceptance and twip rounding for ${sample.input} pt`, () => {
    const actual = wordStoryDistancePoints(String(sample.input));
    if (sample.accepted) expect(actual).toBe(Math.round(sample.observed * 20));
    else expect(actual).toBeNaN();
  });

it.each(['', ' ', 'NaN', 'Infinity', '1e2', '18pt', '18,5', '--1', '1.2.3', '0'.repeat(101)])(
  'rejects malformed distance %j without changing document state',
  (input) => expect(wordStoryDistancePoints(input)).toBeNaN(),
);

it('accepts surrounding whitespace and decimal signs and canonicalizes negative zero', () => {
  expect(wordStoryDistancePoints(' +.025 ')).toBe(1);
  expect(wordStoryDistancePoints('-0.01')).toBe(0);
  expect(Object.is(wordStoryDistancePoints('-0'), -0)).toBe(false);
});
