import { expect, it } from 'vitest';
import { wordStoryGridEligible, wordStoryGridHtml, wordStoryGridOrigin } from './word-story-grid';
import reference from '../tests/fixtures/native-word-story-spacing.json';
import grid from '../tests/fixtures/native-word-story-spacing-grid.json';
import varied from '../tests/fixtures/native-word-story-line-grid.json';

it('matches eighty native positions when zero-gap story paragraphs have different physical line heights', () => {
  for (const row of varied.rows)
    expect(Math.abs(wordStoryGridOrigin(row.origin, 0) * .75 + row.baseline - row.expected), `${row.name}/${row.page}/${row.paragraph}`).toBeLessThanOrEqual(.15);
});

it('qualifies measured regular Arial10 single lines while retaining font/wrap/baseline guards', () => {
  const p = document.createElement('p');
  p.style.cssText = 'font-family:Arial;font-size:13.3333333333px;font-weight:400;font-style:normal;font-stretch:normal;line-height:2;margin:0;background-color:transparent';
  const run = document.createElement('span');
  // jsdom does not inherit computed font properties onto this inline child.
  run.style.cssText = 'font-family:Arial;font-size:13.3333333333px;font-weight:400;font-style:normal;font-stretch:normal';
  run.dataset.wordNativeBaseline = String(9.35 / .75); run.textContent = 'A'; p.append(run); document.body.append(p);
  try {
    expect(wordStoryGridEligible(p, 1)).toBe(false);
    expect(wordStoryGridEligible(p, 1, true)).toBe(true);
    expect(wordStoryGridEligible(p, 2, true)).toBe(false);
    run.style.fontWeight = '700'; expect(wordStoryGridEligible(p, 1, true)).toBe(false); run.style.fontWeight = '400';
    run.dataset.wordNativeBaseline = 'NaN'; expect(wordStoryGridEligible(p, 1, true)).toBe(false);
  } finally { p.remove(); }
});

function painted(kind: string, paragraphs: { before: number; after: number }[]) {
  const height = paragraphs.reduce((n, p, i) => n + 40 + Math.max(p.before, paragraphs[i - 1]?.after || 0), 0)
    + paragraphs.at(-1)!.after;
  let top = kind === 'header' ? 36 : 841.9 - 36 - height;
  return paragraphs.map((p, i) => {
    top += Math.max(p.before, paragraphs[i - 1]?.after || 0);
    const baseline = wordStoryGridOrigin(top / .75, p.after / .75) * .75 + 32.05;
    top += 40;
    return baseline;
  });
}

it('matches all560 independently captured baseline positions in the original56 spacing cases', () => {
  for (const row of reference.rows) for (const kind of ['header', 'footer'] as const) {
    const paragraphs = kind === row.kind ? row.paragraphs : Array.from({ length: 5 }, () => ({ before: 0, after: 0 }));
    const actual = painted(kind, paragraphs);
    for (let i = 0; i < actual.length; i++)
      expect(Math.abs(actual[i] - row.origins[kind][i]), `${row.name}/${kind}/${i}`).toBeLessThanOrEqual(.15);
  }
});

it('holds on180 native metafile baselines at independent fractional spacing boundaries', () => {
  for (const row of grid.rows) {
    const actual = painted(row.kind, row.paragraphs);
    actual.forEach((value, i) => expect(Math.abs(value - row.baselines[i]), `${row.name}/${i}`).toBeLessThanOrEqual(.15));
  }
});

it('keeps unsupported inputs unchanged and confines paint transforms to marked ephemeral stories', () => {
  for (const after of [-1, NaN, Infinity, 54]) expect(wordStoryGridOrigin(100, after)).toBe(100);
  const plain = '<p>Unqualified text</p>';
  expect(wordStoryGridHtml(plain, 48)).toBe(plain);
  const html = '<p data-word-story-paragraph-top="106.6666666667" data-word-story-grid-after="53.3333333333">A\tB</p>';
  const output = wordStoryGridHtml(html, 48);
  const node = document.createElement('div'); node.innerHTML = output;
  expect(node.textContent).toBe('A\tB');
  expect(node.querySelector('p')!.style.translate).not.toBe('');
  expect(wordStoryGridHtml(html, 48)).toBe(output);
});
