import { it, expect } from 'vitest';
import native from '../tests/fixtures/native-word-tab-leader-fonts.json';
import { wordTabLeaderFont } from './word-tab-leader-font';
import type { WordTabStop } from './word-tab-stops';
import storyRanges from '../tests/fixtures/native-word-story-font-range-geometry.json';
import fontSteps from '../tests/fixtures/native-word-story-font-step-geometry.json';
import scripts from '../tests/fixtures/native-word-script-leaders.json';

it('matches the eight native Arial10 script dot profiles without qualifying unmeasured scripts or leaders', () => {
  for (const row of scripts.rows) {
    const font = { family: 'Arial', size: 6.48, weight: '400', style: 'normal', stretch: 'normal',
      letterSpacing: 'normal', normalRatio: 1.1499, script: row.script as 'superscript' | 'subscript' };
    const actual = wordTabLeaderFont('dot', font); expect(actual).not.toBeNull();
    for (const pitch of row.pitches) expect(Math.abs(actual!.pitch - pitch)).toBeLessThanOrEqual(.15);
    for (const leader of ['hyphen', 'underscore', 'heavy', 'middleDot'] as const)
      expect(wordTabLeaderFont(leader, font)).toBeNull();
    expect(wordTabLeaderFont('dot', { ...font, size: 13 })).toBeNull();
    expect(wordTabLeaderFont('dot', { ...font, script: undefined })).toBeNull();
  }
});

it('matches measured Arial9/11 dot pitches without extrapolating other leader kinds', () => {
  for (const size of [9, 11]) {
    const font = { family: 'Arial', size, weight: '400', style: 'normal', stretch: 'normal', letterSpacing: 'normal', normalRatio: 1.1499 };
    const actual = wordTabLeaderFont('dot', font);expect(actual).not.toBeNull();
    const groups = fontSteps.rows.flatMap(row => row.leaders).filter(group => Math.abs(group.first.size - size) < .15);
    expect(groups).toHaveLength(4);
    for (const group of groups) for (const pitch of group.pitches)
      expect(Math.abs(actual!.pitch - pitch)).toBeLessThanOrEqual(.15);
    for (const leader of ['hyphen', 'underscore', 'heavy', 'middleDot'] as const)
      expect(wordTabLeaderFont(leader, font)).toBeNull();
  }
});

it('matches independently formatted Times20 story dots without extrapolating other leaders', () => {
  const font = { family: 'Times New Roman', size: 20, weight: '400', style: 'normal', stretch: 'normal', letterSpacing: 'normal', normalRatio: 1.1499 };
  const actual = wordTabLeaderFont('dot', font);
  expect(actual).not.toBeNull();
  const groups = storyRanges.rows.flatMap(row => row.leaders).filter(group => group.first.font === 'TimesNewRomanPSMT');
  expect(groups).toHaveLength(2);
  for (const group of groups) for (const pitch of group.pitches)
    expect(Math.abs(actual!.pitch - pitch)).toBeLessThanOrEqual(.15);
  for (const leader of ['hyphen', 'underscore', 'heavy', 'middleDot'] as const)
    expect(wordTabLeaderFont(leader, font)).toBeNull();
});

for (const [index, row] of native.rows.entries()) it(`native leader font cell ${index}: ${row.family} ${row.size} ${row.weight} ${row.style}`, () => {
  const leaders: WordTabStop['leader'][] = row.glyph === '_' ? ['underscore', 'heavy']
    : [{ '.': 'dot', '-': 'hyphen', '\u00b7': 'middleDot' }[row.glyph] as WordTabStop['leader']];
  for (const leader of leaders) expect(wordTabLeaderFont(leader, { ...row, stretch: 'normal', letterSpacing: 'normal' }))
    .toEqual({ glyph: row.glyph, pitch: row.pitch });
});
it('does not extrapolate leader cells to unknown fonts, styles, sizes or metric substitutions', () => {
  const font = { family: 'Arial', size: 10, weight: '400', style: 'normal', stretch: 'normal', letterSpacing: 'normal', normalRatio: 1.1499 };
  expect(wordTabLeaderFont('none', font)).toBeNull();
  for (const changed of [{ family: 'Unverified' }, { size: 12 }, { size: NaN }, { normalRatio: 1.2 }, { normalRatio: NaN },
    { weight: '700', style: 'italic' }, { stretch: 'condensed' }, { letterSpacing: '.1px' }])
    expect(wordTabLeaderFont('dot', { ...font, ...changed })).toBeNull();
});
