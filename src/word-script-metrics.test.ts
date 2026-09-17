import { expect, it } from 'vitest';
import native from '../tests/fixtures/native-word-script-fonts.json';
import { wordScriptMetrics } from './word-script-metrics';
import single from '../tests/fixtures/native-word-script-single-lines.json';
import { wordNativeBaseline } from './word-native-baseline';
import { wordUniformLineHeight } from './word-font-line-metrics';

const face = { fontFamily: 'Arial', fontWeight: '400', fontStyle: 'normal', fontStretch: '100%' };
const spacing = { rule: 'exact' as const, line: 800 };
for (const row of native.rows) it(`matches native painted script size and baseline: ${row.name}`, () => {
  const result = wordScriptMetrics(row.script.toLowerCase() as 'superscript' | 'subscript', row.size / .75, row.size / .75, face, spacing)!;
  expect(result).not.toBeNull();
  expect(Math.abs(result.size * .75 - row.geometry.b.size)).toBeLessThan(.0001);
  expect(Math.abs(result.shift * .75 - row.geometry.baselineShift)).toBeLessThan(.0001);
});
it('rejects unqualified descriptors, paragraph sizes and line contexts', () => {
  for (const changed of [{fontFamily:'Calibri'}, {fontWeight:'700'}, {fontStyle:'italic'}, {fontStretch:'condensed'}])
    expect(wordScriptMetrics('superscript', 10/.75, 10/.75, {...face, ...changed}, spacing)).toBeNull();
  for (const size of [0, 9, 11, 12, 19.5, 20.5, NaN, Infinity])
    expect(wordScriptMetrics('subscript', size/.75, size/.75, face, spacing)).toBeNull();
  expect(wordScriptMetrics('subscript', 10/.75, 20/.75, face, spacing)).toBeNull();
  for (const changed of [{rule:'auto' as const,line:240}, {rule:'atLeast' as const,line:800}, {rule:'exact' as const,line:799}])
    expect(wordScriptMetrics('superscript', 10/.75, 10/.75, face, changed)).toBeNull();
});
for (const row of single.rows) it(`retains native Single leading and shifted glyphs: ${row.name}`, () => {
  const spacing = { rule: 'auto' as const, line: 240 }, size = 10 / .75, ratio = 1.1499;
  const height = wordUniformLineHeight(size, ratio, spacing);
  const baseline = wordNativeBaseline(size, size * ratio, ratio, 'Arial', '400', 'normal', '100%', spacing)!;
  const script = wordScriptMetrics(row.script.toLowerCase() as 'superscript' | 'subscript', size, size, face, spacing, true)!;
  expect(script).not.toBeNull(); expect(baseline).not.toBeNull();
  row.geometry.forEach((paragraph, index) => {
    expect(Math.abs(height * .75 - paragraph.height)).toBeLessThanOrEqual(.15);
    const characters = row.stages[0].paragraphs[index].characters.filter(c => c.text !== '\r');
    expect(paragraph.glyphs.map(g => g.text)).toEqual(characters.map(c => c.text));
    characters.forEach((character, i) => {
      const scripted = character.superscript === -1 || character.subscript === -1;
      expect(Math.abs((baseline + (scripted ? script.shift : 0)) * .75 - paragraph.offsets[i])).toBeLessThanOrEqual(.15);
      expect(Math.abs((scripted ? script.size : size) * .75 - paragraph.glyphs[i].size)).toBeLessThanOrEqual(.15);
    });
  });
});
it('keeps unmeasured Single-script sizes and other automatic line multiples guarded', () => {
  expect(wordScriptMetrics('superscript', 20 / .75, 20 / .75, face, { rule: 'auto', line: 240 }, true)).toBeNull();
  for (const line of [239, 241, 360, 480])
    expect(wordScriptMetrics('subscript', 10 / .75, 10 / .75, face, { rule: 'auto', line }, true)).toBeNull();
});
