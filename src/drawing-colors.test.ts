import { expect, it } from 'vitest';
import fs from 'node:fs';
import { createHash } from 'node:crypto';
import native from '../tests/fixtures/native-powerpoint-paints.json';
import inputs from '../tests/fixtures/paint-oracle-cases.json';
import { drawingColor, presentationPaints, paintCSS } from './drawing-colors';
const xml = (value: string) => new DOMParser().parseFromString(value, 'application/xml');
const color = (value: string) =>
  xml(`<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${value}</root>`)
    .documentElement.firstElementChild;
it('binds native paint results to exact authored inputs', () => {
  expect(
    createHash('sha256')
      .update(fs.readFileSync('tests/fixtures/paint-oracle-cases.json'))
      .digest('hex'),
  ).toBe(native.inputSha256);
  expect(native.cases.map((c) => c.name)).toEqual(inputs.map((c) => c.name));
});
it.each(native.cases)('matches native PowerPoint paint: $name', (c) => {
  const actual = drawingColor(color(inputs.find((i) => i.name === c.name)!.xml))!;
  // Native Office quantizes color transforms; tolerate at most one 8-bit channel step.
  for (const offset of [1, 3, 5])
    expect(
      Math.abs(
        parseInt(actual.color.slice(offset, offset + 2), 16) -
          parseInt(c.color.slice(offset, offset + 2), 16),
      ),
    ).toBeLessThanOrEqual(1);
  expect(actual.opacity).toBeCloseTo(c.opacity, 6);
});
it('resolves color-map overrides, style-matrix placeholders and background inheritance', () => {
  const theme = xml(
    '<theme><clrScheme><accent1><srgbClr val="345678"/></accent1><accent2><srgbClr val="987654"/></accent2><lt1><sysClr val="window" lastClr="ffffff"/></lt1></clrScheme><fmtScheme><fillStyleLst><solidFill><schemeClr val="phClr"/></solidFill></fillStyleLst><bgFillStyleLst><solidFill><schemeClr val="phClr"/></solidFill></bgFillStyleLst></fmtScheme></theme>',
  );
  const master = xml(
    '<sldMaster><cSld><bg><bgRef idx="1001"><schemeClr val="bg1"/></bgRef></bg></cSld><clrMap bg1="accent1"/></sldMaster>',
  );
  const layout = xml(
    '<sldLayout><clrMapOvr><overrideClrMapping bg1="accent2"/></clrMapOvr></sldLayout>',
  );
  const slide = xml('<sld/>');
  expect(presentationPaints(theme, master, layout, slide).background().color).toBe('#987654');
  expect(
    presentationPaints(
      theme,
      master,
      layout,
      xml('<sld><clrMapOvr><masterClrMapping/></clrMapOvr></sld>'),
    ).background().color,
  ).toBe('#345678');
  const resolver = presentationPaints(theme, master, layout, slide);
  expect(
    resolver.shape([
      xml('<sp><style><fillRef idx="1"><schemeClr val="accent1"/></fillRef></style></sp>')
        .documentElement,
    ]),
  ).toEqual({ color: '#345678', opacity: 1 });
  expect(
    resolver.shape([
      xml('<sp><spPr><noFill/><ln><solidFill><srgbClr val="FF0000"/></solidFill></ln></spPr></sp>')
        .documentElement,
    ]),
  ).toEqual({ color: 'transparent', opacity: 0 });
  expect(
    resolver.shape([xml('<sp><spPr><gradFill/></spPr></sp>').documentElement]),
  ).toBeUndefined();
  expect(presentationPaints(null, null, null, slide).background()).toEqual({
    color: '#ffffff',
    opacity: 1,
  });
});
it('rejects unresolved and unsupported colors rather than selecting unrelated descendant colors', () => {
  expect(drawingColor(color('<a:schemeClr val="unknown"/>'))).toBeUndefined();
  expect(drawingColor(color('<a:srgbClr val="red"/>'))).toBeUndefined();
  expect(drawingColor(color('<a:srgbClr val="123456"><a:unknown/></a:srgbClr>'))).toBeUndefined();
  expect(paintCSS('#123456', 0)).toBe('transparent');
  expect(paintCSS('#123456', 1)).toBe('#123456');
});
