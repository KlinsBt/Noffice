import { describe, expect, it } from 'vitest';
import { docxStyles } from './docx-styles';
import { child, val, wordXml, WORD_NS, descendants } from './docx-import';
const xml = `<w:styles xmlns:w="${WORD_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Cambria"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:b/></w:rPr><w:pPr><w:spacing w:before="80" w:after="120"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Derived"><w:basedOn w:val="Base"/><w:rPr><w:sz w:val="32"/></w:rPr><w:pPr><w:spacing w:after="240"/></w:pPr></w:style><w:style w:type="character" w:styleId="Emphasis"><w:rPr><w:i/></w:rPr></w:style></w:styles>`;
describe('Word style inheritance for browser rendering', () => {
  it('resolves a paragraph mark independently from paragraph, character and direct text fonts', () => {
    const styles = `<w:styles xmlns:w="${WORD_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="true" w:styleId="Normal"><w:rPr><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="character" w:styleId="Accent"><w:rPr><w:sz w:val="36"/></w:rPr></w:style></w:styles>`;
    const doc = wordXml(
      `<w:p xmlns:w="${WORD_NS}"><w:pPr><w:rPr><w:sz w:val="48"/></w:rPr></w:pPr><w:r><w:t>Inherited</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Accent"/></w:rPr><w:t>Character</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Accent"/><w:sz w:val="40"/></w:rPr><w:t>Direct</w:t></w:r></w:p>`,
    );
    const pr = docxStyles(styles)(doc.documentElement);
    expect(val(child(child(pr, 'rPr')!, 'sz'))).toBe('48');
    expect(descendants(doc, 'r').map((r) => val(child(child(r, 'rPr')!, 'sz')))).toEqual([
      '28',
      '36',
      '40',
    ]);
  });

  it('resolves defaults, paragraph ancestry, character styles and direct overrides', () => {
    const doc = wordXml(
      `<w:p xmlns:w="${WORD_NS}"><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:sz w:val="40"/></w:rPr><w:t>Text</w:t></w:r></w:p>`,
    );
    docxStyles(xml)(doc.documentElement);
    const run = child(descendants(doc, 'r')[0], 'rPr')!,
      pr = descendants(doc, 'pPr')[0];
    expect(val(child(run, 'rFonts'), 'ascii')).toBe('Cambria');
    expect(val(child(run, 'sz'))).toBe('40');
    expect(child(run, 'b')).toBeDefined();
    expect(child(run, 'i')).toBeDefined();
    expect(val(child(pr, 'spacing'), 'before')).toBe('80');
    expect(val(child(pr, 'spacing'), 'after')).toBe('240');
  });
  it('resolves Latin theme font references', () => {
    const doc = wordXml(
      `<w:p xmlns:w="${WORD_NS}"><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr><w:t>Theme heading</w:t></w:r></w:p>`,
    );
    const theme =
      '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Georgia"/></a:majorFont></a:fontScheme></a:themeElements></a:theme>';
    docxStyles(xml, theme)(doc.documentElement);
    expect(val(child(child(descendants(doc, 'r')[0], 'rPr')!, 'rFonts'), 'ascii')).toBe('Georgia');
  });
  it('bounds circular basedOn chains', () => {
    const doc = wordXml(
      `<w:p xmlns:w="${WORD_NS}"><w:pPr><w:pStyle w:val="Loop"/></w:pPr><w:r><w:t>Safe</w:t></w:r></w:p>`,
    );
    expect(() =>
      docxStyles(
        `<w:styles xmlns:w="${WORD_NS}"><w:style w:styleId="Loop"><w:basedOn w:val="Loop"/></w:style></w:styles>`,
      )(doc.documentElement),
    ).not.toThrow();
    expect(descendants(doc, 't')[0].textContent).toBe('Safe');
  });
  it('applies style toggles and absolute direct run overrides', () => {
    const styles = `<w:styles xmlns:w="${WORD_NS}"><w:style w:styleId="Base"><w:rPr><w:b/></w:rPr></w:style><w:style w:styleId="Derived"><w:basedOn w:val="Base"/><w:rPr><w:b/></w:rPr></w:style></w:styles>`;
    const doc = wordXml(
      `<w:p xmlns:w="${WORD_NS}"><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:t>Off</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>On</w:t></w:r></w:p>`,
    );
    docxStyles(styles)(doc.documentElement);
    const bold = descendants(doc, 'r').map((r) => child(child(r, 'rPr')!, 'b'));
    expect(val(bold[0])).toBe('0');
    expect(val(bold[1])).toBe('');
  });
});
