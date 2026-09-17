import { describe, expect, it } from 'vitest';
import { validateWordPdfSnapshot, wordPdfDrawFontKey, type WordPdfSnapshot } from './word-pdf-model';
import { fragmentWordSurfaces } from './word-fragment-plan';

const sample = (): WordPdfSnapshot => ({
  content: 'bound revision',
  pages: [
    {
      width: 595.3,
      height: 841.9,
      glyphs: [{ text: 'A', x: 72, y: 81.6, size: 10, family: 'Arial', color: [0, 0, 0] }],
    },
  ],
});
describe('PDF input boundary', () => {
  it('admits source-qualified bullet paint without accepting arbitrary marker text', () => {
    const input=sample(), glyph=input.pages[0].glyphs[0];
    glyph.text='\u2022';expect(()=>validateWordPdfSnapshot(input)).toThrow();
    glyph.listMarker=true;expect(validateWordPdfSnapshot(input)).toEqual(['Arial']);
    for(const text of ['A','\u25cf','\u00ad','1.']) {
      glyph.text=text;expect(()=>validateWordPdfSnapshot(input)).toThrow();
    }
    for(const text of ['1','.']) {
      glyph.text=text;expect(validateWordPdfSnapshot(input)).toEqual(['Arial']);
    }
  });
  it('admits only qualified literal hyphens and separates their aliased font mappings', () => {
    const input = sample(), glyph = input.pages[0].glyphs[0];
    glyph.text = '\u00ad'; glyph.literalHyphen = true;
    expect(validateWordPdfSnapshot(input)).toEqual(['Arial']);
    expect(wordPdfDrawFontKey(glyph)).not.toBe(wordPdfDrawFontKey({ ...glyph, text: '-', literalHyphen: undefined }));
    for (const text of ['-', 'A', '\u001f']) {
      glyph.text = text; expect(() => validateWordPdfSnapshot(input)).toThrow();
    }
    glyph.text = '\u00ad'; delete glyph.literalHyphen;
    expect(() => validateWordPdfSnapshot(input)).toThrow();
  });
  it('retains printable Latin-1 while rejecting controls and conditional glyphs', () => {
    for (const text of ['\u00a1', '\u00ac', '\u00ae', '\u00b5', '\u00c9', '\u00df', '\u00ff']) {
      const input = sample(); input.pages[0].glyphs[0].text = text;
      expect(validateWordPdfSnapshot(input)).toEqual(['Arial']);
      expect(input.pages[0].glyphs[0].text).toBe(text);
    }
    for (const text of ['\x7f', '\u0085', '\u009f', '\u00ad', '\u0100', 'e\u0301', 'ab']) {
      const input = sample(); input.pages[0].glyphs[0].text = text;
      expect(() => validateWordPdfSnapshot(input)).toThrow();
    }
  });
  it('retains nonbreaking-space identity in the layout snapshot', () => {
    const input = sample(); input.pages[0].glyphs[0].text = '\u00a0';
    expect(validateWordPdfSnapshot(input)).toEqual(['Arial']);
    expect(input.pages[0].glyphs[0].text).toBe('\u00a0');
  });
  it('accepts bounded text paint metadata and rejects forged decoration values', () => {
    for (const decoration of [0, 1, 2, 3]) {
      const input = sample(); input.pages[0].glyphs[0].decoration = decoration;
      expect(validateWordPdfSnapshot(input)).toEqual(['Arial']);
    }
    for (const decoration of [NaN, Infinity, -1, 4, 1.5]) {
      const input = sample(); input.pages[0].glyphs[0].decoration = decoration;
      expect(() => validateWordPdfSnapshot(input)).toThrow();
    }
  });
  it('bounds font-feature variants independently of installed face count', () => {
    const document = sample(), glyph = document.pages[0].glyphs[0];
    document.pages[0].glyphs = Array.from({ length: 32 }, (_, features) => ({ ...glyph, features, run: 1 }));
    expect(validateWordPdfSnapshot(document)).toEqual(['Arial']);
    document.pages[0].glyphs.push({ ...glyph, family: 'Calibri', features: 1, run: 1 });
    expect(() => validateWordPdfSnapshot(document)).toThrow(/32 font-feature/);
    for (const invalid of [NaN, -1, 32, 1.5]) {
      const input = sample(); input.pages[0].glyphs[0].features = invalid;
      expect(() => validateWordPdfSnapshot(input)).toThrow();
    }
    for (const invalid of [NaN, -1, 50001, 1.5]) {
      const input = sample(); input.pages[0].glyphs[0].run = invalid;
      expect(() => validateWordPdfSnapshot(input)).toThrow();
    }
    const unboundedRun = sample(); unboundedRun.pages[0].glyphs[0].features = 1;
    expect(() => validateWordPdfSnapshot(unboundedRun)).toThrow();
  });
  it('keeps font faces distinct and validates middle-dot glyphs without accepting forged font keys', () => {
    const valid = sample();
    const first = valid.pages[0].glyphs[0];
    valid.pages[0].glyphs.push(...(['bold', 'italic', 'boldItalic'] as const).map((face) => ({ ...first, text: '\u00b7', face })));
    expect(validateWordPdfSnapshot(valid)).toEqual(['Arial', 'Arial\0bold', 'Arial\0italic', 'Arial\0boldItalic']);
    first.family = 'Arial\0bold';
    expect(() => validateWordPdfSnapshot(valid)).toThrow();
    first.family = 'Arial';
    (first as unknown as { face: string }).face = 'light';
    expect(() => validateWordPdfSnapshot(valid)).toThrow();
  });
  it('preserves physical points and rejects nonfinite, unbounded and unsupported content', () => {
    const valid = sample();
    expect(validateWordPdfSnapshot(valid)).toEqual(['Arial']);
    expect(valid.pages[0].width).toBe(595.3);
    for (const corrupt of [
      (s: WordPdfSnapshot) => {
        s.pages[0].width = Infinity;
      },
      (s: WordPdfSnapshot) => {
        s.pages[0].glyphs[0].x = NaN;
      },
      (s: WordPdfSnapshot) => {
        s.pages[0].glyphs[0].color[0] = -1;
      },
      (s: WordPdfSnapshot) => {
        s.pages[0].glyphs[0].text = '\u202e';
      },
      (s: WordPdfSnapshot) => {
        s.pages[0].glyphs = Array(50001).fill(s.pages[0].glyphs[0]);
      },
    ]) {
      const invalid = sample();
      corrupt(invalid);
      expect(() => validateWordPdfSnapshot(invalid)).toThrow();
    }
  });
  it('includes empty physical pages without requiring an invented font', () => {
    const blank = sample();
    blank.pages[0].glyphs = [];
    expect(validateWordPdfSnapshot(blank)).toEqual([]);
  });
  it('preserves paragraph margins and semantic extents in a short-page PDF layout', () => {
    const original = {
      width: 800,
      height: 1100,
      pages: [{ sectionId: 'body', left: 0, top: 0, width: 800, height: 1100 }],
      blocks: [{ from: 0, to: 7, section: 0, width: 600, left: 100, top: 88 }],
    };
    const snapshot = structuredClone(original);
    const fragments = fragmentWordSurfaces(original, [{ height: 16, before: 12, after: 8 }]);
    expect(fragments.fragments[0]).toMatchObject({
      from: 0,
      to: 5,
      top: 100,
      left: 100,
      height: 16,
      page: 0,
    });
    expect(original).toEqual(snapshot);
    expect(() => fragmentWordSurfaces(original, [])).toThrow('Incomplete');
  });
});
