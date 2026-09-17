import { Extension } from '@tiptap/core';
import { TextStyle } from '@tiptap/extension-text-style';

/** Kerning metadata can be the only attribute on a retained run. The ordinary
 * text-style parser accepts only style-bearing spans. */
export const WordTextStyle = TextStyle.extend({
  parseHTML() {
    return [
      ...(this.parent?.() || []),
      { tag: 'span[data-word-font-features]', consuming: false },
      {
        tag: 'span[data-word-kerning]',
        consuming: false,
        getAttrs: (el) =>
          wordKerningValue(el.getAttribute('data-word-kerning')) === null ? false : {},
      },
    ];
  },
});

/** Retain Word's minimum kerning size in half-points, including an explicit
 * zero override. The pinned Word treats zero as disabled. */
export function wordKerningValue(value: unknown): number | null {
  if (typeof value !== 'string' && typeof value !== 'number') return null;
  if (!/^\d{1,4}$/.test(String(value))) return null;
  const n = Number(value);
  return Number.isInteger(n) && n >= 0 && n <= 3276 ? n : null;
}

/** DOCX also permits explicit physical units. Keep the half-point precision
 * supported by Word font editing; reject unknown or finer measurements. */
export function wordKerningXmlValue(value: string | null): number | null {
  const integer = wordKerningValue(value);
  if (integer !== null) return integer;
  const match = /^(\d+(?:\.\d+)?)(pt|in|cm|mm|pc|pi)$/.exec(value || '');
  if (!match) return null;
  const factors: Record<string, number> = {
    pt: 2,
    in: 144,
    cm: 144 / 2.54,
    mm: 144 / 25.4,
    pc: 24,
    pi: 24,
  };
  const halfPoints = Number(match[1]) * factors[match[2]];
  if (!Number.isFinite(halfPoints) || Math.abs(halfPoints - Math.round(halfPoints)) > 1e-6)
    return null;
  return wordKerningValue(Math.round(halfPoints));
}

export function wordKerningStyle(threshold: number, fontSize: unknown): 'normal' | 'none' {
  if (typeof fontSize !== 'string') return 'none';
  const match = /^(\d+(?:\.\d+)?)(pt|px)$/.exec(fontSize);
  const halfPoints = match ? Number(match[1]) * (match[2] === 'px' ? 1.5 : 2) : 0;
  return threshold > 0 && halfPoints >= threshold ? 'normal' : 'none';
}

export const WordKerning = Extension.create({
  name: 'wordKerning',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          paragraphKerning: {
            default: null,
            parseHTML: (el) => wordKerningValue(el.getAttribute('data-word-paragraph-kerning')),
            renderHTML: (attrs) => {
              const value = wordKerningValue(attrs.paragraphKerning);
              return value === null ? {} : { 'data-word-paragraph-kerning': String(value) };
            },
          },
        },
      },
      {
        types: ['textStyle'],
        attributes: {
          wordKerning: {
            default: null,
            // Nested edit-session spans inherit the outer run's metadata just
            // as TextStyle merges its CSS. An explicit inner zero wins.
            parseHTML: (el) =>
              wordKerningValue(
                el.closest('[data-word-kerning]')?.getAttribute('data-word-kerning'),
              ),
            renderHTML: (attrs) => {
              const threshold = wordKerningValue(attrs.wordKerning);
              return threshold === null
                ? {}
                : {
                    'data-word-kerning': String(threshold),
                    style: `font-kerning:${wordKerningStyle(threshold, attrs.fontSize)}`,
                  };
            },
          },
        },
      },
    ];
  },
});
