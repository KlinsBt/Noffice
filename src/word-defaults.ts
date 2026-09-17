import { defaultFont } from './fonts';

/** The application's existing Normal paragraph defaults, in semantic units.
 * These must travel with newly authored content and its OOXML defaults. */
export const wordDefaults = {
  fontFamily: defaultFont,
  fontSize: 12,
  lineMultiple: 1.65,
  spaceAfter: 12,
} as const;

export const emptyWordParagraph = `<p data-word-paragraph-font-features="0" style="font-family:${wordDefaults.fontFamily};font-size:${wordDefaults.fontSize}pt;line-height:${wordDefaults.lineMultiple};margin-top:0pt;margin-bottom:${wordDefaults.spaceAfter}pt"></p>`;
