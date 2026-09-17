import { Extension } from '@tiptap/core';
import { wordFontFeaturesStyle, wordFontFeaturesValue } from './word-font-features';

export const WordFontFeatures = Extension.create({
  name: 'wordFontFeatures',
  addGlobalAttributes() {
    return [
      { types: ['paragraph', 'heading'], attributes: {
        paragraphFontFeatures: {
          default: null,
          parseHTML: (el) => wordFontFeaturesValue(el.getAttribute('data-word-paragraph-font-features')),
          renderHTML: (attrs) => {
            const value = wordFontFeaturesValue(attrs.paragraphFontFeatures);
            return value === null ? {} : { 'data-word-paragraph-font-features': String(value),
              style: `font-feature-settings:${wordFontFeaturesStyle(value)}` };
          },
        },
      } },
      { types: ['textStyle'], attributes: {
        wordFontFeatures: {
          default: null,
          parseHTML: (el) => wordFontFeaturesValue(el.closest('[data-word-font-features]')
            ?.getAttribute('data-word-font-features')),
          renderHTML: (attrs) => {
            const value = wordFontFeaturesValue(attrs.wordFontFeatures);
            return value === null ? {} : { 'data-word-font-features': String(value),
              style: `font-feature-settings:${wordFontFeaturesStyle(value)}` };
          },
        },
      } },
    ];
  },
});
