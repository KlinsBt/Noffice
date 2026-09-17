import { Extension } from '@tiptap/core';
import { wordLineSpacing } from './word-line-spacing';
import { wordTabStops, wordDefaultTab, wordDecimalSymbol } from './word-tab-stops';
import { wordParagraphSpace } from './word-paragraph-spacing';

export const paragraphIndents = [
  ['indentStart', 'margin-inline-start', 'Indent before text'],
  ['indentEnd', 'margin-inline-end', 'Indent after text'],
  ['firstLineIndent', 'text-indent', 'First line (negative for hanging)'],
] as const;
export const paragraphBreaks = [
  ['keepNext', 'break-after', 'avoid', 'Keep with next'],
  ['keepLines', 'break-inside', 'avoid', 'Keep lines together'],
  ['pageBreakBefore', 'break-before', 'page', 'Page break before'],
  ['widowControl', 'orphans', '2', 'Widow/orphan control'],
] as const;

/** Paragraph layout is stored in semantic inline CSS for native/HTML portability. */
export const ParagraphLayout = Extension.create({
  name: 'paragraphLayout',
  addGlobalAttributes() {
    return [
      {
        types: ['paragraph', 'heading'],
        attributes: {
          paragraphStyleId: {
            default: null,
            parseHTML: (el: HTMLElement) => el.getAttribute('data-word-paragraph-style')?.slice(0, 253) || null,
            renderHTML: (attrs: Record<string, unknown>) => attrs.paragraphStyleId
              ? { 'data-word-paragraph-style': String(attrs.paragraphStyleId) } : {},
          },
          paragraphContextualSpacing: {
            default: null,
            parseHTML: (el: HTMLElement) => el.hasAttribute('data-word-contextual-spacing')
              ? el.getAttribute('data-word-contextual-spacing') === 'true' : null,
            renderHTML: (attrs: Record<string, unknown>) => attrs.paragraphContextualSpacing == null
              ? {} : { 'data-word-contextual-spacing': String(!!attrs.paragraphContextualSpacing) },
          },
          ...Object.fromEntries(['before', 'after'].map(side => [
            'paragraphSpace' + (side === 'before' ? 'Before' : 'After'), {
              default: null,
              parseHTML: (el: HTMLElement) => wordParagraphSpace(el.getAttribute('data-word-space-' + side)),
              renderHTML: (attrs: Record<string, unknown>) => {
                const value = wordParagraphSpace(attrs['paragraphSpace' + (side === 'before' ? 'Before' : 'After')]);
                return value ? { ['data-word-space-' + side]: JSON.stringify(value) } : {};
              },
            },
          ])),
          paragraphTabs: {
            default: null,
            parseHTML: (el: HTMLElement) => wordTabStops(el.getAttribute('data-word-tabs')),
            renderHTML: (attrs: Record<string, unknown>) => {
              const stops = wordTabStops(attrs.paragraphTabs);
              return stops ? { 'data-word-tabs': JSON.stringify(stops) } : {};
            },
          },
          paragraphDefaultTab: {
            default: null,
            parseHTML: (el: HTMLElement) => wordDefaultTab(el.getAttribute('data-word-default-tab')),
            renderHTML: (attrs: Record<string, unknown>) => {
              const interval = wordDefaultTab(attrs.paragraphDefaultTab);
              return interval === null ? {} : { 'data-word-default-tab': String(interval) };
            },
          },
          paragraphDecimalSymbol: {
            default: null,
            parseHTML: (el: HTMLElement) => wordDecimalSymbol(el.getAttribute('data-word-decimal-symbol')),
            renderHTML: (attrs: Record<string, unknown>) => {
              const symbol = wordDecimalSymbol(attrs.paragraphDecimalSymbol);
              return symbol ? { 'data-word-decimal-symbol': symbol } : {};
            },
          },
          paragraphTabUnsupported: {
            default: false,
            parseHTML: (el: HTMLElement) => el.getAttribute('data-word-tab-layout-unsupported') === 'true',
            renderHTML: (attrs: Record<string, unknown>) => attrs.paragraphTabUnsupported
              ? { 'data-word-tab-layout-unsupported': 'true' } : {},
          },
          ...Object.fromEntries(
            [
              ['paragraphFontSize', 'font-size'],
              ['paragraphFontFamily', 'font-family'],
            ].map(([name, css]) => [
              name,
              {
                default: null,
                parseHTML: (el: HTMLElement) =>
                  el.style.getPropertyValue(css).replace(/^(["'])(.*)\1$/, '$2') || null,
                renderHTML: (attrs: Record<string, unknown>) =>
                  attrs[name]
                    ? {
                        style: `${css}: ${name === 'paragraphFontFamily' ? JSON.stringify(String(attrs[name])) : attrs[name]}`,
                      }
                    : {},
              },
            ]),
          ),
          ...Object.fromEntries(
            paragraphIndents.map(([name, css]) => [
              name,
              {
                default: null,
                parseHTML: (element: HTMLElement) => element.style.getPropertyValue(css) || null,
                renderHTML: (attrs: Record<string, unknown>) =>
                  attrs[name] ? { style: `${css}: ${attrs[name]}` } : {},
              },
            ]),
          ),
          ...Object.fromEntries(
            paragraphBreaks.map(([name, css, on]) => [
              name,
              {
                default: null,
                parseHTML: (element: HTMLElement) => {
                  const value = element.style.getPropertyValue(css);
                  return value ? value === on : null;
                },
                renderHTML: (attrs: Record<string, unknown>) =>
                  attrs[name] == null
                    ? {}
                    : {
                        style:
                          name === 'widowControl'
                            ? `orphans: ${attrs[name] ? 2 : 1}; widows: ${attrs[name] ? 2 : 1}`
                            : `${css}: ${attrs[name] ? on : 'auto'}`,
                      },
              },
            ]),
          ),
          paragraphLineHeight: {
            default: null,
            parseHTML: (element) => element.style.lineHeight || null,
            renderHTML: (attrs) =>
              attrs.paragraphLineHeight
                ? {
                    style: `line-height: ${attrs.paragraphLineHeight}`,
                    ...(/(?:pt|px)$/.test(attrs.paragraphLineHeight)
                      ? {
                          'data-word-line-height':
                            wordLineSpacing(attrs.paragraphLineHeight, attrs.paragraphLineRule)
                              ?.rule || 'exact',
                        }
                      : {}),
                  }
                : {},
          },
          paragraphLineRule: {
            default: null,
            parseHTML: (element) =>
              wordLineSpacing(element.style.lineHeight, element.getAttribute('data-word-line-rule'))
                ?.rule === 'atLeast'
                ? 'atLeast'
                : null,
            renderHTML: (attrs) =>
              wordLineSpacing(attrs.paragraphLineHeight, attrs.paragraphLineRule)?.rule ===
              'atLeast'
                ? { 'data-word-line-rule': 'atLeast' }
                : {},
          },
          spaceBefore: {
            default: null,
            parseHTML: (element) => element.style.marginTop || null,
            renderHTML: (attrs) =>
              attrs.spaceBefore ? { style: `margin-top: ${attrs.spaceBefore}` } : {},
          },
          spaceAfter: {
            default: null,
            parseHTML: (element) => element.style.marginBottom || null,
            renderHTML: (attrs) =>
              attrs.spaceAfter ? { style: `margin-bottom: ${attrs.spaceAfter}` } : {},
          },
        },
      },
    ];
  },
});
