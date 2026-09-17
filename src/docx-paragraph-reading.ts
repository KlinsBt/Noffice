import { child, val } from './word-xml';
import { wordKerningXmlValue } from './word-kerning';
import { paragraphBreaks } from './paragraph-layout';
import { docxTabStops } from './docx-tabs';
import { wordDecimalSymbol } from './word-tab-stops';
import { wordFontFeaturesXml, wordFontFeaturesStyle } from './word-font-features';
import { wordScriptValue } from './word-script';
import { wordParagraphSpace, wordParagraphSpaceTwips } from './word-paragraph-spacing';

/** Apply the resolved paragraph/mark cascade to the reading view. */
export function applyDocxParagraphReading(
  el: HTMLElement,
  pr: Element,
  tabs?: { interval: number | null; decimal: string; decimalInvalid?: boolean },
  legacyFontFeatures = false,
  legacyParagraphScripts = false,
  legacyParagraphSpacing = false,
) {
  if (child(pr, 'tabs')) {
    const stops = docxTabStops(pr);
    if (stops) el.dataset.wordTabs = JSON.stringify(stops);
    else el.dataset.wordTabLayoutUnsupported = 'true';
  }
  if (tabs) {
    if (tabs.interval === null) el.dataset.wordTabLayoutUnsupported = 'true';
    else el.dataset.wordDefaultTab = String(tabs.interval);
    if (tabs.decimalInvalid) el.dataset.wordTabLayoutUnsupported = 'true';
    const decimal = wordDecimalSymbol(tabs.decimal);
    if (decimal) el.dataset.wordDecimalSymbol = decimal;
  }
  const mark = child(pr, 'rPr');
  if (!legacyParagraphScripts) {
    const align = mark && child(mark, 'vertAlign');
    const value = align ? wordScriptValue(val(align)) : 'baseline';
    if (!value) throw Error('The document has an unsupported paragraph script setting.');
    el.dataset.wordParagraphScript = value;
  }
  if (!legacyFontFeatures) {
    const features = wordFontFeaturesXml(mark);
    el.dataset.wordParagraphFontFeatures = String(features);
    el.style.fontFeatureSettings = wordFontFeaturesStyle(features);
  }
  const kerning = wordKerningXmlValue(val(mark && child(mark, 'kern')));
  if (kerning !== null) el.dataset.wordParagraphKerning = String(kerning);
  const markSize = Number(val(mark && child(mark, 'sz'))) / 2;
  if (markSize > 0 && markSize <= 1638) el.style.fontSize = `${markSize}pt`;
  const markFamily = val(mark && child(mark, 'rFonts'), 'ascii');
  if (markFamily) el.style.fontFamily = JSON.stringify(markFamily.slice(0, 200));
  const align = val(child(pr, 'jc'));
  if (['left', 'right', 'center', 'both', 'start', 'end'].includes(align))
    el.style.textAlign = align === 'both' ? 'justify' : align;
  const bidi = child(pr, 'bidi');
  if (bidi) el.dir = ['0', 'false', 'off'].includes(val(bidi)) ? 'ltr' : 'rtl';
  const spacing = child(pr, 'spacing');
  if (!legacyParagraphSpacing) {
    el.dataset.wordParagraphStyle = pr.getAttribute('data-noffice-style-id') || 'Normal';
    const contextual = child(pr, 'contextualSpacing');
    if (contextual) el.dataset.wordContextualSpacing = String(!['0', 'false', 'off'].includes(val(contextual)));
  }
  for (const [name, css, on] of paragraphBreaks) {
    const flag = child(pr, name);
    if (flag) {
      const enabled = !['0', 'false', 'off'].includes(val(flag));
      el.style.setProperty(css, enabled ? on : name === 'widowControl' ? '1' : 'auto');
      if (name === 'widowControl') el.style.widows = enabled ? '2' : '1';
    }
  }
  const indent = child(pr, 'ind');
  for (const [modern, legacy, css] of [
    ['start', 'left', 'margin-inline-start'],
    ['end', 'right', 'margin-inline-end'],
  ]) {
    const amount = val(indent, modern) || val(indent, legacy);
    if (/^-?\d+$/.test(amount)) el.style.setProperty(css, `${Number(amount) / 20}pt`);
  }
  const hanging = val(indent, 'hanging'),
    first = val(indent, 'firstLine');
  if (/^\d+$/.test(hanging || first))
    el.style.textIndent = `${hanging ? -Number(hanging) / 20 : Number(first) / 20}pt`;
  for (const [attr, property] of [
    ['before', 'marginTop'],
    ['after', 'marginBottom'],
  ] as const) {
    const amount = val(spacing, attr);
    if (/^\d+$/.test(amount)) el.style[property] = `${Number(amount) / 20}pt`;
    if (!legacyParagraphSpacing) {
      const automatic = val(spacing, attr + 'Autospacing');
      const lines = val(spacing, attr + 'Lines');
      const mode = ['1', 'true', 'on'].includes(automatic) ? { unit: 'auto' }
        : /^\d+$/.test(lines) && Number(lines) > 0 ? { unit: 'lines', value: Number(lines) } : null;
      if (mode) {
        const space = wordParagraphSpace(mode);
        if (!space) throw Error('This document has unsupported paragraph spacing.');
        el.setAttribute('data-word-space-' + attr, JSON.stringify(space));
        el.style[property] = `${wordParagraphSpaceTwips(space) / 20}pt`;
      }
    }
  }
  if (val(spacing, 'line') && (!val(spacing, 'lineRule') || val(spacing, 'lineRule') === 'auto'))
    el.style.lineHeight = String(Number(val(spacing, 'line')) / 240);
  else if (
    ['exact', 'atLeast'].includes(val(spacing, 'lineRule')) &&
    /^\d+$/.test(val(spacing, 'line')) &&
    Number(val(spacing, 'line')) > 0 &&
    Number(val(spacing, 'line')) <= 31680
  ) {
    el.style.lineHeight = `${Number(val(spacing, 'line')) / 20}pt`;
    if (val(spacing, 'lineRule') === 'atLeast') el.setAttribute('data-word-line-rule', 'atLeast');
  } else if (
    !val(spacing, 'line') &&
    (!val(spacing, 'lineRule') || val(spacing, 'lineRule') === 'auto')
  )
    el.style.lineHeight = '1';
}
