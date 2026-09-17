import { child, val, WORD_NS } from './word-xml';
import { wordDefaultTab, wordDecimalSymbol, wordTabStops, type WordTabStop } from './word-tab-stops';

const unsupported = 'data-noffice-unsupported-tabs';
function validTabs(tabs: Element) {
  if (tabs.hasAttribute(unsupported) || tabs.children.length > 256) return false;
  const values = [];
  for (const tab of tabs.children) {
    if (tab.namespaceURI !== WORD_NS || tab.localName !== 'tab' || tab.children.length ||
      [...tab.attributes].some((a) => a.namespaceURI !== 'http://www.w3.org/2000/xmlns/' &&
        (a.namespaceURI !== WORD_NS || !['val', 'pos', 'leader'].includes(a.localName))))
      return false;
    values.push({ position: /^[+-]?\d+$/.test(val(tab, 'pos')) ? Number(val(tab, 'pos')) : NaN,
      alignment: val(tab) === 'clear' ? 'left' : val(tab), leader: val(tab, 'leader') || 'none' });
  }
  return wordTabStops(values) !== null;
}

/** Tab lists extend the style cascade by position. A clear removes one inherited
 * custom stop; an empty list does not erase the inherited list. */
export function mergeDocxTabs(previous: Element, next: Element) {
  // Validate before a clear or overriding stop can erase an invalid definition.
  // This marker lives only in the disposable style-reading cascade.
  if (!validTabs(previous) || !validTabs(next)) next.setAttribute(unsupported, 'true');
  const stops = new Map<string, Element>();
  for (const source of [previous, next])
    for (const tab of source.children) {
      if (tab.namespaceURI !== WORD_NS || tab.localName !== 'tab') continue;
      const raw = val(tab, 'pos');
      const key = /^[+-]?\d+$/.test(raw) && Number.isSafeInteger(Number(raw)) ? String(Number(raw)) : raw;
      if (val(tab) === 'clear') stops.delete(key);
      else stops.set(key, tab);
    }
  next.replaceChildren(...[...stops.values()]
    .sort((a, b) => Number(val(a, 'pos')) - Number(val(b, 'pos')))
    .map((tab) => next.ownerDocument.importNode(tab, true)));
}

export function docxTabStops(properties: Element): WordTabStop[] | null {
  const tabs = child(properties, 'tabs');
  if (!tabs) return [];
  if (!validTabs(tabs)) return null;
  const values = [...tabs.children]
    .filter((tab) => tab.namespaceURI === WORD_NS && tab.localName === 'tab' && val(tab) !== 'clear')
    .map((tab) => ({
      position: /^[+-]?\d+$/.test(val(tab, 'pos')) ? Number(val(tab, 'pos')) : NaN,
      alignment: val(tab),
      leader: val(tab, 'leader') || 'none',
    }));
  return wordTabStops(values);
}

export function docxDefaultTab(settings?: Element) {
  const element = settings && child(settings, 'defaultTabStop');
  if (!element) return 720;
  const raw = val(element);
  return /^\+?\d+$/.test(raw) ? wordDefaultTab(Number(raw)) : null;
}

export function docxTabSettings(settings?: Element) {
  const element = settings && child(settings, 'decimalSymbol'), decimal = val(element);
  return { interval: docxDefaultTab(settings), decimal,
    decimalInvalid: !!element && wordDecimalSymbol(decimal) === null };
}
