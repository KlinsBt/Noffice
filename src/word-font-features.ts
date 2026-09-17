/** Word's four ligature switches are independent of contextual alternates.
 * Values follow the pinned Word bit flags, not the OpenXML enum ordinals. */
export const WORD_2010_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
export const wordLigatureNames = [
  'none', 'standard', 'contextual', 'standardContextual',
  'historical', 'standardHistorical', 'contextualHistorical', 'standardContextualHistorical',
  'discretional', 'standardDiscretional', 'contextualDiscretional', 'standardContextualDiscretional',
  'historicalDiscretional', 'standardHistoricalDiscretional', 'contextualHistoricalDiscretional', 'all',
] as const;

export function wordFontFeaturesValue(value: unknown): number | null {
  if (typeof value !== 'string' && typeof value !== 'number') return null;
  if (!/^(?:[0-9]|[12][0-9]|3[01])$/.test(String(value))) return null;
  return Number(value);
}

export function wordFontFeaturesXml(properties?: Element): number {
  const find = (name: string) => properties && [...properties.children]
    .find((element) => element.namespaceURI === WORD_2010_NS && element.localName === name);
  const ligatures = find('ligatures'), alternates = find('cntxtAlts');
  const value = ligatures?.getAttributeNS(WORD_2010_NS, 'val');
  const flags = ligatures ? wordLigatureNames.findIndex((name) => name === value) : 0;
  const alternateValue = alternates?.getAttributeNS(WORD_2010_NS, 'val');
  if (flags < 0 || (alternates && alternateValue !== null
    && !['0', '1', 'true', 'false', 'on', 'off'].includes(alternateValue!)))
    throw Error('The Word document contains an invalid font-feature setting.');
  return flags | (alternates && !['0', 'false', 'off'].includes(alternateValue || '') ? 16 : 0);
}

export function wordFontFeaturesStyle(flags: number): string {
  return ['liga', 'clig', 'hlig', 'dlig', 'calt']
    .map((tag, index) => `"${tag}" ${flags & (1 << index) ? 1 : 0}`).join(', ');
}

/** Write absolute overrides for both independently inherited properties. */
export function writeWordFontFeatures(properties: Element, value: unknown) {
  const flags = wordFontFeaturesValue(value);
  if (flags === null) throw Error('Invalid Word font-feature setting.');
  for (const [name, setting] of [['ligatures', wordLigatureNames[flags & 15]],
    ['cntxtAlts', flags & 16 ? '1' : '0']]) {
    let element = [...properties.children]
      .find((child) => child.namespaceURI === WORD_2010_NS && child.localName === name);
    if (!element) {
      element = properties.ownerDocument.createElementNS(WORD_2010_NS, `w14:${name}`);
      properties.append(element);
    }
    element.setAttributeNS(WORD_2010_NS, 'w14:val', setting);
  }
}
