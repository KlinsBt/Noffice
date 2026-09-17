import { child, val, WORD_NS, wordXml } from './word-xml';
import { wordListSourceSchema, type WordListSource } from './word-list-layout';

const children = (el: Element, name: string) => [...el.children]
  .filter(e => e.namespaceURI === WORD_NS && e.localName === name);
const unique = (el: Element, name: string) => {
  const matches = children(el, name);
  if (matches.length > 1) throw Error('Duplicate numbering property.');
  return matches[0];
};
const integer = (value: string) => /^\d+$/.test(value) ? Number(value) : NaN;
const only = (el: Element | undefined, names: string[]) => !el || [...el.children]
  .every(e => e.namespaceURI === WORD_NS && names.includes(e.localName));

/** Resolve the qualified flat-list definitions. This does not mutate pPr or
 * promote unsupported multi-level, linked-style, picture or custom markers. */
export function docxNumbering(xml?: string) {
  const root = xml ? wordXml(xml).documentElement : undefined;
  const index = (name: string, attribute: string) => {
    const map = new Map<string, Element | null>();
    for (const element of root ? children(root, name) : []) {
      const id = val(element, attribute);
      map.set(id, map.has(id) ? null : element);
    }
    return map;
  };
  const instances = index('num', 'numId'), abstracts = index('abstractNum', 'abstractNumId');
  return (properties: Element): WordListSource | null => {
    const numbering = child(properties, 'numPr');
    if (!numbering) return null;
    try {
      if (unique(properties, 'numPr') !== numbering || !only(numbering, ['numId', 'ilvl']))
        throw Error('Unsupported paragraph numbering properties.');
      const numId = val(unique(numbering, 'numId'));
      if (numId === '0') return null;
      const level = integer(val(unique(numbering, 'ilvl')) || '0');
      if (level !== 0) throw Error('Nested numbering needs additional layout support.');
      const instance = instances.get(numId);
      if (!instance || !only(instance, ['abstractNumId', 'lvlOverride']))
        throw Error('Missing, duplicate or unsupported numbering instance.');
      const abstract = abstracts.get(val(unique(instance, 'abstractNumId')));
      if (!abstract || !only(abstract, ['nsid', 'multiLevelType', 'tmpl', 'name', 'lvl']))
        throw Error('Missing, duplicate or linked abstract numbering definition.');
      const restart = abstract.getAttributeNS('http://schemas.microsoft.com/office/word/2012/wordml', 'restartNumberingAfterBreak');
      if (restart && !['0', 'false', 'off'].includes(restart))
        throw Error('Section-dependent numbering restart needs additional support.');
      const levels = children(abstract, 'lvl').filter(e => val(e, 'ilvl') === '0');
      const overrides = children(instance, 'lvlOverride').filter(e => val(e, 'ilvl') === '0');
      if (levels.length !== 1 || overrides.length > 1)
        throw Error('Missing or duplicate numbering level.');
      const override = overrides[0];
      if (!only(override, ['startOverride', 'lvl'])) throw Error('Unsupported numbering override.');
      const levelElement = (override && unique(override, 'lvl')) || levels[0];
      if (val(levelElement, 'ilvl') !== '0' || !only(levelElement,
        ['start', 'numFmt', 'lvlText', 'lvlJc', 'suff', 'pPr', 'rPr']))
        throw Error('Unsupported numbering level properties.');
      // Duplicate properties are invalid even if the first one looks supported.
      for (const e of levelElement.children) unique(levelElement, e.localName);
      for (const name of ['start', 'numFmt', 'lvlText', 'lvlJc', 'suff']) {
        const e = child(levelElement, name);
        if (e && (e.children.length || [...e.attributes].some(a => a.namespaceURI !== WORD_NS || a.localName !== 'val')))
          throw Error('Unsupported numbering value attributes.');
      }
      const suffix = val(child(levelElement, 'suff'));
      const alignment = val(child(levelElement, 'lvlJc'));
      if ((suffix && suffix !== 'tab') || (alignment && alignment !== 'left'))
        throw Error('Non-left or non-tab numbering needs additional layout support.');
      const pPr = child(levelElement, 'pPr'), ind = pPr && unique(pPr, 'ind');
      if (!only(pPr, ['ind']) || !ind || [...ind.attributes].some(a =>
        a.namespaceURI !== WORD_NS || !['left', 'hanging'].includes(a.localName)))
        throw Error('Unsupported numbering indentation or tab properties.');
      // Do not guess precedence for direct/style paragraph indents until the
      // native geometry matrix for those overrides is qualified.
      if (child(properties, 'ind') || child(properties, 'tabs') || child(properties, 'bidi'))
        throw Error('Paragraph overrides need additional numbering layout support.');
      const mark = child(properties, 'rPr'), marker = child(levelElement, 'rPr');
      if (mark && ['b', 'i', 'bCs', 'iCs', 'vanish', 'rtl', 'color', 'highlight', 'u', 'strike', 'vertAlign', 'position', 'spacing', 'w']
        .some(name => {
          const property = child(mark, name);
          const value = val(property);
          return property && !((['b', 'i', 'bCs', 'iCs', 'vanish', 'rtl', 'strike'].includes(name)
            && ['0', 'false', 'off'].includes(value)) || (name === 'vertAlign' && value === 'baseline') ||
            (['position', 'spacing'].includes(name) && value === '0') ||
            (name === 'w' && value === '100') || (name === 'u' && value === 'none') ||
            (name === 'color' && ['auto', '000000'].includes(value)));
        })) throw Error('Paragraph-mark typography needs additional numbering support.');
      if (!only(marker, ['rFonts', 'sz'])) throw Error('Unsupported numbering marker typography.');
      if (marker) for (const e of marker.children) unique(marker, e.localName);
      const fonts = marker && child(marker, 'rFonts') || mark && child(mark, 'rFonts');
      if (fonts && [...fonts.attributes].some(a => a.namespaceURI !== WORD_NS ||
        !['ascii', 'hAnsi', 'cs', 'eastAsia', 'hint'].includes(a.localName)))
        throw Error('The numbering marker font is unresolved.');
      const size = marker && child(marker, 'sz') || mark && child(mark, 'sz');
      const start = override && unique(override, 'startOverride') || child(levelElement, 'start');
      const result = wordListSourceSchema.safeParse({ status: 'resolved', definition: {
        numId, level, start: integer(val(start)), format: val(child(levelElement, 'numFmt')),
        text: val(child(levelElement, 'lvlText')), left: integer(val(ind, 'left')),
        hanging: integer(val(ind, 'hanging')), font: val(fonts, 'ascii'), size: integer(val(size)) / 2,
      } });
      if (!result.success) throw Error('Unsupported numbering format, font, size or geometry.');
      return result.data;
    } catch (error) {
      return { status: 'unsupported', reason: error instanceof Error ? error.message : 'Invalid numbering definition.' };
    }
  };
}
