import { child, descendants, val, wElement, wordXml, WORD_NS } from './word-xml';
import { docxStyleDefaults } from './docx-style-defaults';
import { mergeDocxTabs } from './docx-tabs';
import { WORD_2010_NS } from './word-font-features';

/** Resolve Word's document defaults and paragraph/character basedOn chains for the reading copy. */
export function docxStyles(xml?: string, themeXml?: string, legacyDefaults = false) {
  const styles = xml ? wordXml(xml) : undefined;
  const theme = themeXml ? wordXml(themeXml) : undefined;
  const drawingNS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const themeFonts = new Map(
    ['major', 'minor'].map((name) => [
      name,
      theme
        ?.getElementsByTagNameNS(drawingNS, `${name}Font`)[0]
        ?.getElementsByTagNameNS(drawingNS, 'latin')[0]
        ?.getAttribute('typeface'),
    ]),
  );
  const entries = new Map(
    styles ? descendants(styles, 'style').map((style) => [val(style, 'styleId'), style]) : [],
  );
  const defaults = styles && descendants(styles, 'docDefaults')[0];
  const application = docxStyleDefaults(defaults, legacyDefaults);
  const paragraphDefault = [...entries.values()].find(
    (s) => val(s, 'type') === 'paragraph' && ['1', 'true', 'on'].includes(val(s, 'default')),
  );
  const cache = new Map<string, Element[]>();
  function chain(id: string, seen = new Set<string>()): Element[] {
    if (!id || seen.has(id) || seen.size > 32) return [];
    const cached = cache.get(id);
    if (cached) return cached;
    const style = entries.get(id);
    if (!style) return [];
    seen.add(id);
    const result = [...chain(val(child(style, 'basedOn')), seen), style];
    cache.set(id, result);
    return result;
  }
  function merge(doc: Document, name: string, sources: (Element | undefined)[]) {
    const result = wElement(doc, name);
    for (const source of sources)
      if (source)
        for (const original of source.children) {
          if ((original.namespaceURI !== WORD_NS && !(original.namespaceURI === WORD_2010_NS
            && ['ligatures', 'cntxtAlts'].includes(original.localName))) || /Change$/.test(original.localName)) continue;
          const old = [...result.children].find((element) => element.namespaceURI === original.namespaceURI
            && element.localName === original.localName);
          const next = doc.importNode(original, true);
          if (old && original.localName === 'tabs') mergeDocxTabs(old, next);
          if (
            source.parentElement?.localName === 'style' &&
            [
              'b',
              'bCs',
              'i',
              'iCs',
              'caps',
              'smallCaps',
              'strike',
              'outline',
              'shadow',
              'emboss',
              'imprint',
              'vanish',
            ].includes(original.localName)
          ) {
            // Toggle properties in styles invert inherited state. Direct run formatting is absolute.
            if (['0', 'false', 'off'].includes(val(original))) continue;
            if (old && !['0', 'false', 'off'].includes(val(old)))
              next.setAttributeNS(WORD_NS, 'w:val', '0');
          }
          // Composite font, spacing and indentation properties inherit unspecified attributes.
          if (old && ['rFonts', 'spacing', 'ind', 'lang'].includes(original.localName)) {
            for (const attr of old.attributes)
              if (!next.hasAttributeNS(attr.namespaceURI, attr.localName)) {
                if (
                  original.localName === 'rFonts' &&
                  /Theme$/i.test(attr.localName) &&
                  next.hasAttributeNS(WORD_NS, attr.localName.replace(/Theme$/i, ''))
                )
                  continue;
                next.setAttributeNS(attr.namespaceURI, attr.name, attr.value);
              }
          }
          if (old) old.replaceWith(next);
          else result.append(next);
        }
    return result;
  }
  function resolveTheme(properties: Element) {
    if (!legacyDefaults && !child(properties, 'rFonts'))
      properties.append(
        wElement(properties.ownerDocument, 'rFonts', {
          ascii: 'Times New Roman',
          hAnsi: 'Times New Roman',
        }),
      );
    const fonts = child(properties, 'rFonts');
    const family = themeFonts.get(val(fonts, 'asciiTheme').startsWith('major') ? 'major' : 'minor');
    if (fonts && val(fonts, 'asciiTheme') && family)
      fonts.setAttributeNS(WORD_NS, 'w:ascii', family);
  }
  return (p: Element) => {
    const doc = p.ownerDocument;
    const direct = child(p, 'pPr');
    const inherited = chain(
      val(direct && child(direct, 'pStyle')) || val(paragraphDefault, 'styleId'),
    );
    const paragraph = merge(doc, 'pPr', [
      application.paragraph,
      ...inherited.map((s) => child(s, 'pPr')),
      direct,
    ]);
    paragraph.setAttribute('data-noffice-style-id', val(direct && child(direct, 'pStyle')) || val(paragraphDefault, 'styleId') || 'Normal');
    const mark = merge(doc, 'rPr', [
      application.run,
      ...inherited.map((s) => child(s, 'rPr')),
      child(paragraph, 'rPr'),
    ]);
    if (!child(mark, 'sz')) mark.append(wElement(doc, 'sz', { val: '20' }));
    resolveTheme(mark);
    const oldMark = child(paragraph, 'rPr');
    if (oldMark) oldMark.replaceWith(mark);
    else paragraph.append(mark);
    if (direct) direct.replaceWith(paragraph);
    else p.prepend(paragraph);
    for (const run of descendants(p, 'r')) {
      // Text boxes have their own paragraph cascade and are resolved independently.
      let parent = run.parentElement;
      while (parent && parent !== p && parent.localName !== 'p') parent = parent.parentElement;
      if (parent !== p) continue;
      const directRun = child(run, 'rPr');
      const character = chain(val(directRun && child(directRun, 'rStyle')));
      const resolved = merge(doc, 'rPr', [
        application.run,
        ...inherited.map((s) => child(s, 'rPr')),
        ...character.map((s) => child(s, 'rPr')),
        directRun,
      ]);
      // Pinned Word uses 10pt when no size exists anywhere in the cascade.
      // This is an application fallback, not an OOXML-mandated default.
      if (!child(resolved, 'sz')) resolved.append(wElement(doc, 'sz', { val: '20' }));
      resolveTheme(resolved);
      if (directRun) directRun.replaceWith(resolved);
      else run.prepend(resolved);
    }
    return paragraph;
  };
}
