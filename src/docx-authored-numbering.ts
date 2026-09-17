import type { INumberingOptions } from 'docx';
import { wordNumberingSchema, type WordNumbering, type WordListDefinition } from './word-list-layout';

const attribute = 'data-noffice-export-numbering';
const unavailable = () => Error('This list needs its retained DOCX numbering source. Export a Noffice backup to preserve it.');

/** Assign export-local identities before section ranges clone their list wrappers.
 * Source numIds continue across wrappers; independently authored lists restart.
 * These attributes belong only to the disposable export DOM, never saved HTML. */
export function authoredWordNumbering(bodies: HTMLElement[], metadata?: WordNumbering) {
  const source = new Map(wordNumberingSchema.parse(metadata || { version: 1, paragraphs: [] })
    .paragraphs.map(p => [p.source, p.numbering]));
  const definitions = new Map<string, WordListDefinition>();
  const config: INumberingOptions['config'][number][] = [];
  let serial = 0;
  for (const body of bodies) {
    for (const el of body.querySelectorAll(`[${attribute}]`)) el.removeAttribute(attribute);
    for (const list of body.querySelectorAll('ol,ul')) {
      const paragraphs = [...list.querySelectorAll('p,h1,h2,h3,h4,h5,h6')]
        .filter(p => p.closest('ol,ul') === list);
      const entries = paragraphs.map(p => source.get(p.getAttribute('data-source-paragraph') || ''));
      if (paragraphs.some(p => p.hasAttribute('data-source-paragraph')) && entries.some(e => !e)) throw unavailable();
      if (list.hasAttribute('reversed') || [...list.children].some(li => li.hasAttribute('value'))) throw unavailable();
      if (entries.some(Boolean)) {
        // The qualified source profile is flat, with one paragraph per item.
        if (list.parentElement !== body || [...list.children].some(li => li.tagName !== 'LI' ||
            li.children.length !== 1 || li.firstElementChild?.tagName !== 'P')) throw unavailable();
        paragraphs.forEach((p, index) => {
          const entry = entries[index];
          if (entry?.status !== 'resolved' || (entry.definition.format === 'decimal') !== (list.tagName === 'OL'))
            throw unavailable();
          const d = entry.definition, reference = `source-list-${d.numId}`;
          if (list.tagName === 'OL' && (list.hasAttribute('type') ||
              ![1, d.start].includes(Number(list.getAttribute('start') || '1')))) throw unavailable();
          const previous = definitions.get(d.numId);
          if (previous && JSON.stringify(previous) !== JSON.stringify(d)) throw unavailable();
          if (!previous) {
            definitions.set(d.numId, d);
            config.push({ reference, levels: [{ level: 0, format: d.format, text: d.text,
              start: d.start, alignment: 'left', suffix: 'tab', style: {
                paragraph: { indent: { left: d.left, hanging: d.hanging } },
                run: { font: d.font, size: d.size * 2, sizeComplexScript: false },
              } }] });
          }
          p.setAttribute(attribute, reference);
        });
      } else if (list.tagName === 'OL') {
        const rawStart = list.getAttribute('start') || '1';
        if (!/^\d{1,7}$/.test(rawStart) || Number(rawStart) > 1000000 || list.hasAttribute('reversed') ||
            list.getAttribute('type') && list.getAttribute('type') !== '1') throw unavailable();
        const start = Number(rawStart), reference = `authored-list-${serial++}`;
        let level = 0;
        for (let parent = list.parentElement; parent && parent !== body; parent = parent.parentElement)
          if (parent.tagName === 'OL' || parent.tagName === 'UL') level++;
        if (level > 8 || level && start !== 1) throw unavailable();
        config.push({ reference, levels: Array.from({ length: 9 }, (_, level) => ({
          level, format: 'decimal', text: `%${level + 1}.`, alignment: 'start', start: level ? 1 : start,
        })) });
        for (const p of paragraphs) p.setAttribute(attribute, reference);
      }
    }
  }
  return { config, reference: (p: Element) => p.getAttribute(attribute) || undefined };
}
