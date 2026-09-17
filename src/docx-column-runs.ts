import type { WordExportLayout } from './word-export-layout';
import { writeWordText } from './word-xml';

const word = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const value = (node: Element): string => {
  if (node.namespaceURI !== word) return '';
  if (node.localName === 't') return node.textContent || '';
  if (node.localName === 'tab') return '\t';
  if (node.localName === 'softHyphen') return '\u001f';
  if (node.localName === 'br') {
    const type = node.getAttributeNS(word, 'type');
    return type === 'page' ? '\f' : type === 'column' ? '\u000e' : '\n';
  }
  return [...node.children].map(value).join('');
};

/** Word Save splits plain text runs at physical column/page transitions. Run
 * boundaries affect first-open native PDF advances even with identical fonts.
 * Split after ordinary run normalization using current semantic layout offsets;
 * never add page controls or cached lastRenderedPageBreak elements. */
export function splitWordColumnRuns(doc: XMLDocument, layout: WordExportLayout) {
  const body = doc.getElementsByTagNameNS(word, 'body')[0];
  const paragraphs = [...body.children].filter(
    (e) => e.namespaceURI === word && e.localName === 'p',
  );
  if (paragraphs.length !== layout.paragraphs.length)
    throw Error('Document paragraphs changed before export. Try exporting again.');
  paragraphs.forEach((p, index) => {
    const { text, boundaries } = layout.paragraphs[index];
    if (
      value(p) !== text ||
      boundaries.some(
        (at, i) =>
          !Number.isSafeInteger(at) ||
          at <= 0 ||
          at >= text.length ||
          (i > 0 && at <= boundaries[i - 1]) ||
          (/[\uD800-\uDBFF]/.test(text[at - 1]) && /[\uDC00-\uDFFF]/.test(text[at])),
      )
    )
      throw Error('Document text changed before export. Try exporting again.');
    for (const at of boundaries) {
      let offset = 0;
      for (const run of [...p.children]) {
        const length = value(run).length;
        if (at === offset || at === offset + length) break;
        if (at > offset && at < offset + length) {
          if (
            run.namespaceURI !== word ||
            run.localName !== 'r' ||
            [...run.children].some(
              (e) =>
                e.namespaceURI !== word ||
                !['rPr', 't', 'br', 'tab', 'softHyphen', 'lastRenderedPageBreak'].includes(e.localName) ||
                (e.localName === 'softHyphen' && (e.childNodes.length > 0 ||
                  [...e.attributes].some(attr => attr.namespaceURI !== 'http://www.w3.org/2000/xmlns/'))),
            )
          )
            throw Error('This column boundary crosses unsupported document content.');
          const right = run.cloneNode(false) as Element;
          const properties = [...run.children].find((e) => e.localName === 'rPr');
          if (properties) right.append(properties.cloneNode(true));
          let position = offset;
          for (const node of [...run.children]) {
            if (node.localName === 'rPr') continue;
            const end = position + value(node).length;
            if (position >= at) right.append(node);
            else if (end > at) {
              if (node.localName !== 't') throw Error('Invalid text boundary during export.');
              const rest = node.cloneNode(false) as Element;
              writeWordText(rest, node.textContent!.slice(at - position));
              writeWordText(node, node.textContent!.slice(0, at - position));
              right.append(rest);
            }
            position = end;
          }
          run.after(right);
          break;
        }
        offset += length;
      }
    }
  });
}
