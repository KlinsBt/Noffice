import { child, wElement, wordXml, WORD_NS } from './word-xml';

/** Pinned native Word's application defaults apply only when the corresponding
 * default container is absent. An explicitly empty container suppresses them.
 * These nodes belong to the reading copy; retained source XML stays unchanged. */
export function docxStyleDefaults(defaults?: Element, legacy = false) {
  const doc = defaults?.ownerDocument || wordXml(`<w:styles xmlns:w="${WORD_NS}"/>`);
  const runContainer = defaults && child(defaults, 'rPrDefault');
  const paragraphContainer = defaults && child(defaults, 'pPrDefault');
  let run = runContainer && child(runContainer, 'rPr');
  let paragraph = paragraphContainer && child(paragraphContainer, 'pPr');
  if (!legacy && !runContainer) {
    run = wElement(doc, 'rPr');
    run.append(
      wElement(doc, 'rFonts', {
        ascii: 'Calibri',
        hAnsi: 'Calibri',
        asciiTheme: 'minorHAnsi',
        hAnsiTheme: 'minorHAnsi',
      }),
      wElement(doc, 'sz', { val: '22' }),
    );
  }
  if (!legacy && !paragraphContainer) {
    paragraph = wElement(doc, 'pPr');
    paragraph.append(wElement(doc, 'spacing', { after: '160', line: '259', lineRule: 'auto' }));
  }
  return { run, paragraph };
}
