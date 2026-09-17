import type JSZip from 'jszip';
import { child, descendants, val, wElement, wordXml } from './word-xml';
import { emptyStoryDocument } from './docx-empty-stories';

/** New packages use the same native built-ins as retained creation. Apply them
 * to ordinary story paragraphs after packing; explicit heading/list styles
 * remain intact. This never rewrites body paragraphs or embedded fonts. */
export async function styleAuthoredDocxStories(zip: JSZip) {
  for (const kind of ['header', 'footer'] as const) {
    const parts = Object.values(zip.files).filter(
      (p) => !p.dir && new RegExp(`^word/${kind}\\d+\\.xml$`).test(p.name),
    );
    if (!parts.length) continue;
    const template = await emptyStoryDocument(zip, kind);
    const styleId = val(child(child(descendants(template, 'p')[0], 'pPr')!, 'pStyle'));
    for (const part of parts) {
      const doc = wordXml(await part.async('string'));
      for (const p of descendants(doc, 'p')) {
        let props = child(p, 'pPr');
        if (!props) {
          props = wElement(doc, 'pPr');
          p.prepend(props);
        }
        const style = child(props, 'pStyle');
        if (style && !['Header', 'Footer'].includes(val(style))) continue;
        if (style) style.remove();
        props.prepend(wElement(doc, 'pStyle', { val: styleId }));
      }
      zip.file(part.name, new XMLSerializer().serializeToString(doc), { createFolders: false });
    }
  }
}
