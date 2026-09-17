import type JSZip from 'jszip';
import { child, descendants, val, wElement, wordXml, WORD_NS } from './word-xml';
import { docxStyles } from './docx-styles';
import { applyDocxParagraphReading } from './docx-paragraph-reading';
import { emptyStylesXml, ensureStoryStylePart } from './docx-story-style-part';
import { docxTabSettings } from './docx-tabs';
import { readWordSettings } from './docx-settings';

function privateStyleIdentity(occupied: Set<string>, base: string) {
  let id = base,
    n = 0;
  while (occupied.has(id)) id = base + ++n;
  occupied.add(id);
  return id;
}

/** Native Word restores missing default base styles as unformatted roots. Keep
 * the document's existing defaults/cascade; never replace an occupied style ID. */
function storyBaseStyles(styles: XMLDocument, occupied: Set<string>) {
  const all = descendants(styles, 'style');
  const roots = (type: string) =>
    all.filter((s) => val(s, 'type') === type && ['1', 'true', 'on'].includes(val(s, 'default')));
  const paragraphs = roots('paragraph'),
    characters = roots('character');
  if (paragraphs.length > 1 || characters.length > 1)
    throw Error('The document has ambiguous default style definitions.');
  let normal = paragraphs[0],
    character = characters[0];
  if (!normal) {
    normal = wElement(styles, 'style', {
      type: 'paragraph',
      default: '1',
      styleId: privateStyleIdentity(occupied, 'NofficeNormal'),
    });
    normal.append(wElement(styles, 'name', { val: 'Normal' }), wElement(styles, 'qFormat'));
    styles.documentElement.append(normal);
  }
  if (!character) {
    character = wElement(styles, 'style', {
      type: 'character',
      default: '1',
      styleId: privateStyleIdentity(occupied, 'NofficeDefaultParagraphFont'),
    });
    character.append(
      wElement(styles, 'name', { val: 'Default Paragraph Font' }),
      wElement(styles, 'uiPriority', { val: '1' }),
      wElement(styles, 'semiHidden'),
      wElement(styles, 'unhideWhenUsed'),
    );
    styles.documentElement.append(character);
  }
  return { normal, character };
}

/** The pinned Word built-ins inherit document Normal/Default Paragraph Font.
 * Private identities avoid overwriting localized or custom styles. The measured
 * 4536/9072-twip tab stops belong to the baseline built-in, not the page width. */
function storyStyle(styles: XMLDocument, kind: 'header' | 'footer', create: boolean) {
  const all = descendants(styles, 'style');
  const matches = all.filter(
    (s) => val(s, 'type') === 'paragraph' && val(child(s, 'name')).toLowerCase() === kind,
  );
  if (matches.length > 1) throw Error('The document has ambiguous header/footer styles.');
  if (matches[0] || !create) return matches[0];
  const occupied = new Set(all.map((s) => val(s, 'styleId')));
  const { normal, character } = storyBaseStyles(styles, occupied);
  const identity = (base: string) => privateStyleIdentity(occupied, base);
  const label = kind === 'header' ? 'Header' : 'Footer';
  const id = identity('Noffice' + label),
    charId = identity('Noffice' + label + 'Char');
  const paragraph = wElement(styles, 'style', { type: 'paragraph', styleId: id });
  paragraph.append(
    wElement(styles, 'name', { val: kind }),
    wElement(styles, 'basedOn', { val: val(normal, 'styleId') }),
    wElement(styles, 'link', { val: charId }),
    wElement(styles, 'uiPriority', { val: '99' }),
    wElement(styles, 'unhideWhenUsed'),
  );
  const properties = wElement(styles, 'pPr'),
    tabs = wElement(styles, 'tabs');
  tabs.append(
    wElement(styles, 'tab', { val: 'center', pos: '4536' }),
    wElement(styles, 'tab', { val: 'right', pos: '9072' }),
  );
  properties.append(
    tabs,
    wElement(styles, 'spacing', { after: '0', line: '240', lineRule: 'auto' }),
  );
  paragraph.append(properties);
  const linked = wElement(styles, 'style', {
    type: 'character',
    customStyle: '1',
    styleId: charId,
  });
  linked.append(
    wElement(styles, 'name', { val: 'Noffice ' + label + ' Char' }),
    wElement(styles, 'basedOn', { val: val(character, 'styleId') }),
    wElement(styles, 'link', { val: id }),
    wElement(styles, 'uiPriority', { val: '99' }),
  );
  styles.documentElement.append(paragraph, linked);
  return paragraph;
}

/** Derive templates without changing the retained archive. Missing built-ins
 * use the measured native cascade; missing base style definitions stay explicit. */
export async function docxEmptyStoryTemplates(zip: JSZip, legacyFontFeatures = false, legacyParagraphScripts = false, legacyParagraphSpacing = false) {
  const settings = (await readWordSettings(zip))?.document.documentElement;
  const stylesXml = await zip.file('word/styles.xml')?.async('string');
  const output: { header: string | null; footer: string | null } = { header: null, footer: null };
  const styles = wordXml(stylesXml || emptyStylesXml);
  const ids = {
    header: storyStyle(styles, 'header', true),
    footer: storyStyle(styles, 'footer', true),
  };
  const resolve = docxStyles(
    new XMLSerializer().serializeToString(styles),
    await zip.file('word/theme/theme1.xml')?.async('string'),
  );
  for (const kind of ['header', 'footer'] as const) {
    const style = ids[kind];
    if (!style) continue;
    const paragraph = emptyStoryParagraph(kind, val(style, 'styleId'));
    const html = document.createElement('p');
    html.dataset.sourceParagraph = 'empty:0';
    applyDocxParagraphReading(html, resolve(paragraph), docxTabSettings(settings), legacyFontFeatures, legacyParagraphScripts, legacyParagraphSpacing);
    output[kind] = html.outerHTML;
  }
  return output;
}

export function emptyStoryParagraph(kind: 'header' | 'footer', styleId: string) {
  const doc = wordXml(`<w:${kind === 'header' ? 'hdr' : 'ftr'} xmlns:w="${WORD_NS}"/>`);
  const p = wElement(doc, 'p'),
    props = wElement(doc, 'pPr');
  props.append(wElement(doc, 'pStyle', { val: styleId }));
  p.append(props);
  doc.documentElement.append(p);
  return p;
}

export async function emptyStoryDocument(zip: JSZip, kind: 'header' | 'footer') {
  await ensureStoryStylePart(zip);
  const xml = await zip.file('word/styles.xml')?.async('string');
  if (!xml) throw Error('Creating this empty story requires the document style definitions.');
  const styles = wordXml(xml),
    existing = storyStyle(styles, kind, false),
    style = existing || storyStyle(styles, kind, true);
  if (!style)
    throw Error('Creating this empty story requires Normal and Default Paragraph Font styles.');
  if (!existing)
    zip.file('word/styles.xml', new XMLSerializer().serializeToString(styles), {
      createFolders: false,
    });
  return emptyStoryParagraph(kind, val(style, 'styleId')).ownerDocument;
}
