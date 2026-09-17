import JSZip from 'jszip';
import { assertXmlComplexity } from './office-preservation';
import { docxTypography } from './docx-typography';
import type { WordContent } from './model';
import { docxStyles } from './docx-styles';
import { applyDocxParagraphReading } from './docx-paragraph-reading';
import { docxEmptyStoryTemplates } from './docx-empty-stories';
import { liveSectionProperties, readDocxStructure } from './docx-sections';
import { readWordStorySources } from './word-story-source';
import { appendWordStoryReading, splitWordStoryReading } from './word-story-reading';
import { wordStoriesSchema } from './word-stories';
import { docxTabSettings } from './docx-tabs';
import { readWordSettings } from './docx-settings';
import { readWordCompatibility } from './word-compatibility';
import { wordFinalSectionProperties } from './word-section-defaults';
import { wordHyphenText } from './word-hyphen';
import { docxNumbering } from './docx-numbering';
import type { WordNumbering } from './word-list-layout';

import { WORD_NS, descendants, child, val, wordXml, wElement } from './word-xml';
export { WORD_NS, descendants, child, val, wordXml, wElement } from './word-xml';

/** Annotate a disposable reading copy. The original package is never changed by import. */
export async function readDocx(data: ArrayBuffer, options: { legacyStyleDefaults?: boolean; legacyRunColors?: boolean; legacyFontFeatures?: boolean; legacyParagraphScripts?: boolean; legacyParagraphSpacing?: boolean; legacySectionDefaults?: boolean; legacyHyphens?: boolean } = {}) {
  const zip = await JSZip.loadAsync(data);
  for (const part of Object.values(zip.files))
    if (!part.dir && part.name.endsWith('.xml')) assertXmlComplexity(await part.async('string'));
  const mainText = await zip.file('word/document.xml')?.async('string');
  if (!mainText) throw new Error('This file has no Word document.');
  const sourceStructure = readDocxStructure(wordXml(mainText));
  const settingsPart = await readWordSettings(zip);
  sourceStructure.compatibility = readWordCompatibility(settingsPart);
  const settings = settingsPart?.document.documentElement;
  const tabSettings = docxTabSettings(settings);
  const storySources = await readWordStorySources(zip, sourceStructure);
  const paragraphs = new Map<string, Element>();
  const displayProperties = new Map<string, Element>();
  const readNumbering = docxNumbering(await zip.file('word/numbering.xml')?.async('string'));
  const numbering: WordNumbering = { version: 1, paragraphs: [] };
  const styleParagraph = docxStyles(
    await zip.file('word/styles.xml')?.async('string'),
    await zip.file('word/theme/theme1.xml')?.async('string'),
    options.legacyStyleDefaults,
  );
  // Reading-only run styles must never make an absent source styles part
  // appear available to the retained exporter.
  const emptyTemplates = await docxEmptyStoryTemplates(zip, options.legacyFontFeatures, options.legacyParagraphScripts, options.legacyParagraphSpacing);
  const documents = new Map<string, XMLDocument>();
  const copies = new Map<string, XMLDocument>();
  const typography = docxTypography(options);
  const token = `NOFFICE${crypto.randomUUID().replaceAll('-', '')}SOURCE`;
  const paths = [
    'word/document.xml',
    'word/footnotes.xml',
    'word/endnotes.xml',
    ...storySources.parts.map((s) => s.path),
  ];
  let hyphens = 0;
  for (const path of paths) {
    const text = await zip.file(path)?.async('string');
    if (!text) continue;
    const doc = wordXml(text);
    documents.set(path, doc);
    const copy = doc.cloneNode(true) as XMLDocument;
    copies.set(path, copy);
    const originals = descendants(doc, 'p');
    descendants(copy, 'p').forEach((p, index) => {
      const key = `${paths.indexOf(path)}:${index}`;
      paragraphs.set(key, originals[index]);
      const properties = styleParagraph(p);
      displayProperties.set(key, properties);
      const list = readNumbering(properties);
      if (list) {
        if (numbering.paragraphs.length >= 50000) throw Error('The Word document exceeds the supported numbered paragraph limit.');
        numbering.paragraphs.push({ source: key, numbering: list });
      }
      for (const run of descendants(p, 'r')) {
        let owner = run.parentElement;
        while (owner && owner !== p && owner.localName !== 'p') owner = owner.parentElement;
        if (owner === p) typography.retainRunStyle(run, token);
      }
      const run = wElement(copy, 'r');
      const marker = wElement(copy, 't');
      marker.textContent = `${token}${key}END`;
      run.append(marker);
      const props = child(p, 'pPr');
      p.insertBefore(run, props ? props.nextSibling : p.firstChild);
    });
    // Mammoth omits flow breaks. Mark their exact inline positions in the reading copy.
    for (const br of descendants(copy, 'br')) {
      if (!['page', 'column'].includes(val(br, 'type'))) continue;
      const marker = wElement(copy, 't');
      marker.textContent = `${token}${val(br, 'type').toUpperCase()}BREAKEND`;
      br.replaceWith(marker);
    }
    if (!options.legacyHyphens) {
      // Literal U+00AD is not w:softHyphen in Word. Preserve that distinction
      // before Mammoth maps both inputs to the same HTML text character.
      for (const text of descendants(copy, 't')) text.textContent = (text.textContent || '').replace(/\u00ad/g, () => {
        if (++hyphens > 50000) throw Error('The Word document exceeds the supported hyphen limit.');
        return `${token}HYPHENLITERALEND`;
      });
      for (const hyphen of descendants(copy, 'softHyphen')) {
        if (hyphen.childNodes.length || [...hyphen.attributes].some(attribute =>
          attribute.namespaceURI !== 'http://www.w3.org/2000/xmlns/')) continue;
        if (++hyphens > 50000) throw Error('The Word document exceeds the supported hyphen limit.');
        const marker = wElement(copy, 't');marker.textContent = `${token}HYPHENOPTIONALEND`;
        hyphen.replaceWith(marker);
      }
    }
    zip.file(path, new XMLSerializer().serializeToString(copy), { createFolders: false });
  }
  const main = documents.get(paths[0]);
  if (!main) throw new Error('This file has no Word document.');
  const storyBoundaries = await appendWordStoryReading(zip, copies, storySources.parts, token);
  const readingStyles = wordXml(
    (await zip.file('word/styles.xml')?.async('string')) || `<w:styles xmlns:w="${WORD_NS}"/>`,
  );
  if (typography.readingStyles(readingStyles))
    zip.file('word/styles.xml', new XMLSerializer().serializeToString(readingStyles), {
      createFolders: false,
    });
  // Explicit paragraphs inside list items keep source identities when ProseMirror normalizes HTML.
  for (let depth = 1; depth <= 9; depth++) {
    const parents = Array.from({ length: depth - 1 }, () => 'ul|ol > li').join(' > ');
    for (const [kind, tag] of [
      ['unordered', 'ul'],
      ['ordered', 'ol'],
    ])
      typography.styleMap.push(
        `p:${kind}-list(${depth}) => ${parents ? parents + ' > ' : ''}${tag} > li:fresh > p:fresh`,
      );
  }
  const mammoth = await import('mammoth');
  const result = await mammoth.convertToHtml(
    { arrayBuffer: await zip.generateAsync({ type: 'arraybuffer' }) },
    {
      styleMap: typography.styleMap,
      transformDocument: typography.transformDocument,
      includeEmbeddedStyleMap: false,
      externalFileAccess: false,
    },
  );
  const html = new DOMParser().parseFromString(typography.decorate(result.value), 'text/html');
  const storyContainers = splitWordStoryReading(html, storyBoundaries);
  const roots = [html.body, ...storyContainers];
  const all = <T extends Element = Element>(selector: string) =>
    roots.flatMap((root) => [...root.querySelectorAll<T>(selector)]);
  const texts: Text[] = [];
  for (const root of roots) {
    const walker = html.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) texts.push(walker.currentNode as Text);
  }
  const pattern = new RegExp(`${token}(\\d+:\\d+)END`, 'g');
  const ambiguous = new Set<Element>();
  for (const text of texts) {
    text.data = text.data.replace(pattern, (_, key: string) => {
      const parent = text.parentElement?.closest('p,h1,h2,h3,h4,h5,h6');
      if (parent) {
        if (parent.hasAttribute('data-source-paragraph')) ambiguous.add(parent);
        else parent.setAttribute('data-source-paragraph', key);
      }
      return '';
    });
  }
  for (const el of ambiguous) el.removeAttribute('data-source-paragraph');
  for (const text of texts)
    if (text.data.includes('\t') || /(?:PAGE|COLUMN)BREAKEND|HYPHEN(?:OPTIONAL|LITERAL)END/.test(text.data)) {
      const fragment = html.createDocumentFragment();
      text.data.split(new RegExp(`(\\t|${token}(?:PAGE|COLUMN)BREAKEND|${token}HYPHEN(?:OPTIONAL|LITERAL)END)`)).forEach((part) => {
        const hyphen = part === `${token}HYPHENOPTIONALEND` ? 'optional'
          : part === `${token}HYPHENLITERALEND` ? 'literal' : null;
        if (part !== '\t' && part !== `${token}PAGEBREAKEND` && part !== `${token}COLUMNBREAKEND` && !hyphen) {
          fragment.append(html.createTextNode(part));
          return;
        }
        const marker = html.createElement('span');
        if (hyphen) {
          marker.dataset.wordHyphen = hyphen;
          marker.textContent = wordHyphenText(hyphen)!;
        } else if (part === '\t') {
          marker.dataset.wordTab = 'true';
          marker.style.whiteSpace = 'pre';
          marker.style.tabSize = '4';
          marker.textContent = '\t';
        } else {
          const column = part === `${token}COLUMNBREAKEND`;
          marker.setAttribute(column ? 'data-word-column-break' : 'data-word-page-break', 'true');
          marker.style.display = 'block';
          marker.style.breakAfter = column ? 'column' : 'page';
        }
        fragment.append(marker);
      });
      text.replaceWith(fragment);
    }
  // Mammoth merges some deleted paragraphs and drawing text. Preserve those blocks as
  // unchanged opaque content rather than guessing which original paragraph to mutate.
  roots.forEach((root, story) =>
    root.querySelectorAll('p,h1,h2,h3,h4,h5,h6').forEach((el, index) => {
      if (!el.hasAttribute('data-source-paragraph'))
        el.setAttribute(
          'data-source-paragraph',
          `unmapped:${story ? `story${story}:` : ''}${index}`,
        );
    }),
  );
  all('a[href^="#footnote-ref-"],a[href^="#endnote-ref-"]').forEach((link) => {
    const previous = link.previousSibling;
    if (previous?.nodeType === 3 && previous.textContent?.endsWith(' '))
      previous.textContent = previous.textContent.slice(0, -1);
    link.remove();
  });
  all<HTMLElement>('[data-source-paragraph]').forEach((el) => {
    const pr = displayProperties.get(el.dataset.sourceParagraph!);
    if (!pr) return;
    applyDocxParagraphReading(el, pr, tabSettings, options.legacyFontFeatures, options.legacyParagraphScripts, options.legacyParagraphSpacing);
  });
  const body = descendants(main, 'body')[0];
  const section = options.legacySectionDefaults || sourceStructure.compatibility.mode !== 15
    ? liveSectionProperties(body).at(-1)
    : wordFinalSectionProperties(child(body, 'sectPr') || null, sourceStructure.compatibility.mode);
  const size = section ? child(section, 'pgSz') : undefined;
  const margin = section ? child(section, 'pgMar') : undefined;
  const content: WordContent = {
    kind: 'word',
    tabStopsVersion: 1,
    runColorsVersion: options.legacyRunColors ? undefined : 1,
    fontFeaturesVersion: options.legacyFontFeatures ? undefined : 1,
    paragraphScriptsVersion: options.legacyParagraphScripts ? undefined : 1,
    paragraphSpacingVersion: options.legacyParagraphSpacing ? undefined : 1,
    sectionDefaultsVersion: options.legacySectionDefaults ? undefined : 1,
    hyphenVersion: options.legacyHyphens ? undefined : 1,
    numbering,
    lineSpacingVersion: 1,
    fontMetricsVersion: 1,
    kerningVersion: 1,
    paragraphKerningVersion: 1,
    styleDefaultsVersion: options.legacyStyleDefaults ? undefined : 1,
    docxStructure: sourceStructure,
    stories: wordStoriesSchema.parse({
      version: 1,
      evenAndOddHeaders: storySources.evenAndOddHeaders,
      emptyTemplates,
      templateVersion: 4,
      parts: storySources.parts.map((part, index) => ({
        path: part.path,
        kind: part.kind,
        relationshipIds: part.relationshipIds,
        html: storyContainers[index].innerHTML,
      })),
    }),
    html: html.body.innerHTML,
    paper: [Number(val(size, 'w')), Number(val(size, 'h'))].includes(12240) ? 'letter' : 'a4',
    margin:
      val(margin, 'left') === '720' ? 'narrow' : val(margin, 'left') === '2160' ? 'wide' : 'normal',
    orientation: val(size, 'orient') === 'landscape' ? 'landscape' : 'portrait',
  };
  const unsupportedTabs = '[data-word-tab-layout-unsupported="true"]';
  const messages = html.querySelector(unsupportedTabs) || storyContainers.some((story) => story.querySelector(unsupportedTabs))
    ? [...result.messages, { type: 'warning' as const, message: 'This document contains tab definitions that cannot be laid out accurately. The original DOCX is preserved.' }]
    : result.messages;
  if (sourceStructure.compatibility.mode === null || sourceStructure.compatibility.wordPerfectJustification === null)
    messages.push({ type: 'warning', message: 'This document has unresolved Word compatibility settings. Its layout may differ; the original DOCX is preserved.' });
  if (numbering.paragraphs.some(p => p.numbering.status === 'unsupported'))
    messages.push({ type: 'warning', message: 'This document contains numbering definitions that need additional layout support. The original DOCX is preserved.' });
  return { content, messages, paragraphs, documents, tabSettings };
}
