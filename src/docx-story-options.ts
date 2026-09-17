import type JSZip from 'jszip';
import type { WordContent } from './model';
import { WORD_NS, child, wElement } from './docx-import';
import { liveSectionProperties } from './docx-sections';
import { wordStoriesSchema } from './word-stories';
import { readWordSettings } from './docx-settings';
import { liveWordSectionIds } from './word-section-breaks';

/** Apply validated overrides to the exporter's private package copy. Story
 * parts and relationship identities are unchanged by page-option commands. */
export async function writeWordStoryOptions(
  zip: JSZip,
  body: Element,
  original: WordContent,
  current: WordContent,
  live = false,
) {
  const stories = current.stories && wordStoriesSchema.parse(current.stories);
  const properties = liveSectionProperties(body);
  let changed = false;
  let settingsChanged = false;
  for (const options of stories?.sectionOptions || []) {
    // Clone source properties before applying any live identity's overrides;
    // otherwise resetting a clone can inherit its sibling's current options.
    if (!live && current.sectionState?.version === 2) continue;
    const index =
      live && current.sectionState
        ? liveWordSectionIds(current.sectionState).indexOf(options.sectionId)
        : (original.docxStructure?.sections.findIndex((s) => s.id === options.sectionId) ?? -1);
    const section = index >= 0 && properties[index];
    if (!section) throw Error('The header/footer options refer to an unavailable source section.');
    if (options.differentFirstPage !== undefined) {
      const prior = child(section, 'titlePg');
      if (!options.differentFirstPage) prior?.remove();
      else if (prior) prior.setAttributeNS(WORD_NS, 'w:val', '1');
      else {
        const anchor = [...section.children].find(
          (e) =>
            e.namespaceURI === WORD_NS &&
            [
              'textDirection',
              'bidi',
              'rtlGutter',
              'docGrid',
              'printerSettings',
              'sectPrChange',
            ].includes(e.localName),
        );
        section.insertBefore(wElement(section.ownerDocument, 'titlePg'), anchor || null);
      }
      changed = true;
    }
    for (const [option, attribute] of [
      ['headerDistance', 'header'],
      ['footerDistance', 'footer'],
    ] as const) {
      if (options[option] === undefined) continue;
      const margins = child(section, 'pgMar');
      if (!margins)
        throw Error('Editing header/footer distances requires explicit section margins.');
      margins.setAttributeNS(WORD_NS, `w:${attribute}`, String(options[option]));
      changed = true;
    }
  }
  if (!!stories?.evenAndOddHeaders !== !!original.stories?.evenAndOddHeaders) {
    const part = await readWordSettings(zip);
    if (!part) throw Error('Editing odd/even headers requires retained document settings.');
    const settings = part.document;
    const root = settings.documentElement;
    if (root.namespaceURI !== WORD_NS || root.localName !== 'settings')
      throw Error('Invalid Word settings.');
    const prior = child(root, 'evenAndOddHeaders');
    if (!stories?.evenAndOddHeaders) prior?.remove();
    else if (prior) prior.setAttributeNS(WORD_NS, 'w:val', '1');
    else {
      const anchor = [...root.children].find(
        (e) =>
          e.namespaceURI === WORD_NS &&
          [
            'bookFoldRevPrinting',
            'bookFoldPrinting',
            'bookFoldPrintingSheets',
            'drawingGridHorizontalSpacing',
            'drawingGridVerticalSpacing',
            'displayHorizontalDrawingGridEvery',
            'displayVerticalDrawingGridEvery',
            'doNotUseMarginsForDrawingGridOrigin',
            'drawingGridHorizontalOrigin',
            'drawingGridVerticalOrigin',
            'doNotShadeFormData',
            'noPunctuationKerning',
            'characterSpacingControl',
            'printTwoOnOne',
            'strictFirstAndLastChars',
            'noLineBreaksAfter',
            'noLineBreaksBefore',
            'savePreviewPicture',
            'doNotValidateAgainstSchema',
            'saveInvalidXml',
            'ignoreMixedContent',
            'alwaysShowPlaceholderText',
            'doNotDemarcateInvalidXml',
            'saveXmlDataOnly',
            'useXSLTWhenSaving',
            'saveThroughXslt',
            'showXMLTags',
            'alwaysMergeEmptyNamespace',
            'updateFields',
            'hdrShapeDefaults',
            'footnotePr',
            'endnotePr',
            'compat',
            'docVars',
            'rsids',
            'mathPr',
            'themeFontLang',
            'clrSchemeMapping',
          ].includes(e.localName),
      );
      root.insertBefore(wElement(settings, 'evenAndOddHeaders'), anchor || null);
    }
    zip.file(part.path, new XMLSerializer().serializeToString(settings), {
      createFolders: false,
    });
    settingsChanged = true;
  }
  return { bodyChanged: changed, settingsChanged };
}
