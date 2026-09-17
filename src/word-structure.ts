import JSZip from 'jszip';
import type { OfficeFile, WordContent } from './model';
import { inspectZip, sanitizeWordContent } from './formats';
import { wordXml, readDocx } from './docx-import';
import { readDocxStructure } from './docx-sections';
import { contentFingerprint } from './office-preservation';
import { hydrateParagraphFonts } from './word-font-migration';
import { hydrateParagraphKerning, hydrateParagraphMarkKerning } from './word-kerning-migration';
import { hydrateStoryParagraphKerning } from './word-story-kerning-migration';
import { hydrateWordTabMetadata } from './word-tab-migration';
import { authoredWordSection } from './word-authored-section';
import { authoredWordStoryTemplates } from './word-authored-stories';
import { hydrateWordFontFeatures } from './word-font-feature-migration';
import { hydrateWordParagraphScripts } from './word-script-migration';
import { hydrateWordParagraphSpacing } from './word-paragraph-spacing-migration';
import { readWordSettings } from './docx-settings';
import { readWordCompatibility } from './word-compatibility';
import { usesWordSectionDefaults } from './word-section-defaults';
import { hydrateWordSectionDefaults } from './word-section-defaults-migration';
import { hydrateWordHyphens } from './word-hyphen-migration';

export function needsWordStructure(file: OfficeFile) {
  if (file.content.kind === 'word' && !file.original && file.content.tabStopsVersion !== 1 &&
      file.content.stories?.parts.length && authoredWordSection(file.content)) return true;
  return (
    file.content.kind === 'word' &&
    (!file.content.docxStructure ||
      !file.content.docxStructure.compatibility ||
      file.content.lineSpacingVersion !== 1 ||
      file.content.fontMetricsVersion !== 1 ||
      file.content.kerningVersion !== 1 ||
      file.content.paragraphKerningVersion !== 1 ||
      file.content.styleDefaultsVersion !== 1 ||
      file.content.tabStopsVersion !== 1 ||
      file.content.runColorsVersion !== 1 ||
      file.content.fontFeaturesVersion !== 1 ||
      file.content.paragraphScriptsVersion !== 1 ||
      file.content.paragraphSpacingVersion !== 1 ||
      file.content.hyphenVersion !== 1 ||
      !file.content.numbering ||
      (file.content.sectionDefaultsVersion !== 1 && usesWordSectionDefaults(file.content.docxStructure)) ||
      !file.content.stories?.emptyTemplates ||
      file.content.stories.templateVersion !== 4) &&
    !!file.original?.name.toLowerCase().endsWith('.docx')
  );
}
/** Add original section identities to legacy snapshots without replacing their edited HTML. */
export async function hydrateWordStructure(file: OfficeFile): Promise<OfficeFile> {
  if (!needsWordStructure(file) || file.content.kind !== 'word') return file;
  if (file.original && file.content.hyphenVersion !== 1) {
    inspectZip(file.original.data);
    const content = hydrateWordHyphens(file.content, sanitizeWordContent((await readDocx(file.original.data)).content));
    const unchanged = file.original.contentFingerprint === await contentFingerprint(file.content);
    file = { ...file, content, original: unchanged
      ? { ...file.original, contentFingerprint: await contentFingerprint(content) } : file.original };
    if (!needsWordStructure(file)) return file;
    if (file.content.kind !== 'word') return file;
  }
  if (!file.original) {
    const templates = authoredWordStoryTemplates(file.content);
    if (!templates) throw Error('This saved document has no source for its tab settings. Export a backup before reopening it.');
    const source: WordContent = {
      ...file.content,
      html: '<p data-word-default-tab="720"></p>',
      stories: { version: 1, evenAndOddHeaders: false, parts: [], emptyTemplates: templates, templateVersion: 4 },
    };
    const content = hydrateWordTabMetadata(file.content, source);
    content.stories = { ...content.stories!, emptyTemplates: templates, templateVersion: 4 };
    return { ...file, content };
  }
  inspectZip(file.original!.data);
  const zip = await JSZip.loadAsync(file.original!.data);
  const main = await zip.file('word/document.xml')?.async('string');
  if (!main) throw Error('The retained Word document has no main part.');
  let html = file.content.html;
  let source: Awaited<ReturnType<typeof readDocx>> | undefined;
  const legacyFontFeatures = file.content.fontFeaturesVersion !== 1;
  const legacyParagraphScripts = file.content.paragraphScriptsVersion !== 1;
  const legacyParagraphSpacing = file.content.paragraphSpacingVersion !== 1;
  const legacySectionDefaults = file.content.sectionDefaultsVersion !== 1;
  const readSource = async () => (source ||= await readDocx(file.original!.data,
    { legacyFontFeatures, legacyParagraphScripts, legacyParagraphSpacing }));
  let restoredDefaults = false;
  let runColorsVersion = file.content.runColorsVersion;
  const unchanged =
    file.original!.contentFingerprint &&
    file.original!.contentFingerprint === (await contentFingerprint(file.content));
  if (file.content.styleDefaultsVersion !== 1) {
    const currentSource = sanitizeWordContent((file.content.runColorsVersion === 1
      ? await readSource()
      : await readDocx(file.original!.data, { legacyRunColors: true, legacyFontFeatures, legacyParagraphScripts, legacyParagraphSpacing })).content);
    const legacySource = sanitizeWordContent(
      (await readDocx(file.original!.data, { legacyStyleDefaults: true, legacyRunColors: file.content.runColorsVersion !== 1, legacyFontFeatures, legacyParagraphScripts, legacyParagraphSpacing })).content,
    );
    const appearance = (content: typeof currentSource) =>
      JSON.stringify({
        html: content.html,
        stories: content.stories?.parts.map((p) => ({ path: p.path, html: p.html })),
      });
    if (appearance(currentSource) !== appearance(legacySource)) {
      if (!unchanged)
        throw Error(
          'This saved document uses older Word style defaults and contains edits that cannot be migrated safely. Export a Noffice backup to preserve those edits, then reopen the original DOCX as a separate document.',
        );
      html = currentSource.html;
      restoredDefaults = true;
    }
  }
  if (runColorsVersion !== 1) {
    const currentSource = sanitizeWordContent((await readSource()).content);
    const legacySource = sanitizeWordContent((await readDocx(file.original!.data, { legacyRunColors: true, legacyFontFeatures, legacyParagraphScripts, legacyParagraphSpacing })).content);
    const appearance = (content: WordContent) => JSON.stringify({
      html: content.html,
      stories: content.stories?.parts.map((part) => ({ path: part.path, html: part.html })),
    });
    if (unchanged) {
      html = currentSource.html;
      restoredDefaults = true;
      runColorsVersion = 1;
    } else if (appearance(currentSource) === appearance(legacySource)) {
      runColorsVersion = 1;
    }
    // Edited legacy snapshots keep their original reading semantics. Without a
    // run identity map, missing color cannot be distinguished from a user edit.
    // Their retained exporter must compare against the same legacy baseline.
  }
  if (
    file.content.lineSpacingVersion !== 1 ||
    file.content.fontMetricsVersion !== 1 ||
    file.content.kerningVersion !== 1 ||
    file.content.paragraphKerningVersion !== 1
  ) {
    const current = new DOMParser().parseFromString(html, 'text/html');
    if (
      [...current.querySelectorAll<HTMLElement>('[data-source-paragraph]')].some(
        (p) =>
          !p.style.lineHeight ||
          (file.content.kind === 'word' &&
            (file.content.fontMetricsVersion !== 1 ||
              file.content.kerningVersion !== 1 ||
              file.content.paragraphKerningVersion !== 1)),
      )
    ) {
      const source = await readSource();
      const previous = new DOMParser().parseFromString(source.content.html, 'text/html');
      const sourceSpacing = new Map(
        [...previous.querySelectorAll<HTMLElement>('[data-source-paragraph]')].map((p) => [
          p.dataset.sourceParagraph,
          { height: p.style.lineHeight, minimum: p.dataset.wordLineRule === 'atLeast', element: p },
        ]),
      );
      let changed = false;
      for (const p of current.querySelectorAll<HTMLElement>('[data-source-paragraph]')) {
        const spacing = sourceSpacing.get(p.dataset.sourceParagraph);
        // Earlier UI commands wrote explicit values. Missing spacing is an old
        // importer omission, not a request to overwrite the original spacing.
        if (file.content.lineSpacingVersion !== 1 && spacing?.height && !p.style.lineHeight) {
          p.style.lineHeight = spacing.height;
          if (spacing.minimum) p.dataset.wordLineRule = 'atLeast';
          changed = true;
        }
        if (spacing && file.content.fontMetricsVersion !== 1) {
          hydrateParagraphFonts(p, spacing.element);
          changed = true;
        }
        if (spacing && file.content.paragraphKerningVersion !== 1) {
          hydrateParagraphMarkKerning(p, spacing.element);
          changed = true;
        }
        if (spacing && file.content.kerningVersion !== 1) {
          hydrateParagraphKerning(p, spacing.element);
          changed = true;
        }
      }
      if (changed) html = current.body.innerHTML;
    }
  }
  const stories = restoredDefaults
    ? sanitizeWordContent((await readSource()).content).stories
    : file.content.stories || sanitizeWordContent((await readSource()).content).stories;
  let migratedStories =
    stories &&
    (!stories.emptyTemplates ||
      stories.templateVersion !== 4 ||
      file.content.tabStopsVersion !== 1 ||
      file.content.paragraphKerningVersion !== 1)
      ? {
          ...stories,
          templateVersion: 4 as const,
          emptyTemplates: (await readSource()).content.stories!.emptyTemplates,
        }
      : stories;
  if (migratedStories && file.content.paragraphKerningVersion !== 1)
    migratedStories = hydrateStoryParagraphKerning(
      migratedStories,
      (await readSource()).content.stories!,
    );
  let content: WordContent = {
    ...file.content,
    html,
    numbering: file.content.numbering || (await readSource()).content.numbering,
    stories: migratedStories,
    lineSpacingVersion: 1 as const,
    fontMetricsVersion: 1 as const,
    kerningVersion: 1 as const,
    paragraphKerningVersion: 1 as const,
    styleDefaultsVersion: 1 as const,
    runColorsVersion,
    docxStructure: {
      ...(file.content.docxStructure || readDocxStructure(wordXml(main))),
      compatibility: file.content.docxStructure?.compatibility || readWordCompatibility(await readWordSettings(zip)),
    },
  };
  if (file.content.tabStopsVersion !== 1)
    content = { ...content, ...hydrateWordTabMetadata(content, (await readSource()).content) };
  if (file.content.fontFeaturesVersion !== 1)
    content = hydrateWordFontFeatures(content, sanitizeWordContent((await readDocx(file.original!.data)).content));
  if (legacyParagraphScripts)
    content = hydrateWordParagraphScripts(content, sanitizeWordContent((await readDocx(file.original!.data)).content));
  if (legacyParagraphSpacing)
    content = hydrateWordParagraphSpacing(content,
      sanitizeWordContent((await readDocx(file.original!.data)).content),
      sanitizeWordContent((await readDocx(file.original!.data, { legacyParagraphSpacing: true })).content));
  if (legacySectionDefaults) {
    if (usesWordSectionDefaults(content.docxStructure))
      content = hydrateWordSectionDefaults(content,
        sanitizeWordContent((await readDocx(file.original!.data)).content),
        sanitizeWordContent((await readDocx(file.original!.data, { legacySectionDefaults: true })).content));
    else content.sectionDefaultsVersion = 1;
  }
  return {
    ...file,
    content,
    original:
      unchanged && (content.html !== file.content.html || content.stories !== file.content.stories
        || content.paper !== file.content.paper || content.margin !== file.content.margin
        || content.orientation !== file.content.orientation)
        ? { ...file.original!, contentFingerprint: await contentFingerprint(content) }
        : file.original,
  };
}
