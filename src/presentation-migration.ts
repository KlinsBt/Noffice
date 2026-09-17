import type { OfficeFile, DeckContent } from './model';
import { readPresentation } from './pptx-import';
import { inspectZip } from './formats';
import { contentFingerprint } from './office-preservation';

export function needsPresentationText(file: OfficeFile) {
  return (
    file.content.kind === 'powerpoint' &&
    (file.content.textModelVersion !== 1 ||
      file.content.groupModelVersion !== 1 ||
      file.content.slides.some((s) =>
        s.elements.some((e) => e.sourceGroupTransform && !e.sourceGroupTransform.dimensions),
      )) &&
    !!file.original?.name.toLowerCase().endsWith('.pptx')
  );
}
export async function hydratePresentationText(file: OfficeFile): Promise<OfficeFile> {
  if (!needsPresentationText(file) || file.content.kind !== 'powerpoint') return file;
  inspectZip(file.original!.data);
  const old = await readPresentation(
    file.original!.data,
    file.content.textModelVersion !== 1,
    file.content.groupModelVersion !== 1,
  );
  const fresh = await readPresentation(file.original!.data);
  const unchanged = (await contentFingerprint(file.content)) === file.original!.contentFingerprint;
  const content: DeckContent = {
    ...file.content,
    textModelVersion: 1,
    groupModelVersion: 1,
    slides: file.content.slides.map((slide) => {
      const source = fresh.slides.find((s) => s.sourcePath === slide.sourcePath);
      const previous = old.slides.find((s) => s.sourcePath === slide.sourcePath);
      if (!source || !previous) return slide;
      return {
        ...slide,
        sourceLayoutPath: source.sourceLayoutPath,
        sourceMasterPath: source.sourceMasterPath,
        sourceThemePath: source.sourceThemePath,
        elements: slide.elements.map((item) => {
          const before = previous.elements.find((e) => e.sourceShapeId === item.sourceShapeId);
          const after = source.elements.find((e) => e.sourceShapeId === item.sourceShapeId);
          if (!before || !after || !item.sourceShapeId) return item;
          const next = {
            ...item,
            sourceText: after.sourceText,
            sourcePlaceholder: after.sourcePlaceholder,
            sourceGroupIds: after.sourceGroupIds,
            sourceGroupTransform: after.sourceGroupTransform,
          };
          for (const key of [
            'text',
            'fontSize',
            'fontFamily',
            'bold',
            'italic',
            'underline',
            'color',
            'align',
            'x',
            'y',
            'w',
            'h',
            'rotation',
            'flipH',
            'flipV',
          ] as const) {
            if (item[key] === before[key]) Object.assign(next, { [key]: after[key] });
          }
          return next;
        }),
      };
    }),
  };
  return {
    ...file,
    content,
    original: unchanged
      ? { ...file.original!, contentFingerprint: await contentFingerprint(content) }
      : file.original,
  };
}
