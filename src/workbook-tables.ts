import type { OfficeFile } from './model';
import { inspectZip } from './formats';
import { importWorkbook } from './xlsx-import';
import { restoreTableMetadata, legacyCalculatedResizeBlock } from './sheet-tables';

export function needsTableMetadata(file: OfficeFile) {
  return (
    file.content.kind === 'excel' &&
    file.content.sheets.every((s) => !!s.sourcePath) &&
    (!!file.content.xlsxStructureBase || !!file.original?.name.toLowerCase().endsWith('.xlsx')) &&
    file.content.sheets.some(
      (s) =>
        !s.arrayFormulas ||
        !s.nameDefinitions ||
        s.nameDefinitions.some((d) => !d.referenceOrigin) ||
        !s.tableTheme ||
        s.tables?.some(
          (t) =>
            !t.sourcePath ||
            !t.calculatedColumns ||
            !t.totalLabels ||
            !t.totalFormulas ||
            t.resizeBlocked === legacyCalculatedResizeBlock,
        ),
    )
  );
}
export async function hydrateTableMetadata(file: OfficeFile): Promise<OfficeFile> {
  if (!needsTableMetadata(file) || file.content.kind !== 'excel') return file;
  const data = file.content.xlsxStructureBase
    ? Uint8Array.from(atob(file.content.xlsxStructureBase), (c) => c.charCodeAt(0)).buffer
    : file.original!.data;
  inspectZip(data);
  const next = { ...file, content: restoreTableMetadata(file.content, await importWorkbook(data)) };
  if (needsTableMetadata(next))
    throw Error(
      'This saved table no longer matches its source metadata. Export a Noffice backup and reopen the original XLSX.',
    );
  return next;
}
