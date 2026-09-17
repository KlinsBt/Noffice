import type { OfficeFile, WorkbookContent } from './model';
import { exportOffice, inspectZip } from './formats';
import { importWorkbook } from './xlsx-import';
import { restructureXlsx } from './xlsx-structure';
import { checkedStructureEdit, type StructureEdit } from './sheet-structure';

export async function editWorkbookStructure(
  file: OfficeFile,
  edit: StructureEdit,
): Promise<WorkbookContent> {
  if (file.content.kind !== 'excel') throw Error('Select a workbook.');
  checkedStructureEdit(edit);
  const previous = file.content;
  const sheet = previous.sheets.find((s) => s.name === edit.sheet);
  if (!sheet) throw Error('The worksheet no longer exists.');
  if (sheet.protected) throw Error('This worksheet is protected.');
  if (edit.at + edit.count > (edit.axis === 'row' ? sheet.rows : sheet.cols))
    throw Error('Select rows or columns within the editable grid.');
  const input = await (await exportOffice(file)).arrayBuffer();
  inspectZip(input);
  const data = await restructureXlsx(input, [edit]);
  inspectZip(data);
  const next = await importWorkbook(data);
  // Re-importing the transformed package gives UI and preservation writer the same metadata,
  // including rules on blank cells. Existing cell edits have already been patched into this copy.
  next.sheets.forEach((s, i) => {
    s.id = previous.sheets[i].id;
    const delta = s.name === edit.sheet ? (edit.action === 'insert' ? edit.count : -edit.count) : 0;
    s.rows = Math.min(
      10000,
      Math.max(s.rows, previous.sheets[i].rows + (edit.axis === 'row' ? delta : 0)),
    );
    s.cols = Math.min(
      256,
      Math.max(s.cols, previous.sheets[i].cols + (edit.axis === 'column' ? delta : 0)),
    );
  });
  // New/model-generated workbooks have no opaque XLSX source to retain. Keep them
  // in model mode so ordinary sheet creation/renaming remains available afterwards.
  if (!previous.xlsxStructureBase && !file.original?.name.toLowerCase().endsWith('.xlsx')) {
    for (const s of next.sheets) {
      delete s.sourcePath;
      for (const filter of s.autoFilters || []) delete filter.sourcePath;
    }
    return next;
  }
  if (data.byteLength > 50000000) throw Error('Structural checkpoints are limited to 50 MB.');
  const bytes = new Uint8Array(data);
  let binary = '';
  for (let i = 0; i < bytes.length; i += 8192)
    binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
  next.xlsxStructureBase = btoa(binary);
  return next;
}
