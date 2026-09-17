import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import spec from './excel-formulas-input.json' with { type: 'json' };
import oracle from './excel-formulas-oracle.json' with { type: 'json' };

/** Independently authored input with recorded desktop results, never engine-generated caches. */
export async function formulaWorkbook() {
  const book = new ExcelJS.Workbook(),
    sheet = book.addWorksheet(spec.sheet);
  for (const [ref, value] of Object.entries(spec.cells))
    sheet.getCell(ref).value =
      typeof value === 'string' && value.startsWith('=') ? { formula: value.slice(1) } : value;
  sheet.addTable({
    name: spec.table.name,
    ref: 'A1',
    headerRow: true,
    columns: spec.table.columns.map((name) => ({ name })),
    rows: [2, 3, 4, 5].map((r) =>
      ['A', 'B', 'C', 'D'].map((c) => spec.cells[(c + r) as keyof typeof spec.cells]),
    ),
  });
  for (const { ref, formula, value, kind } of oracle.cases)
    sheet.getCell(ref).value = {
      formula: formula.slice(1),
      result: kind === 'error' ? { error: value as ExcelJS.CellErrorValue['error'] } : value,
    };
  sheet.getCell('L2').value = { formula: 'J12*2', result: 240 };
  sheet.getCell('L3').value = { formula: 'IFERROR(UNSUPPORTED_FUNCTION(B2),0)', result: 987 };
  const zip = await JSZip.loadAsync(await book.xlsx.writeBuffer());
  const xml = await zip.file('xl/workbook.xml')!.async('string');
  const escape = (s: string) =>
    s.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');
  zip.file(
    'xl/workbook.xml',
    xml.replace(
      '<calcPr',
      '<definedNames>' +
        Object.entries(spec.names)
          .map(
            ([name, formula]) =>
              `<definedName name="${escape(name)}">${escape(formula)}</definedName>`,
          )
          .join('') +
        '</definedNames><calcPr',
    ),
  );
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
