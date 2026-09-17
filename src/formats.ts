import { readPresentation } from './pptx-import';
import { importWorkbook, parseXML, elements, workbookSheets } from './xlsx-import';
import { writeFilterColumns } from './sheet-filters';
import { writeWorksheetLinks } from './xlsx-links';
import { writeValidationChanges } from './xlsx-validation';
import { writeSheetTables } from './xlsx-tables';
import { tableAt } from './sheet-tables';
import { contentFingerprint } from './office-preservation';
import { exportAddedPresentationObjects, exportRetainedPresentation } from './pptx-preserve';
import { exportRetainedWorkbook } from './xlsx-preserve';
import { defaultFont } from './fonts';
import { wordDefaults } from './word-defaults';
import { resolveWordSections } from './word-section-layout';
import { validateSectionState } from './word-section-breaks';
import { wordStoriesSchema } from './word-stories';
import { wordLineSpacing } from './word-line-spacing';
import { wordParagraphSpace, wordParagraphSpaceXml } from './word-paragraph-spacing';
import { wordKerningValue } from './word-kerning';
import { WORD_2010_NS, wordFontFeaturesValue, wordLigatureNames } from './word-font-features';
import { wordScriptValue, type WordScript } from './word-script';
import type { IRunOptions, IParagraphOptions, IContext } from 'docx';
import { wordRunSession } from './word-edit-run';
import { authoredWordNumbering } from './docx-authored-numbering';
import { readDocx } from './docx-import';
import DOMPurify from 'dompurify';
import JSZip from 'jszip';
import { z } from 'zod';
import {
  contentSchema,
  newFile,
  uid,
  type OfficeFile,
  type Sheet,
  type Cell,
  type WordContent,
} from './model';
import { calculator, coordinates, address } from './formulas';

export const MAX_FILE_BYTES = 30 * 1024 * 1024;
export function sanitizeHTML(html: string) {
  const clean = DOMPurify.sanitize(html, {
    USE_PROFILES: { html: true },
    FORBID_TAGS: ['style', 'iframe', 'form', 'input', 'video', 'audio'],
    FORBID_ATTR: ['srcset'],
  });
  const dom = new DOMParser().parseFromString(clean, 'text/html');
  dom.querySelectorAll('img').forEach((img) => {
    if (!/^data:image\/(png|jpeg|gif|webp);base64,/i.test(img.src)) img.remove();
  });
  dom.querySelectorAll('[style]').forEach((el) => {
    if (/url\s*\(|expression|@import/i.test(el.getAttribute('style') || ''))
      el.removeAttribute('style');
  });
  dom.querySelectorAll('a').forEach((a) => {
    if (!/^(https?:|mailto:|#)/i.test(a.getAttribute('href') || '')) a.removeAttribute('href');
    a.setAttribute('rel', 'noopener noreferrer');
  });
  return dom.body.innerHTML;
}
export function sanitizeWordContent(content: WordContent): WordContent {
  return {
    ...content,
    html: sanitizeHTML(content.html),
    ...(content.stories
      ? {
          stories: {
            ...content.stories,
            ...(content.stories.emptyTemplates
              ? {
                  emptyTemplates: {
                    header:
                      content.stories.emptyTemplates.header &&
                      sanitizeHTML(content.stories.emptyTemplates.header),
                    footer:
                      content.stories.emptyTemplates.footer &&
                      sanitizeHTML(content.stories.emptyTemplates.footer),
                  },
                }
              : {}),
            parts: content.stories.parts.map((part) => ({
              ...part,
              html: sanitizeHTML(part.html),
            })),
          },
        }
      : {}),
  };
}
export function download(
  data: Blob | ArrayBuffer | string,
  name: string,
  type = 'application/octet-stream',
) {
  const blob = data instanceof Blob ? data : new Blob([data], { type });
  const url = URL.createObjectURL(blob),
    a = document.createElement('a');
  a.href = url;
  a.download = name.replace(/[<>:"/\\|?*\x00-\x1f]/g, '_');
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 30000);
}
export async function imageData(file: File): Promise<string> {
  if (
    !['image/png', 'image/jpeg', 'image/gif', 'image/webp'].includes(file.type) ||
    file.size > 8 * 1024 * 1024
  )
    throw new Error('Choose a PNG, JPEG, GIF, or WebP image smaller than 8 MB.');
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(new Error('Could not read this image.'));
    reader.readAsDataURL(file);
  });
}
export function inspectZip(data: ArrayBuffer) {
  const v = new DataView(data);
  let end = data.byteLength - 22;
  for (; end >= Math.max(0, data.byteLength - 65557); end--)
    if (v.getUint32(end, true) === 0x06054b50) break;
  if (end < 0 || end < data.byteLength - 65557)
    throw new Error('This file is not a supported Office ZIP package.');
  const count = v.getUint16(end + 10, true),
    start = v.getUint32(end + 16, true);
  if (count > 5000 || start >= end || v.getUint16(end + 4, true) || v.getUint16(end + 6, true))
    throw new Error('This archive is too large or uses an unsupported ZIP format.');
  let offset = start,
    expanded = 0;
  for (let i = 0; i < count; i++) {
    if (offset + 46 > end || v.getUint32(offset, true) !== 0x02014b50)
      throw new Error('The archive directory is damaged.');
    if (v.getUint16(offset + 8, true) & 1)
      throw new Error('Password-protected files are not supported.');
    const size = v.getUint32(offset + 24, true);
    expanded += size;
    if (size > 32 * 1024 * 1024 || expanded > 100 * 1024 * 1024)
      throw new Error('The expanded file exceeds the 100 MB import limit.');
    offset +=
      46 +
      v.getUint16(offset + 28, true) +
      v.getUint16(offset + 30, true) +
      v.getUint16(offset + 32, true);
  }
  if (offset > end) throw new Error('The archive directory is damaged.');
}
export function parseCSV(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [],
    field = '',
    quoted = false;
  const input = text.replace(/^\uFEFF/, '');
  for (let i = 0; i < input.length; i++) {
    const c = input[i];
    if (c === '"') {
      if (quoted && input[i + 1] === '"') {
        field += '"';
        i++;
      } else if (quoted || field === '') quoted = !quoted;
      else field += c;
    } else if (c === ',' && !quoted) {
      row.push(field);
      field = '';
    } else if ((c === '\n' || c === '\r') && !quoted) {
      row.push(field);
      rows.push(row);
      field = '';
      row = [];
      if (c === '\r' && input[i + 1] === '\n') i++;
    } else field += c;
  }
  if (quoted) throw new Error('The CSV contains an unclosed quoted field.');
  if (field || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}
export function writeCSV(sheet: Sheet, sheets: Sheet[]): string {
  const calc = calculator(sheets),
    refs = Object.keys(sheet.cells).map(coordinates);
  const rows = Math.max(1, ...refs.map(([r]) => r + 1)),
    cols = Math.max(1, ...refs.map(([, c]) => c + 1));
  return Array.from({ length: rows }, (_, r) =>
    Array.from({ length: cols }, (_, c) => {
      let value = String(calc(sheet, address(r, c)));
      // Prevent external spreadsheet applications from interpreting text as a formula.
      if (/^[=+@\-\t\r]/.test(value) && !Number.isFinite(Number(value))) value = "'" + value;
      return /[",\r\n]/.test(value) ? `"${value.replaceAll('"', '""')}"` : value;
    }).join(','),
  ).join('\r\n');
}
export async function importFile(input: File): Promise<OfficeFile> {
  if (input.size > MAX_FILE_BYTES) throw new Error('Choose a file smaller than 30 MB.');
  const ext = input.name.split('.').pop()?.toLowerCase(),
    name = input.name.replace(/\.[^.]+$/, '');
  if (ext === 'noffice') {
    const schema = z.object({
      format: z.literal('noffice'),
      version: z.literal(1),
      name: z.string().min(1).max(200),
      content: contentSchema,
      warnings: z.array(z.string()).default([]),
      original: z
        .object({
          name: z.string(),
          base64: z.string(),
          contentFingerprint: z
            .string()
            .regex(/^[a-f0-9]{64}$/)
            .optional(),
        })
        .optional(),
    });
    const data = schema.parse(JSON.parse(await input.text()));
    if (data.content.kind === 'word') data.content = sanitizeWordContent(data.content);
    if (data.content.kind === 'excel')
      data.content.sheets.forEach((sheet) =>
        sheet.images?.forEach((image) => {
          if (!/^data:image\/(png|jpeg|gif|webp);base64,/i.test(image.src))
            throw new Error('Backup contains an unsupported spreadsheet image source.');
        }),
      );
    if (data.content.kind === 'powerpoint')
      data.content.slides.forEach((s) =>
        s.elements.forEach((e) => {
          if (e.src && !/^data:image\/(png|jpeg|gif|webp);base64,/i.test(e.src))
            throw new Error('Backup contains an unsupported image source.');
        }),
      );
    const file = newFile(data.content.kind, data.name, data.content);
    file.warnings = data.warnings;
    if (data.original)
      file.original = {
        name: data.original.name,
        data: Uint8Array.from(atob(data.original.base64), (c) => c.charCodeAt(0)).buffer,
        contentFingerprint: data.original.contentFingerprint,
      };
    if (file.content.kind === 'word') {
      const { hydrateWordStructure } = await import('./word-structure');
      return hydrateWordStructure(file);
    }
    if (file.content.kind === 'powerpoint') {
      const { hydratePresentationText } = await import('./presentation-migration');
      return hydratePresentationText(file);
    }
    return file;
  }
  if (ext === 'txt' || ext === 'html' || ext === 'htm') {
    const text = await input.text();
    const escaped = text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
    return newFile('word', name, {
      kind: 'word',
      html:
        ext === 'txt'
          ? escaped
              .split(/\r?\n/)
              .map((line) => `<p>${line}</p>`)
              .join('')
          : sanitizeHTML(text),
      paper: 'a4',
      margin: 'normal',
    });
  }
  if (ext === 'csv') {
    const rows = parseCSV(await input.text());
    if (rows.length > 10000 || rows.some((r) => r.length > 256))
      throw new Error('CSV imports support up to 10,000 rows and 256 columns.');
    const cells: Sheet['cells'] = {};
    rows.forEach((row, r) =>
      row.forEach((value, c) => {
        if (value) cells[address(r, c)] = { value: /^[=+@]/.test(value) ? "'" + value : value };
      }),
    );
    return newFile('excel', name, {
      kind: 'excel',
      sheets: [
        {
          id: uid(),
          name: 'Sheet 1',
          cells,
          rows: Math.max(100, rows.length),
          cols: Math.max(26, ...rows.map((r) => r.length)),
        },
      ],
    });
  }
  if (!['docx', 'xlsx', 'pptx'].includes(ext || ''))
    throw new Error(
      'Open DOCX, XLSX, PPTX, CSV, TXT, HTML, or a .noffice backup. Legacy and macro-enabled files are not supported yet.',
    );
  const data = await input.arrayBuffer();
  inspectZip(data);
  let file: OfficeFile;
  if (ext === 'docx') {
    const result = await readDocx(data);
    file = newFile('word', name, sanitizeWordContent(result.content));
    file.warnings = [
      'DOCX exports retain the original package and patch supported paragraph edits. Headers, notes, review structures, pictures and section settings remain in the file. Complex structural edits that cannot yet be preserved are reported at export. Browser layout and advanced editing remain incomplete; the original remains available to download.',
    ];
    if (result.messages.length)
      file.warnings.push(
        `The DOCX converter reported ${result.messages.length} unsupported or adjusted items.`,
      );
  } else if (ext === 'xlsx') {
    file = newFile('excel', name, await importWorkbook(data));
    file.warnings = [
      'Imported XLSX exports preserve the original package and patch supported cell edits. Unchanged exports retain the original bytes. Charts, drawings, rich text and some workbook features are retained in the file but not yet fully editable on screen. Unsupported calculations are identified separately. Keep the source sheet names and order when exporting imported workbooks.',
    ];
  } else {
    file = newFile('powerpoint', name, await readPresentation(data));
    file.warnings = [
      'PPTX exports retain original parts during supported text, formatting, geometry, background and existing-note edits. Masters, charts, media and timing remain in the file. Unsupported structural edits are reported at export. Browser layout, rich text and playback remain incomplete.',
    ];
    if (
      file.content.kind === 'powerpoint' &&
      file.content.slides.some((s) =>
        s.elements.some((e) => e.sourceGroupTransform && !e.sourceGroupTransform.projected),
      )
    )
      file.warnings.push(
        'Some grouped objects have skewed transforms that cannot yet be displayed accurately. Their original transforms are retained; geometry editing is unavailable.',
      );
  }
  contentSchema.parse(file.content);
  file.original = {
    name: input.name,
    data,
    contentFingerprint: file.kind === 'excel' ? undefined : await contentFingerprint(file.content),
  };
  return file;
}

export function nativeBackup(file: OfficeFile): string {
  let original: { name: string; base64: string; contentFingerprint?: string } | undefined;
  if (file.original) {
    const bytes = new Uint8Array(file.original.data);
    let binary = '';
    for (let i = 0; i < bytes.length; i += 8192)
      binary += String.fromCharCode(...bytes.subarray(i, i + 8192));
    original = {
      name: file.original.name,
      base64: btoa(binary),
      contentFingerprint: file.original.contentFingerprint,
    };
  }
  return JSON.stringify({
    format: 'noffice',
    version: 1,
    name: file.name,
    content: file.content,
    warnings: file.warnings,
    original,
  });
}
/** Unchanged package bytes need no renderer, fonts or derived layout. The full
 * content fingerprint and ZIP validation still guard this preservation path. */
export async function unchangedOfficeOriginal(file: OfficeFile): Promise<Blob | undefined> {
  const originalExtension =
    file.content.kind === 'word' ? 'docx' : file.content.kind === 'powerpoint' ? 'pptx' : null;
  if (
    originalExtension &&
    file.original?.name.toLowerCase().endsWith('.' + originalExtension) &&
    file.original.contentFingerprint === (await contentFingerprint(file.content))
  ) {
    inspectZip(file.original.data);
    return new Blob([file.original.data], {
      type:
        originalExtension === 'docx'
          ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
          : 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    });
  }
  return undefined;
}
export async function exportOffice(
  file: OfficeFile,
  wordLayout?: import('./word-export-layout').WordExportLayout,
): Promise<Blob> {
  if (file.content.kind === 'powerpoint') {
    const { hydratePresentationText } = await import('./presentation-migration');
    file = await hydratePresentationText(file);
  }
  const original = await unchangedOfficeOriginal(file);
  if (original) return original;
  if (file.content.kind === 'powerpoint' && file.original?.name.toLowerCase().endsWith('.pptx')) {
    inspectZip(file.original.data);
    const preserved = await exportAddedPresentationObjects(file);
    if (preserved) return preserved;
    return exportRetainedPresentation(file);
  }
  if (file.content.kind === 'excel') {
    if (file.content.xlsxStructureBase) {
      inspectZip(
        Uint8Array.from(atob(file.content.xlsxStructureBase), (c) => c.charCodeAt(0)).buffer,
      );
      return exportRetainedWorkbook(file);
    }
    if (
      file.original?.name.toLowerCase().endsWith('.xlsx') &&
      !file.content.sheets.every((s) => s.sourcePath)
    )
      throw new Error(
        'This workbook was imported before preservation support. Keep a Noffice backup of your edits and reopen the original XLSX to enable the new importer.',
      );
    if (
      file.original?.name.toLowerCase().endsWith('.xlsx') &&
      file.content.sheets.every((s) => s.sourcePath)
    ) {
      inspectZip(file.original.data);
      return exportRetainedWorkbook(file);
    }
    const ExcelJS = (await import('exceljs')).default,
      book = new ExcelJS.Workbook(),
      calc = calculator(file.content.sheets);
    book.creator = 'Noffice';
    book.calcProperties.fullCalcOnLoad = true;
    for (const sheet of file.content.sheets) {
      const ws = book.addWorksheet(sheet.name, { state: sheet.state || 'visible' });
      ws.columns = Array.from({ length: sheet.cols }, (_, c) => ({
        width: sheet.columnWidths?.[c] ? (sheet.columnWidths[c] - 5) / 7 : 18,
        hidden: sheet.hiddenColumns?.includes(c) || false,
      }));
      for (const [r, height] of Object.entries(sheet.rowHeights || {}))
        ws.getRow(Number(r) + 1).height = height * 0.75;
      for (const r of sheet.hiddenRows || []) ws.getRow(r + 1).hidden = true;
      ws.views = [
        {
          state: sheet.frozenRows || sheet.frozenColumns ? 'frozen' : 'normal',
          ySplit: sheet.frozenRows || 0,
          xSplit: sheet.frozenColumns || 0,
          showGridLines: sheet.gridLines !== false,
        },
      ];
      for (const [ref, model] of Object.entries(sheet.cells)) {
        const cell = ws.getCell(ref),
          calculated = calc(sheet, ref);
        cell.value = model.value.startsWith('=')
          ? {
              formula: model.value.slice(1),
              ...(sheet.arrayFormulas?.find((a) => a.anchor === ref)
                ? {
                    shareType: 'array' as const,
                    ref: sheet.arrayFormulas.find((a) => a.anchor === ref)!.ref,
                  }
                : {}),
              result:
                typeof calculated === 'string' && calculated.startsWith('#')
                  ? undefined
                  : calculated,
            }
          : calculated;
        // A synthesized direct font masks table header colors in Excel. Inherit
        // the workbook font for table cells until the user explicitly formats it.
        if (
          !tableAt(sheet, ref) ||
          ['fontFamily', 'fontSize', 'bold', 'italic', 'underline', 'color'].some(
            (key) => model[key as keyof Cell] !== undefined,
          )
        )
          cell.font = {
            name: model.fontFamily || defaultFont,
            size: model.fontSize || 11,
            underline: model.underline,
            bold: model.bold,
            italic: model.italic,
            color: model.color ? { argb: 'FF' + model.color.slice(1) } : undefined,
          };
        if (model.fill)
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF' + model.fill.slice(1) },
          };
        cell.alignment = {
          horizontal: model.align,
          vertical: model.vertical,
          wrapText: model.wrap,
        };
        if (model.dataType === 'text') cell.value = model.value;
        if (model.dataType === 'boolean') cell.value = model.value.toUpperCase() === 'TRUE';
        cell.numFmt =
          model.numFmt ||
          (model.format === 'currency'
            ? '$#,##0.00'
            : model.format === 'percent'
              ? '0.00%'
              : model.format === 'number'
                ? '0.00'
                : 'General');
      }
      for (const range of sheet.merges || []) ws.mergeCells(range);
    }
    let data: Uint8Array = new Uint8Array(await book.xlsx.writeBuffer());
    if (file.content.sheets.some((s) => s.tables?.length)) {
      const zip = await JSZip.loadAsync(data),
        mapping = await workbookSheets(zip);
      // ExcelJS tags its Office 2007 theme as a newer default, causing Excel to substitute colors.
      for (const props of elements(mapping.book, 'workbookPr'))
        props.removeAttribute('defaultThemeVersion');
      zip.file('xl/workbook.xml', new XMLSerializer().serializeToString(mapping.book));
      const styleDoc = parseXML(await zip.file('xl/styles.xml')!.async('string'));
      const defaultFontNode = elements(styleDoc, 'fonts')[0]?.firstElementChild;
      if (defaultFontNode) {
        elements(defaultFontNode, 'name')[0]?.setAttribute('val', defaultFont);
        elements(defaultFontNode, 'scheme').forEach((e) => e.remove());
        zip.file('xl/styles.xml', new XMLSerializer().serializeToString(styleDoc));
      }
      for (let i = 0; i < file.content.sheets.length; i++) {
        const sheet = file.content.sheets[i],
          path = mapping.sheets[i].path;
        const doc = parseXML(await zip.file(path)!.async('string'));
        if (await writeSheetTables(zip, path, doc.documentElement, { ...sheet, tables: [] }, sheet))
          zip.file(path, new XMLSerializer().serializeToString(doc));
        for (const filter of sheet.autoFilters || []) {
          if (!sheet.tables?.some((t) => t.sourcePath === filter.sourcePath)) continue;
          const tableDoc = parseXML(await zip.file(filter.sourcePath!)!.async('string'));
          writeFilterColumns(elements(tableDoc, 'autoFilter')[0], filter.columns);
          zip.file(filter.sourcePath!, new XMLSerializer().serializeToString(tableDoc));
        }
      }
      data = await zip.generateAsync({ type: 'uint8array' });
    }
    if (file.content.sheets.some((s) => s.autoFilters?.length)) {
      const zip = await JSZip.loadAsync(data),
        mapping = await workbookSheets(zip);
      for (let i = 0; i < file.content.sheets.length; i++) {
        const sheet = file.content.sheets[i];
        const worksheetFilters = (sheet.autoFilters || []).filter(
          (f) => !sheet.tables?.some((t) => t.sourcePath === f.sourcePath),
        );
        if (!worksheetFilters.length) continue;
        if (worksheetFilters.length !== 1 || worksheetFilters[0].sourcePath)
          throw new Error('This new workbook cannot export imported table filter metadata.');
        const path = mapping.sheets[i].path,
          doc = parseXML(await zip.file(path)!.async('string')),
          root = doc.documentElement;
        const filter = doc.createElementNS(root.namespaceURI, 'autoFilter');
        filter.setAttribute('ref', worksheetFilters[0].ref);
        writeFilterColumns(filter, worksheetFilters[0].columns);
        const after = [
          'mergeCells',
          'conditionalFormatting',
          'dataValidations',
          'hyperlinks',
          'printOptions',
          'pageMargins',
          'pageSetup',
          'headerFooter',
          'drawing',
          'tableParts',
          'extLst',
        ];
        root.insertBefore(
          filter,
          Array.from(root.children).find((e) => after.includes(e.localName)) || null,
        );
        if (sheet.filterMode) {
          const props =
            elements(root, 'sheetPr')[0] || doc.createElementNS(root.namespaceURI, 'sheetPr');
          props.setAttribute('filterMode', '1');
          root.insertBefore(props, root.firstChild);
        }
        zip.file(path, new XMLSerializer().serializeToString(doc));
      }
      data = await zip.generateAsync({ type: 'uint8array' });
    }
    if (
      file.content.sheets.some(
        (s) => s.validationRanges?.length || Object.values(s.cells).some((c) => c.validation),
      )
    ) {
      const zip = await JSZip.loadAsync(data),
        mapping = await workbookSheets(zip);
      for (let i = 0; i < file.content.sheets.length; i++) {
        const sheet = file.content.sheets[i],
          path = mapping.sheets[i].path;
        const doc = parseXML(await zip.file(path)!.async('string'));
        if (writeValidationChanges(doc.documentElement, {}, sheet.cells, sheet.validationRanges))
          zip.file(path, new XMLSerializer().serializeToString(doc));
      }
      data = await zip.generateAsync({ type: 'uint8array' });
    }
    if (file.content.sheets.some((s) => Object.values(s.cells).some((c) => c.hyperlink))) {
      const zip = await JSZip.loadAsync(data),
        mapping = await workbookSheets(zip);
      for (let i = 0; i < file.content.sheets.length; i++) {
        const path = mapping.sheets[i].path,
          doc = parseXML(await zip.file(path)!.async('string'));
        if (
          await writeWorksheetLinks(
            zip,
            path,
            doc.documentElement,
            {},
            file.content.sheets[i].cells,
          )
        )
          zip.file(path, new XMLSerializer().serializeToString(doc));
      }
      data = await zip.generateAsync({ type: 'uint8array' });
    }
    return new Blob([new Uint8Array(data)], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }
  if (file.content.kind === 'powerpoint') {
    const PptxGenJS = (await import('pptxgenjs')).default,
      pptx = new PptxGenJS();
    const aspectRatio = file.content.aspectRatio || 16 / 9;
    pptx.defineLayout({ name: 'NOFFICE', width: 960 / 72, height: 960 / 72 / aspectRatio });
    pptx.layout = 'NOFFICE';
    pptx.author = 'Noffice';
    pptx.title = file.name;
    for (const model of file.content.slides) {
      const slide = pptx.addSlide();
      slide.background = {
        color: model.background.replace('#', ''),
        transparency: 100 * (1 - (model.backgroundOpacity ?? 1)),
      };
      slide.addNotes(model.notes);
      for (const el of model.elements) {
        const verticalScale = 960 / aspectRatio / 540;
        const bounds = {
          x: el.x / 72,
          rotate: el.rotation || 0,
          y: (el.y * verticalScale) / 72,
          w: el.w / 72,
          h: (el.h * verticalScale) / 72,
        };
        const fill = {
          color: el.fill === 'transparent' ? 'FFFFFF' : el.fill.replace('#', ''),
          transparency: el.fill === 'transparent' ? 100 : 100 * (1 - (el.fillOpacity ?? 1)),
        };
        const line = el.outline
          ? {
              color:
                el.outline.color === 'transparent' ? 'FFFFFF' : el.outline.color.replace('#', ''),
              width: el.outline.width,
              transparency:
                el.outline.color === 'transparent' ? 100 : 100 * (1 - el.outline.opacity),
              // The generator writes DrawingML preset values verbatim; its declaration omits three legal values.
              dashType: el.outline.dash as
                | 'solid'
                | 'dash'
                | 'dashDot'
                | 'lgDash'
                | 'lgDashDot'
                | 'lgDashDotDot'
                | 'sysDash'
                | 'sysDot',
            }
          : { transparency: 100 };
        if (el.type === 'image' && el.src)
          slide.addImage({ ...bounds, data: el.src, flipH: el.flipH, flipV: el.flipV });
        else if (el.type === 'text')
          slide.addText(el.text, {
            ...bounds,
            fontSize: el.fontSize,
            fontFace: el.fontFamily || defaultFont,
            italic: el.italic,
            underline: el.underline ? { style: 'sng' } : undefined,
            bold: el.bold,
            color: el.color.replace('#', ''),
            align: el.align,
            valign: 'top',
            margin: 0,
            breakLine: false,
            fill,
            line,
          });
        else
          slide.addShape(el.type === 'ellipse' ? pptx.ShapeType.ellipse : pptx.ShapeType.rect, {
            ...bounds,
            fill,
            line,
          });
      }
    }
    return (await pptx.write({ outputType: 'blob' })) as Blob;
  }
  if (file.original?.name.toLowerCase().endsWith('.docx')) {
    const { exportRetainedDocument } = await import('./docx-preserve');
    return exportRetainedDocument(file, wordLayout);
  }
  const authoredContent = file.content;
  // Source-derived copies can retain resolved section/list metadata without an
  // original ZIP. Their fresh package still needs those explicit page settings.
  const authoredSections = resolveWordSections(file.content);
  const authoredSection = authoredSections[0];
  const authoredIds = new Set(authoredSections.map((s) => s.id));
  const authoredStories = file.content.stories && wordStoriesSchema.parse(file.content.stories);
  if (
    authoredStories &&
    (!authoredSection ||
      authoredStories.parts.some((p) => !p.created || p.copiedFrom) ||
      authoredStories.references?.some((r) => !authoredIds.has(r.sectionId)) ||
      authoredStories.sectionOptions?.some((o) => !authoredIds.has(o.sectionId)))
  )
    throw Error(
      'These headers or footers need their retained DOCX source. Export a Noffice backup to preserve them.',
    );
  const D = await import('docx');
  const dom = new DOMParser().parseFromString(sanitizeHTML(file.content.html), 'text/html');
  const storyDocuments = new Map(
    (authoredStories?.parts || []).map((p) => [
      p.path,
      new DOMParser().parseFromString(sanitizeHTML(p.html), 'text/html'),
    ]),
  );
  const imageRuns = new Map<Element, InstanceType<typeof D.ImageRun>>();
  const numbering = authoredWordNumbering([dom.body, ...[...storyDocuments.values()].map(d => d.body)], file.content.numbering);
  for (const image of [dom, ...storyDocuments.values()].flatMap((d) => [
    ...d.querySelectorAll('img'),
  ])) {
    const match = /^data:image\/(png|jpeg|gif);base64,(.+)$/i.exec(image.src);
    if (!match) continue;
    const data = Uint8Array.from(atob(match[2]), (char) => char.charCodeAt(0));
    const bitmap = await createImageBitmap(new Blob([data]));
    const scale = Math.min(1, 600 / bitmap.width);
    imageRuns.set(
      image,
      new D.ImageRun({
        type: match[1] === 'jpeg' ? 'jpg' : (match[1].toLowerCase() as 'png' | 'gif'),
        data,
        transformation: {
          width: Math.round(bitmap.width * scale),
          height: Math.round(bitmap.height * scale),
        },
        altText: { title: image.alt, description: image.alt, name: image.alt || 'Document image' },
      }),
    );
    bitmap.close();
  }
  type RunOptions = {
    style?: string;
    bold?: boolean;
    italics?: boolean;
    strike?: boolean;
    subScript?: boolean;
    superScript?: boolean;
    underline?: { type: 'single' };
    color?: string;
    size?: number;
    font?: string;
    kern?: number | '0pt';
    features?: number;
    highlight?: 'yellow';
  };
  type DocRun =
    | InstanceType<typeof D.TextRun>
    | InstanceType<typeof D.ImageRun>
    | InstanceType<typeof D.ExternalHyperlink>;
  const rootHash = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(file.id));
  const rootSession = new DataView(rootHash).getUint32(0) & 0x7fffffff || 1;
  const sessions = new Map<string, string>([
    ['', rootSession.toString(16).padStart(8, '0').toUpperCase()],
  ]);
  class AuthoredTextRun extends D.TextRun {
    constructor(options: IRunOptions & { features?: number }, session = '') {
      const { features, ...run } = options;
      super(run);
      if (features !== undefined) addFontFeatures(this.properties, features);
      let id = sessions.get(session);
      if (!id) {
        if (sessions.size >= 1000000)
          throw Error('This document has too many editing sessions to export.');
        id = (rootSession + sessions.size).toString(16).padStart(8, '0').toUpperCase();
        sessions.set(session, id);
      }
      this.root.push(new D.Attributes({ rsidR: id }));
    }
  }
  function addFontFeatures(properties: InstanceType<typeof D.RunProperties>, value: number) {
    if (wordFontFeaturesValue(value) === null) throw Error('Invalid Word font-feature setting.');
    for (const [name, setting] of [
      ['ligatures', wordLigatureNames[value & 15]],
      ['cntxtAlts', value & 16 ? '1' : '0'],
    ])
      properties.push(
        new D.ImportedXmlComponent(`w14:${name}`, {
          'xmlns:w14': WORD_2010_NS,
          'w14:val': setting,
        }),
      );
  }
  class AuthoredParagraph extends D.Paragraph {
    constructor(
      options: IParagraphOptions,
      features: number | null,
      script: WordScript | null,
      element: HTMLElement,
    ) {
      const { run, ...paragraph } = options;
      const before = wordParagraphSpace(element.getAttribute('data-word-space-before'));
      const after = wordParagraphSpace(element.getAttribute('data-word-space-after'));
      super({
        ...paragraph,
        contextualSpacing: element.hasAttribute('data-word-contextual-spacing')
          ? element.getAttribute('data-word-contextual-spacing') === 'true'
          : undefined,
      });
      const properties = this.root.find((item) => item instanceof D.ParagraphProperties);
      if (!properties) throw Error('The Word paragraph properties are unavailable.');
      if (before || after) {
        const values: Record<string, string> = {};
        if (before) Object.assign(values, wordParagraphSpaceXml('before', before));
        if (after) Object.assign(values, wordParagraphSpaceXml('after', after));
        const serialize = (
          properties as InstanceType<typeof D.ParagraphProperties>
        ).prepForXml.bind(properties);
        // Extend the public serializer's existing spacing element. Appending a
        // second/custom element would put spacing after contextualSpacing and
        // violate the ordered paragraph-properties schema.
        properties.prepForXml = (context: IContext) => {
          const output = serialize(context),
            children = output?.['w:pPr'];
          const spacing =
            Array.isArray(children) && children.find((item) => item?.['w:spacing'])?.['w:spacing'];
          if (!spacing?._attr) throw Error('The Word spacing properties are unavailable.');
          Object.assign(
            spacing._attr,
            Object.fromEntries(Object.entries(values).map(([key, value]) => ['w:' + key, value])),
          );
          return output;
        };
      }
      const mark = new D.RunProperties(run);
      if (features !== null) addFontFeatures(mark, features);
      if (script !== null)
        mark.push(new D.ImportedXmlComponent('w:vertAlign', { 'w:val': script }));
      properties.addChildElement(mark);
    }
  }
  function runs(node: Node, style: RunOptions = { font: defaultFont }, session?: string): DocRun[] {
    if (node.nodeType === 3)
      return [new AuthoredTextRun({ text: node.textContent || '', ...style }, session)];
    if (!(node instanceof HTMLElement)) return [];
    if (node.tagName === 'IMG') return imageRuns.has(node) ? [imageRuns.get(node)!] : [];
    const next = { ...style };
    session =
      wordRunSession([
        { type: 'wordEditRun', attrs: { session: node.getAttribute('data-word-edit-run') } },
      ]) ?? session;
    const kerning = wordKerningValue(node.getAttribute('data-word-kerning'));
    if (kerning !== null) next.kern = kerning === 0 ? '0pt' : kerning;
    const features = wordFontFeaturesValue(
      node.getAttribute('data-word-font-features') ??
        node.getAttribute('data-word-paragraph-font-features'),
    );
    if (features !== null) next.features = features;
    if (['STRONG', 'B'].includes(node.tagName)) next.bold = true;
    if (['EM', 'I'].includes(node.tagName)) next.italics = true;
    if (node.tagName === 'U') next.underline = { type: 'single' };
    if (['S', 'DEL'].includes(node.tagName)) next.strike = true;
    if (node.tagName === 'SUB') next.subScript = true;
    if (node.tagName === 'SUP') next.superScript = true;
    if (node.tagName === 'MARK') next.highlight = 'yellow';
    if (node.style.fontSize)
      next.size = parseFloat(node.style.fontSize) * (node.style.fontSize.endsWith('px') ? 1.5 : 2);
    if (node.style.fontFamily) next.font = node.style.fontFamily.replace(/^['"]|['"]$/g, '');
    if (node.style.color) {
      const m = node.style.color.match(/\d+/g);
      next.color = m
        ? m
            .slice(0, 3)
            .map((n) => Number(n).toString(16).padStart(2, '0'))
            .join('')
        : node.style.color.replace('#', '');
    }
    if (node.tagName === 'A' && /^(https?:|mailto:)/i.test(node.getAttribute('href') || ''))
      return [
        new D.ExternalHyperlink({
          link: node.getAttribute('href')!,
          children: [...node.childNodes].flatMap((child) =>
            runs(child, { ...next, style: 'Hyperlink' }, session),
          ),
        }),
      ];
    if (node.tagName === 'BR') return [new D.TextRun({ ...next, break: 1 })];
    if (node.matches('span[data-word-tab]'))
      return [new D.TextRun({ ...next, children: [new D.Tab()] })];
    if (node.matches('span[data-word-hyphen="optional"]') && node.textContent === '\u00ad')
      return [new D.TextRun({ ...next, children: [new D.SoftHyphen()] })];
    if (node.matches('span[data-word-hyphen="literal"]') && node.textContent === '-')
      return [new D.TextRun({ ...next, text: '\u00ad' })];
    if (node.matches('span[data-word-page-break]')) return [new D.PageBreak()];
    if (node.matches('span[data-word-column-break]')) return [new D.ColumnBreak()];
    return [...node.childNodes].flatMap((child) => runs(child, next, session));
  }
  function blocks(
    root: HTMLElement,
    list?: 'bullet' | 'number',
    level = 0,
    story = false,
  ): (InstanceType<typeof D.Paragraph> | InstanceType<typeof D.Table>)[] {
    return [...root.children].flatMap(
      (node): (InstanceType<typeof D.Paragraph> | InstanceType<typeof D.Table>)[] => {
        const el = node as HTMLElement;
        if (el.tagName === 'IMG') return [new D.Paragraph({ children: runs(el) })];
        if (el.tagName === 'UL' || el.tagName === 'OL')
          return blocks(el, el.tagName === 'UL' ? 'bullet' : 'number', level, story);
        if (el.tagName === 'LI') return blocks(el, list, level + 1, story);
        if (el.tagName === 'BLOCKQUOTE') return blocks(el, undefined, 0, story);
        if (el.tagName === 'TABLE')
          return [
            new D.Table({
              rows: [...el.querySelectorAll('tr')].map(
                (row) =>
                  new D.TableRow({
                    children: [...row.children].map(
                      (cell) =>
                        new D.TableCell({
                          columnSpan: Math.max(1, Number(cell.getAttribute('colspan')) || 1),
                          rowSpan: Math.max(1, Number(cell.getAttribute('rowspan')) || 1),
                          children: blocks(cell as HTMLElement, undefined, 0, story).length
                            ? blocks(cell as HTMLElement, undefined, 0, story)
                            : [new D.Paragraph({ children: runs(cell) })],
                        }),
                    ),
                  }),
              ),
              width: { size: 100, type: D.WidthType.PERCENTAGE },
            }),
          ];
        const heading = /^H([1-6])$/.exec(el.tagName);
        const lineSpacing = wordLineSpacing(
          el.style.lineHeight,
          el.getAttribute('data-word-line-rule'),
        );
        return [
          new AuthoredParagraph(
            {
              children: runs(el),
              run: {
                size: el.style.fontSize
                  ? parseFloat(el.style.fontSize) * (el.style.fontSize.endsWith('px') ? 1.5 : 2)
                  : undefined,
                font: el.style.fontFamily.replace(/^(["'])(.*)\1$/, '$2') || undefined,
              },
              bidirectional: el.dir === 'rtl' ? true : undefined,
              heading: heading ? (`Heading${heading[1]}` as 'Heading1') : undefined,
              keepNext: el.style.breakAfter ? el.style.breakAfter === 'avoid' : undefined,
              keepLines: el.style.breakInside ? el.style.breakInside === 'avoid' : undefined,
              pageBreakBefore: el.style.breakBefore ? el.style.breakBefore === 'page' : undefined,
              widowControl: el.style.orphans ? Number(el.style.orphans) > 1 : undefined,
              indent: ![el.style.marginInlineStart, el.style.marginInlineEnd, el.style.textIndent].some(Boolean) ? undefined : {
                start: el.style.marginInlineStart
                  ? Math.round(
                      parseFloat(el.style.marginInlineStart) *
                        (el.style.marginInlineStart.endsWith('pt') ? 20 : 15),
                    )
                  : undefined,
                end: el.style.marginInlineEnd
                  ? Math.round(
                      parseFloat(el.style.marginInlineEnd) *
                        (el.style.marginInlineEnd.endsWith('pt') ? 20 : 15),
                    )
                  : undefined,
                firstLine:
                  parseFloat(el.style.textIndent) >= 0
                    ? Math.round(
                        parseFloat(el.style.textIndent) *
                          (el.style.textIndent.endsWith('pt') ? 20 : 15),
                      )
                    : undefined,
                hanging:
                  parseFloat(el.style.textIndent) < 0
                    ? Math.round(
                        -parseFloat(el.style.textIndent) *
                          (el.style.textIndent.endsWith('pt') ? 20 : 15),
                      )
                    : undefined,
              },
              alignment: ['left', 'center', 'right', 'justify'].includes(el.style.textAlign)
                ? ((el.style.textAlign === 'justify' ? 'both' : el.style.textAlign) as 'left')
                : undefined,
              bullet: list === 'bullet' && !numbering.reference(el) ? { level: Math.max(0, level - 1) } : undefined,
              numbering:
                list && numbering.reference(el)
                  ? { reference: numbering.reference(el)!, level: Math.max(0, level - 1) }
                  : undefined,
              spacing: {
                before: el.style.marginTop
                  ? Math.round(
                      parseFloat(el.style.marginTop) *
                        (el.style.marginTop.endsWith('pt') ? 20 : 15),
                    )
                  : undefined,
                after: el.style.marginBottom
                  ? Math.round(
                      parseFloat(el.style.marginBottom) *
                        (el.style.marginBottom.endsWith('pt') ? 20 : 15),
                    )
                  : story
                    ? 0
                    : wordDefaults.spaceAfter * 20,
                line:
                  lineSpacing?.line ?? (story ? 240 : Math.round(wordDefaults.lineMultiple * 240)),
                lineRule:
                  lineSpacing?.rule === 'atLeast'
                    ? D.LineRuleType.AT_LEAST
                    : lineSpacing?.rule === 'exact'
                      ? D.LineRuleType.EXACT
                      : D.LineRuleType.AUTO,
              },
            },
            wordFontFeaturesValue(el.getAttribute('data-word-paragraph-font-features')),
            wordScriptValue(el.getAttribute('data-word-paragraph-script')),
            el,
          ),
        ];
      },
    );
  }
  const margin =
    file.content.margin === 'narrow' ? 720 : file.content.margin === 'wide' ? 2160 : 1440;
  const { wordDefaultFontBytes } = await import('./word-default-font');
  const sectionBodies = [dom.body];
  if (file.content.sectionState) {
    const { getSchema } = await import('@tiptap/core');
    const { DOMParser: ParagraphParser } = await import('@tiptap/pm/model');
    const { wordExtensions } = await import('./word-extensions');
    const model = ParagraphParser.fromSchema(getSchema(wordExtensions())).parse(dom.body);
    validateSectionState(file.content.sectionState, file.content.docxStructure, model);
    const paragraphs = [...dom.body.querySelectorAll('p,h1,h2,h3,h4,h5,h6')];
    const copyRange = (range: Range) => {
      let contents: Node = range.cloneContents();
      for (let ancestor: Node | null = range.commonAncestorContainer; ancestor && ancestor !== dom.body;
        ancestor = ancestor.parentNode) {
        const wrapper = ancestor.cloneNode(false);
        wrapper.appendChild(contents);
        contents = wrapper;
      }
      return contents;
    };
    let previous: Element | null = null;
    sectionBodies.length = 0;
    for (const boundary of file.content.sectionState.breaks) {
      const target = paragraphs[boundary.paragraph];
      if (!target) throw Error('The authored section break has no supported output paragraph.');
      const group = dom.createElement('div');
      const range = dom.createRange();
      if (previous) range.setStartAfter(previous); else range.setStart(dom.body, 0);
      range.setEndAfter(target);
      // cloneContents retains the enclosing list/item on either side of an
      // interior boundary. The writer emits one shared numbering definition.
      group.append(copyRange(range));
      sectionBodies.push(group);
      previous = target;
    }
    const group = dom.createElement('div');
    const range = dom.createRange();
    if (previous) range.setStartAfter(previous); else range.setStart(dom.body, 0);
    range.setEnd(dom.body, dom.body.childNodes.length);
    group.append(copyRange(range));
    sectionBodies.push(group);
  }
  const sections = (authoredSections.length ? authoredSections : [undefined]).map(
    (section, index) => {
      const headers: Partial<Record<'default' | 'first' | 'even', InstanceType<typeof D.Header>>> =
        {};
      const footers: Partial<Record<'default' | 'first' | 'even', InstanceType<typeof D.Footer>>> =
        {};
      for (const [kind, key] of [
        ['header', 'headers'],
        ['footer', 'footers'],
      ] as const) {
        for (const slot of ['default', 'first', 'even'] as const) {
          const ref = section?.[key][slot];
          // A body-only copy deliberately has no story model. Historical
          // source references supply geometry, not missing header contents.
          if (!ref || ref.inherited || !authoredStories) continue;
          const part = authoredStories?.parts.find(
            (p) => p.kind === kind && p.relationshipIds.includes(ref.relationshipId),
          );
          if (!part) throw Error('The authored section has an unavailable header or footer.');
          const children = blocks(storyDocuments.get(part.path)!.body, undefined, 0, true);
          if (!children.length) children.push(new D.Paragraph(''));
          if (kind === 'header') headers[slot] = new D.Header({ children });
          else footers[slot] = new D.Footer({ children });
        }
      }
      return {
        headers,
        footers,
        properties: {
          titlePage: section?.differentFirstPage,
          type:
            authoredContent.sectionState && section
              ? (
                  {
                    nextPage: D.SectionType.NEXT_PAGE,
                    continuous: D.SectionType.CONTINUOUS,
                    evenPage: D.SectionType.EVEN_PAGE,
                    oddPage: D.SectionType.ODD_PAGE,
                  } as Record<string, (typeof D.SectionType)[keyof typeof D.SectionType]>
                )[section.start]
              : undefined,
          page: {
            size: {
              orientation:
                (section?.orientation || authoredContent.orientation) === 'landscape'
                  ? D.PageOrientation.LANDSCAPE
                  : D.PageOrientation.PORTRAIT,
              ...(section?.width && section.height
                ? {
                    width: Math.min(section.width, section.height),
                    height: Math.max(section.width, section.height),
                  }
                : authoredContent.paper === 'a4'
                  ? { width: 11906, height: 16838 }
                  : { width: 12240, height: 15840 }),
            },
            margin: {
              top: section?.margins.top ?? margin,
              bottom: section?.margins.bottom ?? margin,
              left: section?.margins.left ?? margin,
              right: section?.margins.right ?? margin,
              header: section?.margins.header ?? undefined,
              footer: section?.margins.footer ?? undefined,
            },
          },
        },
        children: blocks(sectionBodies[index]),
      };
    },
  );
  const document = new D.Document({
    creator: 'Noffice',
    title: file.name,
    evenAndOddHeaderAndFooters: authoredStories?.evenAndOddHeaders,
    // docx types this as Node Buffer; its browser packer consumes Uint8Array.
    fonts: [{ name: wordDefaults.fontFamily, data: (await wordDefaultFontBytes()) as Buffer }],
    styles: {
      default: {
        document: {
          run: { font: wordDefaults.fontFamily, size: wordDefaults.fontSize * 2 },
          paragraph: {
            spacing: {
              before: 0,
              after: wordDefaults.spaceAfter * 20,
              line: Math.round(wordDefaults.lineMultiple * 240),
              lineRule: D.LineRuleType.AUTO,
            },
          },
        },
      },
    },
    numbering: { config: numbering.config },
    sections,
  });
  // New packages own this register. Retained packages continue to use their
  // existing register and collision checks in the preservation adapter.
  const register = new D.ImportedXmlComponent('w:rsids');
  register.addChildElement(new D.ImportedXmlComponent('w:rsidRoot', { 'w:val': sessions.get('') }));
  for (const id of sessions.values())
    register.addChildElement(new D.ImportedXmlComponent('w:rsid', { 'w:val': id }));
  document.Settings.addChildElement(register);
  {
    const zip = await JSZip.loadAsync(await D.Packer.toArrayBuffer(document));
    const { retainAuthoredFontEmbedding } = await import('./docx-authored-fonts');
    await retainAuthoredFontEmbedding(zip);
    if (authoredContent.sectionState?.breaks.length) {
      const { attachAuthoredSectionProperties } = await import('./docx-authored-sections');
      await attachAuthoredSectionProperties(
        zip,
        authoredContent.sectionState.breaks.map((b) => b.paragraph),
      );
    }
    if (authoredStories?.parts.length) {
      const { styleAuthoredDocxStories } = await import('./docx-authored-stories');
      await styleAuthoredDocxStories(zip);
    }
    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
  }
}
