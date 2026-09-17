import type { Cell, Sheet, WorkbookContent } from './model';
import { address, coordinates } from './formulas';

export function sheetLinkAddress(value: string) {
  const href = value.trim();
  if (!href || href.length > 8192 || /[\u0000-\u001f\u007f]/.test(href))
    throw Error('Enter a web, email or workbook address.');
  if (href.startsWith('#') && href.length > 1) return href;
  if (/^(https?:\/\/|mailto:)/i.test(href) && !/\s/.test(href)) {
    try {
      const url = new URL(href);
      if (url.protocol === 'mailto:' ? !!url.pathname : !!url.hostname) return href;
    } catch {
      /* Report one useful validation message below. */
    }
  }
  throw Error('Use https://, http://, mailto: or #Sheet!A1 for a workbook location.');
}
export function internalSheetLink(href: string, sheet: Sheet, sheets: Sheet[]) {
  let source = sheetLinkAddress(href).slice(1);
  if (!href.startsWith('#')) throw Error('This link is not a workbook location.');
  const seen = new Set<string>();
  for (let depth = 0; depth < 10; depth++) {
    const key = source.toLowerCase();
    if (!sheet.definedNames?.[key]) break;
    if (seen.has(key)) throw Error('This defined name contains a circular reference.');
    seen.add(key);
    source = sheet.definedNames[key].replace(/^=/, '');
  }
  const match =
    /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?(\$?[A-Z]+\$?[1-9]\d*)(?::(\$?[A-Z]+\$?[1-9]\d*))?$/i.exec(
      source,
    );
  if (!match) throw Error('Use a cell, range or defined name that refers to a cell range.');
  const name = match[1]?.replaceAll("''", "'") || match[2];
  const target = name ? sheets.find((s) => s.name.toLowerCase() === name.toLowerCase()) : sheet;
  if (!target) throw Error('The linked worksheet does not exist.');
  if (target.state && target.state !== 'visible') throw Error('The linked worksheet is hidden.');
  const [r, c] = coordinates(match[3]),
    [rr, cc] = coordinates(match[4] || match[3]);
  if (Math.max(r, rr) >= target.rows || Math.max(c, cc) >= target.cols)
    throw Error('The linked range is outside this editor’s loaded grid.');
  if (target.hiddenRows?.includes(r) || target.hiddenColumns?.includes(c))
    throw Error('The linked cell is hidden.');
  return { sheetId: target.id, ref: address(r, c), end: address(rr, cc) };
}
export function editSheetLink(
  sheet: Sheet,
  ref: string,
  href: string | null,
  text: string,
  tooltip = '',
): Sheet {
  if (sheet.protected) throw Error('Hyperlinks cannot be changed on a protected worksheet.');
  const before: Cell = sheet.cells[ref] || { value: '' };
  if (text.length > 32767 || tooltip.length > 255)
    throw Error('Link text or ScreenTip exceeds the supported length.');
  const next = { ...before };
  if (href === null) {
    delete next.hyperlink;
    delete next.hyperlinkTooltip;
  } else {
    next.hyperlink = sheetLinkAddress(href);
    next.hyperlinkTooltip = tooltip || undefined;
    if (!(before.value.startsWith('=') && before.dataType !== 'text') && text !== before.value) {
      next.value = text || next.hyperlink;
      next.dataType = 'text';
      delete next.cachedValue;
    } else if (!before.value) {
      next.value = text || next.hyperlink;
      next.dataType = 'text';
    }
    if (!before.hyperlink) {
      next.color = '#0563c1';
      next.underline = true;
    }
  }
  return JSON.stringify(before) === JSON.stringify(next)
    ? sheet
    : { ...sheet, cells: { ...sheet.cells, [ref]: next } };
}
export function removeSheetLinks(content: WorkbookContent, sheet: Sheet, refs: string[]) {
  let next = sheet;
  for (const ref of refs) if (next.cells[ref]?.hyperlink) next = editSheetLink(next, ref, null, '');
  return next === sheet
    ? content
    : { ...content, sheets: content.sheets.map((s) => (s.id === sheet.id ? next : s)) };
}
