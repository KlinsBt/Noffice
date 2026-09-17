<script lang="ts">
  import { arrayContains, arrayEditError, enterArrayFormula, singleArray } from '../sheet-arrays';
  import { tableRowEntry, syncTableTotals } from '../sheet-table-entry';
  import { calculatedColumnEntry, clearRemovedColumnFormulas } from '../sheet-calculated-columns';
  import SearchBar from './SearchBar.svelte';
  import { tick } from 'svelte';
  import {
    workbookMatches,
    replaceWorkbookMatches,
    nextWorkbookMatch,
    type CellMatch,
  } from '../office-search';
  let replacement = $state(''),
    searchCase = $state(false),
    searchWhole = $state(false),
    searchStatus = $state('');
  let searchScope = $state('sheet'),
    searchLookIn = $state<'formulas' | 'values' | 'notes'>('formulas'),
    searchColumns = $state(false);
  let searchWildcards = $state(true);
  let searchSelection: string[] = [],
    searchSheet = '';
  function openSearch() {
    searchSelection = [...refs];
    searchSheet = sheet.id;
    setShowFind(true);
  }
  function searchMatches(forReplacement = false) {
    return workbookMatches(content, search, {
      matchCase: searchCase,
      wholeCell: searchWhole,
      sheetId:
        searchScope === 'workbook'
          ? undefined
          : searchScope === 'selection'
            ? searchSheet
            : sheet.id,
      refs: searchScope === 'selection' ? searchSelection : undefined,
      lookIn: searchLookIn,
      byColumns: searchColumns,
      wildcards: searchWildcards,
      forReplacement,
    });
  }
  async function showCellMatch(match: CellMatch) {
    setSheetId(match.sheetId);
    setSelected(match.ref);
    setAnchor(match.ref);
    setEditing(false);
    await tick();
    if (grid.current) {
      grid.current.scrollTop = layout.offsets[layout.rows.indexOf(coordinates(match.ref)[0])] || 0;
      const cell = grid.current.querySelector<HTMLElement>(`td[aria-label="${match.ref}"]`);
      cell?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
  }
  async function findCell(previous: boolean) {
    try {
      const matches = searchMatches();
      const next = nextWorkbookMatch(content, matches, sheet.id, selected, previous, searchColumns);
      if (next) {
        await showCellMatch(next);
        searchStatus = `${matches.indexOf(next) + 1} of ${matches.length} matching cells`;
      } else searchStatus = 'No matching cells.';
    } catch (error) {
      searchStatus = (error as Error).message;
    }
  }
  function replaceCells(all: boolean) {
    try {
      if (searchLookIn !== 'formulas') {
        searchStatus =
          'Choose Formulas to replace cell contents. Displayed results and notes are find-only.';
        return;
      }
      const found = searchMatches(true);
      const chosen = all
        ? found
        : found.filter((m) => m.sheetId === sheet.id && m.ref === selected);
      if (!all && !chosen.length) {
        void findCell(false);
        return;
      }
      const next = replaceWorkbookMatches(content, chosen, replacement);
      if (chosen.length) commit(next);
      searchStatus = `Replaced ${chosen.length} cells.`;
    } catch (error) {
      searchStatus = (error as Error).message;
    }
  }
  import RibbonGroup from './RibbonGroup.svelte';
  import { sheetLayout } from '../sheet-layout';
  import { freezeSheet, type FreezeCommand } from '../sheet-panes';
  import { validationChoices, validationError, validationSetupError } from '../cell-validation';
  import {
    applyValidation,
    cellValidation,
    checkedValidation,
    type ValidationRule,
  } from '../sheet-validation';
  import ValidationDialog from './ValidationDialog.svelte';
  let validationOpen = $state(false);
  function changeValidation(rule: ValidationRule) {
    const error = validationSetupError(checkedValidation(rule), sheet, content.sheets);
    if (error) throw Error(error);
    const next = applyValidation(sheet, refs, rule);
    if (next !== sheet)
      commit({ ...content, sheets: content.sheets.map((s) => (s.id === sheet.id ? next : s)) });
    validationOpen = false;
  }
  import { onMount } from 'svelte';
  import FontPicker from './FontPicker.svelte';
  import { fontStack, defaultFont } from '../fonts';
  import Tool from './Tool.svelte';
  import Modal from './Modal.svelte';
  import NumberFormatDialog from './NumberFormatDialog.svelte';
  import SheetFilterDialog from './SheetFilterDialog.svelte';
  import SheetLinkDialog from './SheetLinkDialog.svelte';
  import SheetTableDialog from './SheetTableDialog.svelte';
  import {
    createSheetTable,
    editSheetTable,
    tableAt,
    nextTableName,
    tableHeaderError,
    tableCellStyle,
    tableBounds,
    type TableStyle,
  } from '../sheet-tables';
  let tableOpen = $state(false);
  let tableSession = 0;
  function closeTable() {
    tableSession++;
    tableOpen = false;
  }
  async function applyTable(
    name: string,
    ref: string,
    style: TableStyle,
    column?: { index: number; name: string },
  ) {
    const before = content,
      session = tableSession;
    let next = activeTable
      ? editSheetTable(content, sheet.id, activeTable.name, ref, style)
      : createSheetTable(content, sheet.id, ref, name, style);
    if (
      activeTable &&
      (name !== activeTable.name || (column && column.name !== activeTable.columns[column.index]))
    ) {
      next = await onTableRename(next, {
        sheetId: sheet.id,
        table: activeTable.name,
        name,
        column,
      });
      if (!structureAlive || session !== tableSession) return;
      if (content !== before)
        throw Error('The workbook changed while this operation was running. Please retry.');
    }
    commit(next);
    tableOpen = false;
  }
  import {
    editSheetLink,
    internalSheetLink,
    sheetLinkAddress,
    removeSheetLinks,
  } from '../sheet-links';
  let linkOpen = $state(false);
  function applyLink(href: string | null, text: string, tooltip: string) {
    const next = editSheetLink(sheet, selected, href, text, tooltip);
    if (next.cells[selected].value !== current.value) {
      const error = validationError(current, next.cells[selected].value, sheet, content.sheets);
      if (error) throw Error(error);
    }
    if (next !== sheet)
      commit({ ...content, sheets: content.sheets.map((s) => (s.id === sheet.id ? next : s)) });
    linkOpen = false;
  }
  function clearLinks() {
    try {
      const next = removeSheetLinks(content, sheet, refs);
      if (next !== content) commit(next);
    } catch (e) {
      notify((e as Error).message);
    }
  }
  async function followLink() {
    try {
      const href = sheetLinkAddress(current.hyperlink || '');
      if (href.startsWith('#')) {
        const target = internalSheetLink(href, sheet, content.sheets);
        await showCellMatch({ sheetId: target.sheetId, ref: target.ref, matches: [] });
        setSelected(target.end);
      } else window.open(href, '_blank', 'noopener,noreferrer');
    } catch (e) {
      notify((e as Error).message);
    }
  }
  import { applySheetFilters, type SheetFilter } from '../sheet-filters';
  import { css, focusOnMount } from './css';
  import {
    Bold,
    Italic,
    Underline,
    AlignLeft,
    AlignCenter,
    AlignRight,
    Undo2,
    Redo2,
    Plus,
    ArrowDownAZ,
    ArrowUpAZ,
    ArrowDownToLine,
    Trash2,
    Search,
    ChartColumn,
    Sigma,
    Copy,
    PaintBucket,
    WrapText,
  } from '@lucide/svelte';
  import {
    address,
    calculator,
    colName,
    coordinates,
    displayValue,
    rangeAddresses,
    translateFormula,
  } from '../formulas';
  import { uid, type WorkbookContent, type Cell } from '../model';
  import type { StructureEdit } from '../sheet-structure';
  import type { TableRename } from '../table-references';
  import { applyHeaderEntry } from '../sheet-header-entry';
  let {
    content,
    onChange,
    onStructure,
    onTableRename,
    onBusy,
    notify,
  }: {
    content: WorkbookContent;
    onChange: (value: WorkbookContent) => void;
    onStructure: (value: WorkbookContent, edit: StructureEdit) => Promise<WorkbookContent>;
    onTableRename: (value: WorkbookContent, edit: TableRename) => Promise<WorkbookContent>;
    onBusy: (value: boolean) => void;
    notify: (text: string) => void;
  } = $props();
  let structureBusy = $state(false);
  let headerBusy = $state(false);
  $effect(() => onBusy(headerBusy));
  let structureAlive = true;
  onMount(() => () => {
    structureAlive = false;
    onBusy(false);
  });
  async function changeStructure(axis: StructureEdit['axis'], action: StructureEdit['action']) {
    if (structureBusy) return;
    if (editing) saveDraft();
    await tick();
    const before = content;
    const index = axis === 'row' ? 0 : 1;
    const a = coordinates(anchor)[index],
      b = coordinates(selected)[index];
    structureBusy = true;
    try {
      const next = await onStructure(before, {
        sheet: sheet.name,
        axis,
        action,
        at: Math.min(a, b),
        count: Math.abs(a - b) + 1,
      });
      if (!structureAlive) return;
      if (content !== before)
        throw Error('The workbook changed while this operation was running. Please retry.');
      commit(next);
      setEditing(false);
      setAnchor(selected);
    } catch (error) {
      if (structureAlive) notify((error as Error).message);
    } finally {
      structureBusy = false;
    }
  }
  let sheetId = $state.raw('');
  function setSheetId(value: typeof sheetId | ((previous: typeof sheetId) => typeof sheetId)) {
    sheetId = typeof value === 'function' ? value(sheetId) : value;
  }
  let selected = $state.raw('A1');
  function setSelected(value: typeof selected | ((previous: typeof selected) => typeof selected)) {
    selected = typeof value === 'function' ? value(selected) : value;
  }
  let anchor = $state.raw('A1');
  function setAnchor(value: typeof anchor | ((previous: typeof anchor) => typeof anchor)) {
    anchor = typeof value === 'function' ? value(anchor) : value;
  }
  let editing = $state.raw(false);
  function setEditing(value: typeof editing | ((previous: typeof editing) => typeof editing)) {
    editing = typeof value === 'function' ? value(editing) : value;
  }
  let draft = $state.raw('');
  function setDraft(value: typeof draft | ((previous: typeof draft) => typeof draft)) {
    draft = typeof value === 'function' ? value(draft) : value;
  }
  let scroll = $state.raw(0);
  function setScroll(value: typeof scroll | ((previous: typeof scroll) => typeof scroll)) {
    scroll = typeof value === 'function' ? value(scroll) : value;
  }
  let chart = $state.raw(false);
  let numberFormatOpen = $state(false);
  let editingFilter = $state<SheetFilter | null>(null);
  function openFilter() {
    if (sheet.protected) {
      notify('This worksheet is protected.');
      return;
    }
    const [r, c] = coordinates(selected);
    const existing = sheet.autoFilters?.find((f) => {
      const [a, b = a] = f.ref.split(':');
      const [fr, fc] = coordinates(a),
        [rr, cc] = coordinates(b);
      return r >= fr && r <= rr && c >= fc && c <= cc;
    });
    if (existing) {
      editingFilter = existing;
      return;
    }
    if (sheet.autoFilters?.some((f) => f.sourcePath === sheet.sourcePath)) {
      notify('Select a cell inside the existing filter range.');
      return;
    }
    const [ar, ac] = coordinates(anchor);
    if (ar === r) {
      notify('Select a header row and the data rows to filter.');
      return;
    }
    const start = address(Math.min(r, ar), Math.min(c, ac)),
      end = address(Math.max(r, ar), Math.max(c, ac));
    if (
      sheet.tables?.some((t) => {
        const [a, b = a] = t.ref.split(':');
        const [tr, tc] = coordinates(a),
          [rr, cc] = coordinates(b);
        return (
          Math.min(r, ar) <= rr &&
          Math.max(r, ar) >= tr &&
          Math.min(c, ac) <= cc &&
          Math.max(c, ac) >= tc
        );
      })
    ) {
      notify('Select a cell inside the table’s existing filter range.');
      return;
    }
    editingFilter = { ref: `${start}:${end}`, sourcePath: sheet.sourcePath, columns: [] };
  }
  function changeFilters(filters: SheetFilter[]) {
    try {
      const next = applySheetFilters(sheet, content.sheets, filters);
      commit({ ...content, sheets: content.sheets.map((s) => (s.id === sheet.id ? next : s)) });
      editingFilter = null;
      const [r, c] = coordinates(selected);
      if (next.hiddenRows?.includes(r)) {
        const ref = address(coordinates(filters[0]?.ref.split(':')[0] || 'A1')[0], c);
        setSelected(ref);
        setAnchor(ref);
      }
    } catch (error) {
      notify(error instanceof Error ? error.message : 'Could not apply this filter.');
    }
  }
  function setChart(value: typeof chart | ((previous: typeof chart) => typeof chart)) {
    chart = typeof value === 'function' ? value(chart) : value;
  }
  let rename = $state.raw(false);
  function setRename(value: typeof rename | ((previous: typeof rename) => typeof rename)) {
    rename = typeof value === 'function' ? value(rename) : value;
  }
  let sheetName = $state.raw('');
  function setSheetName(
    value: typeof sheetName | ((previous: typeof sheetName) => typeof sheetName),
  ) {
    sheetName = typeof value === 'function' ? value(sheetName) : value;
  }
  let search = $state.raw('');
  function setSearch(value: typeof search | ((previous: typeof search) => typeof search)) {
    search = typeof value === 'function' ? value(search) : value;
  }
  let showFind = $state.raw(false);
  function setShowFind(value: typeof showFind | ((previous: typeof showFind) => typeof showFind)) {
    showFind = typeof value === 'function' ? value(showFind) : value;
  }
  const undo: { current: WorkbookContent[] } = { current: [] };
  const redo: { current: WorkbookContent[] } = { current: [] };
  const grid: { current: HTMLDivElement | null } = { current: null };
  let gridHeight = $state(900);
  onMount(() => {
    if (!grid.current) return;
    const observer = new ResizeObserver(([entry]) => (gridHeight = entry.contentRect.height));
    observer.observe(grid.current);
    return () => observer.disconnect();
  });
  let sheet = $derived(
    content.sheets.find((s) => s.id === sheetId) ||
      content.sheets.find((s) => !s.state || s.state === 'visible') ||
      content.sheets[0],
  );
  const calculation = calculator([]);
  let calc = $derived.by(() => calculation.update(content.sheets));
  let refs = $derived.by(() => rangeAddresses(anchor, selected));
  let selectedSet = $derived.by(() => new Set(refs));
  let current = $derived({
    ...(sheet.cells[selected] || { value: '' }),
    validation: cellValidation(sheet, selected),
  });
  function commit(next: WorkbookContent) {
    if (headerBusy) return;
    next = syncTableTotals(content, clearRemovedColumnFormulas(content, next));
    next = {
      ...next,
      sheets: next.sheets.map((s) => {
        const arrays = s.arrayFormulas?.filter(
          (a) => !singleArray(a) || s.cells[a.anchor]?.value.startsWith('='),
        );
        return arrays?.length !== s.arrayFormulas?.length ? { ...s, arrayFormulas: arrays } : s;
      }),
    };
    for (const old of content.sheets) {
      const changed = next.sheets.find((s) => s.id === old.id);
      if (changed) {
        const arrayError = arrayEditError(old, changed);
        if (arrayError) {
          notify(arrayError);
          return;
        }
      }
      const error = changed && tableHeaderError(old, changed);
      if (error) {
        void commitHeaderEntry(next);
        return;
      }
    }
    recordChange(next);
  }
  function recordChange(next: WorkbookContent) {
    undo.current.push(content);
    if (undo.current.length > 50) undo.current.shift();
    redo.current = [];
    onChange(next);
  }
  async function commitHeaderEntry(proposed: WorkbookContent) {
    const before = content;
    headerBusy = true;
    try {
      const next = await applyHeaderEntry(before, proposed, onTableRename);
      if (!structureAlive) return;
      if (content !== before)
        throw Error('The workbook changed while its headers were updating. Please retry.');
      recordChange(next);
    } catch (error) {
      if (structureAlive) notify((error as Error).message);
    } finally {
      headerBusy = false;
      await tick();
      if (structureAlive && document.activeElement === document.body) grid.current?.focus();
    }
  }
  function updateCells(changes: Record<string, Cell>) {
    const blocked = Object.keys(changes).filter(
      (ref) => sheet.protected && sheet.cells[ref]?.locked !== false,
    );
    if (blocked.length) {
      notify('This worksheet is protected. Edit an unlocked input cell.');
      return;
    }
    for (const [ref, next] of Object.entries(changes)) {
      if (next.value === sheet.cells[ref]?.value) continue;
      const error = validationError(
        sheet.cells[ref] || { value: '' },
        next.value,
        sheet,
        content.sheets,
        ref,
      );
      if (error) {
        notify(`${ref}: ${error}`);
        return;
      }
    }
    for (const [ref, next] of Object.entries(changes)) {
      if (next.value !== sheet.cells[ref]?.value)
        changes[ref] = { ...next, dataType: undefined, cachedValue: undefined };
    }
    commit({
      ...content,
      sheets: content.sheets.map((s) =>
        s.id === sheet.id ? { ...s, cells: { ...s.cells, ...changes } } : s,
      ),
    });
  }
  function patchStyle(style: Partial<Cell>) {
    if ('format' in style && !('numFmt' in style)) style.numFmt = undefined;
    updateCells(
      Object.fromEntries(
        refs.map((ref) => [ref, { ...(sheet.cells[ref] || { value: '' }), ...style }]),
      ),
    );
  }
  function saveDraft(value = draft, array = false) {
    try {
      if (array || sheet.arrayFormulas?.some((a) => arrayContains(a.ref, selected))) {
        if (array && refs.length !== 1)
          throw Error('Multi-cell array entry is not supported yet. Select one cell.');
        commit(enterArrayFormula(content, sheet.id, selected, value, array));
        setEditing(false);
        return;
      }
      const calculated =
        tableRowEntry(content, sheet.id, selected, value) ||
        calculatedColumnEntry(content, sheet.id, selected, value);
      if (calculated) commit(calculated);
      else updateCells({ [selected]: { ...current, value } });
    } catch (error) {
      notify((error as Error).message);
    }
    setEditing(false);
  }
  function choose(ref: string, extend = false) {
    if (editing) saveDraft();
    const [r, c] = coordinates(ref);
    ref = layout.cell(r, c).master;
    setSelected(ref);
    if (!extend) setAnchor(ref);
    setEditing(false);
  }
  function move(dr: number, dc: number, extend = false) {
    const [r, c] = coordinates(selected),
      ri = layout.rows.indexOf(r),
      ci = layout.columns.indexOf(c);
    const nr = layout.rows[Math.max(0, Math.min(layout.rows.length - 1, ri + dr))],
      nc = layout.columns[Math.max(0, Math.min(layout.columns.length - 1, ci + dc))];
    choose(address(nr, nc), extend);
    const y = layout.offsets[layout.rows.indexOf(nr)];
    if (
      grid.current &&
      (y < grid.current.scrollTop || y > grid.current.scrollTop + grid.current.clientHeight - 90)
    )
      grid.current.scrollTop = Math.max(0, y - 90);
  }

  function onKey(e: KeyboardEvent) {
    if ((e.target as HTMLElement).closest('dialog')) return;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 't') {
      e.preventDefault();
      if (editing) saveDraft();
      tableOpen = true;
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
      e.preventDefault();
      linkOpen = true;
      return;
    }
    if ((e.ctrlKey || e.metaKey) && ['f', 'h'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      openSearch();
      return;
    }
    if (
      (e.target as HTMLElement).tagName === 'INPUT' ||
      (e.target as HTMLElement).tagName === 'TEXTAREA' ||
      (e.target as HTMLElement).tagName === 'SELECT'
    )
      return;
    if ((e.ctrlKey || e.metaKey) && e.key === '1') {
      e.preventDefault();
      numberFormatOpen = true;
    } else if (e.key.startsWith('Arrow')) {
      e.preventDefault();
      move(
        e.key === 'ArrowDown' ? 1 : e.key === 'ArrowUp' ? -1 : 0,
        e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : 0,
        e.shiftKey,
      );
    } else if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      if (sheet.protected && current.locked !== false) {
        notify('This cell is locked by worksheet protection.');
        return;
      }
      setDraft(current.value);
      setEditing(true);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      move(0, e.shiftKey ? -1 : 1);
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      updateCells(Object.fromEntries(refs.map((ref) => [ref, { ...sheet.cells[ref], value: '' }])));
    } else if ((e.ctrlKey || e.metaKey) && ['z', 'y'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      history(e.shiftKey || e.key.toLowerCase() === 'y');
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      if (sheet.protected && current.locked !== false) {
        notify('This cell is locked by worksheet protection.');
        return;
      }
      setDraft(e.key);
      setEditing(true);
    }
  }
  function history(forward: boolean) {
    const source = forward ? redo : undo,
      dest = forward ? undo : redo,
      previous = source.current.pop();
    if (previous) {
      dest.current.push(content);
      onChange(previous);
      setEditing(false);
    }
  }
  function paste(text: string) {
    const [r, c] = coordinates(selected),
      rows = text
        .replace(/\r\n/g, '\n')
        .replace(/\n$/, '')
        .split('\n')
        .map((row) => row.split('\t'));
    if (r + rows.length > 10000 || c + Math.max(...rows.map((row) => row.length)) > 256) {
      notify('The pasted range exceeds 10,000 rows or 256 columns.');
      return;
    }
    const changes: Record<string, Cell> = {};
    rows.forEach((row, dr) =>
      row.forEach((value, dc) => {
        const ref = address(r + dr, c + dc);
        changes[ref] = { ...sheet.cells[ref], value, dataType: undefined, cachedValue: undefined };
      }),
    );
    if (sheet.protected && Object.keys(changes).some((ref) => sheet.cells[ref]?.locked !== false)) {
      notify('This selection contains protected cells.');
      return;
    }
    for (const [ref, next] of Object.entries(changes)) {
      const error = validationError(
        sheet.cells[ref] || { value: '' },
        next.value,
        sheet,
        content.sheets,
        ref,
      );
      if (error) {
        notify(`${ref}: ${error}`);
        return;
      }
    }
    commit({
      ...content,
      sheets: content.sheets.map((s) =>
        s.id === sheet.id
          ? {
              ...s,
              rows: Math.max(s.rows, r + rows.length),
              cols: Math.max(s.cols, c + Math.max(...rows.map((row) => row.length))),
              cells: { ...s.cells, ...changes },
            }
          : s,
      ),
    });
  }
  function copyText() {
    const [ar, ac] = coordinates(anchor),
      [br, bc] = coordinates(selected);
    return Array.from({ length: Math.abs(ar - br) + 1 }, (_, r) =>
      Array.from({ length: Math.abs(ac - bc) + 1 }, (_, c) =>
        String(calc(sheet, address(Math.min(ar, br) + r, Math.min(ac, bc) + c))),
      ).join('\t'),
    ).join('\n');
  }
  function sort(descending: boolean) {
    const [ar, ac] = coordinates(anchor),
      [br, bc] = coordinates(selected),
      start = Math.min(ar, br),
      end = Math.max(ar, br),
      left = Math.min(ac, bc),
      right = Math.max(ac, bc);
    if (start === end) {
      notify('Select multiple rows with Shift + click, then sort the selected range.');
      return;
    }
    const rows = Array.from({ length: end - start + 1 }, (_, i) => start + i).sort((a, b) => {
      const av = calc(sheet, address(a, left)),
        bv = calc(sheet, address(b, left));
      return (
        (typeof av === 'number' && typeof bv === 'number'
          ? av - bv
          : String(av).localeCompare(String(bv), undefined, { numeric: true })) *
        (descending ? -1 : 1)
      );
    });
    const changes: Record<string, Cell> = {};
    rows.forEach((r, i) => {
      for (let c = left; c <= right; c++) {
        const source = sheet.cells[address(r, c)] || { value: '' };
        changes[address(start + i, c)] = {
          ...source,
          value: translateFormula(source.value, start + i - r, 0),
        };
      }
    });
    updateCells(changes);
  }
  function fillDown() {
    const [ar, ac] = coordinates(anchor),
      [br, bc] = coordinates(selected),
      top = Math.min(ar, br),
      bottom = Math.max(ar, br);
    if (top === bottom) {
      notify('Select the source row and rows below it with Shift + click.');
      return;
    }
    const changes: Record<string, Cell> = {};
    for (let c = Math.min(ac, bc); c <= Math.max(ac, bc); c++)
      for (let r = top + 1; r <= bottom; r++) {
        const source = sheet.cells[address(top, c)] || { value: '' };
        changes[address(r, c)] = { ...source, value: translateFormula(source.value, r - top, 0) };
      }
    updateCells(changes);
  }
  let layout = $derived(sheetLayout(sheet));
  let activeTable = $derived(tableAt(sheet, selected));
  let choices = $derived(validationChoices(current, sheet, content.sheets));
  let rowWindow = $derived(layout.window(scroll, gridHeight));
  function freeze(command: FreezeCommand) {
    try {
      const next = freezeSheet(sheet, command, selected);
      if (
        (next.frozenRows || 0) === (sheet.frozenRows || 0) &&
        (next.frozenColumns || 0) === (sheet.frozenColumns || 0)
      )
        return;
      commit({ ...content, sheets: content.sheets.map((s) => (s.id === sheet.id ? next : s)) });
      if (grid.current) {
        grid.current.scrollTop = 0;
        grid.current.scrollLeft = 0;
      }
      setScroll(0);
    } catch (error) {
      notify((error as Error).message);
    }
  }
  function changeGeometry(key: 'rowHeights' | 'columnWidths', index: number, value: number) {
    if (!Number.isFinite(value) || value < 8 || value > 1000) return;
    commit({
      ...content,
      sheets: content.sheets.map((s) =>
        s.id === sheet.id ? { ...s, [key]: { ...s[key], [index]: value } } : s,
      ),
    });
  }
  function changeRowVisibility(hide: boolean) {
    if (sheet.protected || sheet.filterMode) {
      notify(
        sheet.protected
          ? 'This worksheet is protected.'
          : 'Clear active filters before changing row visibility.',
      );
      return;
    }
    const hidden = new Set(sheet.hiddenRows);
    const [a] = coordinates(anchor),
      [b, col] = coordinates(selected);
    if (hide) for (let r = Math.min(a, b); r <= Math.max(a, b); r++) hidden.add(r);
    else hidden.clear();
    const visible = Array.from({ length: sheet.rows }, (_, r) => r).filter((r) => !hidden.has(r));
    if (!visible.length) {
      notify('Keep at least one row visible.');
      return;
    }
    commit({
      ...content,
      sheets: content.sheets.map((s) =>
        s.id === sheet.id ? { ...s, hiddenRows: [...hidden].sort((a, b) => a - b) } : s,
      ),
    });
    if (hidden.has(b)) {
      const ref = address(visible.find((r) => r > Math.max(a, b)) ?? visible.at(-1)!, col);
      setSelected(ref);
      setAnchor(ref);
    }
    setRename(false);
  }
  let selectedNumbers = $derived(
    refs.map((ref) => calc(sheet, ref)).filter((v): v is number => typeof v === 'number'),
  );
  let chartData = $derived(
    refs
      .map((ref) => ({ ref, value: calc(sheet, ref) }))
      .filter((v): v is { ref: string; value: number } => typeof v.value === 'number')
      .slice(0, 30),
  );
  let chartMax = $derived(Math.max(1, ...chartData.map((v) => Math.abs(v.value))));
</script>

<!-- svelte-ignore a11y_no_noninteractive_element_interactions (The editor handles keyboard navigation and editing for its focused grid.) -->
<div
  class="sheet-editor editor-body"
  inert={headerBusy}
  aria-busy={headerBusy}
  role="application"
  aria-label="Spreadsheet editor"
  onkeydown={onKey}
>
  <div class="ribbon-tabs">
    <button class="selected">Home</button><button
      class="ribbon-tab-action"
      disabled={sheet.protected || !!sheet.tableEditingBlocked}
      onclick={() => {
        if (editing) saveDraft();
        tableOpen = true;
      }}>{activeTable ? 'Table design' : 'Format as table'}</button
    ><button class="ribbon-tab-action" onclick={() => setChart(true)}
      ><ChartColumn size={15} />Chart selection</button
    ><button class="ribbon-tab-action" onclick={openFilter}>Filter range</button><button
      class="ribbon-tab-action"
      disabled={!sheet.autoFilters?.some((f) => f.columns.length) || sheet.protected}
      onclick={() => changeFilters(sheet.autoFilters || [])}>Reapply filters</button
    ><button
      class="ribbon-tab-action"
      disabled={!sheet.autoFilters?.some((f) => f.columns.length) || sheet.protected}
      onclick={() => changeFilters((sheet.autoFilters || []).map((f) => ({ ...f, columns: [] })))}
      >Clear filters</button
    ><span class="ribbon-spacer"></span><button
      class="ribbon-tab-action"
      onclick={() => (showFind ? setShowFind(false) : openSearch())}
      ><Search size={15} />Find cells</button
    >
  </div>
  <div class="ribbon grouped-ribbon" role="toolbar" aria-label="Spreadsheet formatting">
    <RibbonGroup label="Cells">
      <select
        aria-label="Insert or delete rows and columns"
        value=""
        disabled={structureBusy || sheet.protected}
        onchange={(e) => {
          const [action, axis] = e.currentTarget.value.split('-') as [
            StructureEdit['action'],
            StructureEdit['axis'],
          ];
          e.currentTarget.value = '';
          void changeStructure(axis, action);
        }}
      >
        <option value="" disabled>Rows and columns</option>
        <option value="insert-row">Insert rows</option>
        <option value="insert-column">Insert columns</option>
        <option value="delete-row">Delete rows</option>
        <option value="delete-column">Delete columns</option>
      </select>
      {#if structureBusy}<span role="status">Updating worksheet…</span>{/if}
    </RibbonGroup>
    <RibbonGroup label="Data"
      ><button class="ribbon-action" onclick={() => (validationOpen = true)}>Data validation</button
      ></RibbonGroup
    >
    <RibbonGroup label="Links">
      <button class="ribbon-action" onclick={() => (linkOpen = true)}>Hyperlink</button>
      <button class="ribbon-action" disabled={!current.hyperlink} onclick={followLink}
        >Open link</button
      >
      <button
        class="ribbon-action"
        disabled={!refs.some((ref) => sheet.cells[ref]?.hyperlink)}
        onclick={clearLinks}>Remove links</button
      >
    </RibbonGroup>
    <RibbonGroup label="History" kind="history">
      <Tool label="Undo" disabled={!content || !undo.current.length} onclick={() => history(false)}
        ><Undo2 /></Tool
      ><Tool label="Redo" disabled={!content || !redo.current.length} onclick={() => history(true)}
        ><Redo2 /></Tool
      >
    </RibbonGroup>
    <RibbonGroup label="Font" kind="cell-font"
      ><div class="ribbon-row">
        <FontPicker
          value={current.fontFamily || defaultFont}
          onchange={(fontFamily) => patchStyle({ fontFamily })}
        />
        <input
          class="font-size-input"
          aria-label="Font size"
          title="Font size in points"
          type="number"
          min="1"
          max="409"
          value={current.fontSize || 11}
          onchange={(e) => {
            const fontSize = e.currentTarget.valueAsNumber;
            if (Number.isFinite(fontSize) && fontSize >= 1 && fontSize <= 409)
              patchStyle({ fontSize });
          }}
        />
      </div>
      <div class="ribbon-row">
        <Tool
          label="Bold"
          active={current.bold || false}
          onclick={() => patchStyle({ bold: !current.bold })}><Bold /></Tool
        ><Tool
          label="Italic"
          active={current.italic || false}
          onclick={() => patchStyle({ italic: !current.italic })}><Italic /></Tool
        ><Tool
          label="Underline"
          active={current.underline || false}
          onclick={() => patchStyle({ underline: !current.underline })}><Underline /></Tool
        >
        <label class="color-tool" title="Text color"
          >A<input
            aria-label="Cell text color"
            type="color"
            value={current.color || '#263d34'}
            oninput={(e) => patchStyle({ color: e.currentTarget.value })}
          /></label
        ><label class="color-tool" title="Cell fill"
          ><PaintBucket size={17} /><input
            aria-label="Cell fill"
            type="color"
            value={current.fill || '#ffffff'}
            oninput={(e) => patchStyle({ fill: e.currentTarget.value })}
          /></label
        >
      </div></RibbonGroup
    >
    <RibbonGroup label="Alignment" kind="cell-alignment">
      {#each ['left', 'center', 'right'] as const as align, i}{@const Icon = [
          AlignLeft,
          AlignCenter,
          AlignRight,
        ][i]}<Tool
          label={`Align ${align}`}
          active={current.align === align}
          onclick={() => patchStyle({ align })}><Icon /></Tool
        >{/each}
      <Tool
        label="Wrap text"
        active={current.wrap || false}
        onclick={() => patchStyle({ wrap: !current.wrap })}><WrapText /></Tool
      >
      <select
        aria-label="Vertical alignment"
        value={current.vertical || 'bottom'}
        onchange={(e) => patchStyle({ vertical: e.currentTarget.value as Cell['vertical'] })}
      >
        <option value="top">Align top</option><option value="middle">Align middle</option><option
          value="bottom">Align bottom</option
        >
      </select></RibbonGroup
    >
    <RibbonGroup label="Number" kind="cell-number"
      ><select
        aria-label="Number format"
        value={current.numFmt && current.numFmt !== 'General'
          ? 'custom'
          : current.format || 'general'}
        onchange={(e) => patchStyle({ format: e.currentTarget.value as Cell['format'] })}
        >{#if current.numFmt && current.numFmt !== 'General'}<option value="custom" disabled
            >Custom format</option
          >{/if}<option value="general">General</option><option value="number">Number · 0.00</option
        ><option value="currency">Currency · $</option><option value="percent"
          >Percentage · %</option
        ></select
      >
      <button
        class="text-tool"
        title="Format cells (Ctrl+1)"
        onclick={() => (numberFormatOpen = true)}>Format cells</button
      >
    </RibbonGroup>
    <RibbonGroup label="Editing" kind="cell-editing">
      <Tool label="Sort ascending" onclick={() => sort(false)}><ArrowDownAZ /></Tool><Tool
        label="Sort descending"
        onclick={() => sort(true)}><ArrowUpAZ /></Tool
      ><Tool label="Fill down" onclick={fillDown}><ArrowDownToLine /></Tool><Tool
        label="AutoSum above"
        onclick={() => {
          const [r, c] = coordinates(selected);
          if (r)
            updateCells({
              [selected]: { ...current, value: `=SUM(${address(0, c)}:${address(r - 1, c)})` },
            });
        }}><Sigma /></Tool
      ><Tool
        label="Copy selected cells"
        onclick={async () => {
          try {
            await navigator.clipboard.writeText(copyText());
            notify('Selection copied.');
          } catch {
            notify('Use Ctrl+C with the grid focused to copy.');
          }
        }}><Copy /></Tool
      >
    </RibbonGroup>
    <RibbonGroup label="View" kind="sheet-view">
      <select
        aria-label="Freeze panes"
        value=""
        onchange={(e) => {
          freeze(e.currentTarget.value as FreezeCommand);
          e.currentTarget.value = '';
        }}
      >
        <option value="" disabled
          >{sheet.frozenRows || sheet.frozenColumns ? 'Panes frozen' : 'Choose freeze area'}</option
        >
        <option value="selection">Freeze at active cell</option>
        <option value="row">Freeze top row</option>
        <option value="column">Freeze first column</option>
        <option value="none" disabled={!sheet.frozenRows && !sheet.frozenColumns}
          >Unfreeze panes</option
        >
      </select>
      {#if sheet.frozenRows || sheet.frozenColumns}<span
          >{sheet.frozenRows || 0} rows · {sheet.frozenColumns || 0} columns frozen</span
        >{/if}
    </RibbonGroup>
  </div>
  <div class="formula-bar">
    <span class="cell-address">{selected}</span><span class="fx">ƒx</span><input
      aria-label="Formula bar"
      readonly={sheet.protected && current.locked !== false}
      value={editing
        ? draft
        : sheet.arrayFormulas?.some((a) => arrayContains(a.ref, selected))
          ? `{${current.value}}`
          : current.value}
      onfocus={() => {
        setDraft(current.value);
        setEditing(true);
      }}
      oninput={(e) => setDraft(e.currentTarget.value)}
      onblur={() => {
        if (editing) saveDraft();
      }}
      onkeydown={(e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          saveDraft(draft, (e.ctrlKey || e.metaKey) && e.shiftKey);
          grid.current?.focus();
        }
        if (e.key === 'Escape') {
          setEditing(false);
          grid.current?.focus();
        }
      }}
      placeholder="Enter a value or a formula, like =SUM(A1:A10)"
    />
    {#if choices}<select
        aria-label="Cell choices"
        value={current.value}
        disabled={sheet.protected && current.locked !== false}
        onchange={(e) => saveDraft(e.currentTarget.value)}
      >
        <option value="">Choose a value</option>{#each choices as choice}<option value={choice}
            >{choice}</option
          >{/each}
      </select>{/if}
  </div>
  <div
    class="cell-note"
    class:has-note={!!(
      current.note ||
      (current.validation?.showInputMessage !== false && current.validation?.prompt)
    )}
    role="note"
    title={current.note ||
      (current.validation?.showInputMessage !== false ? current.validation?.prompt : '')}
  >
    {current.note ||
      (current.validation?.showInputMessage !== false ? current.validation?.prompt : '') ||
      ''}
  </div>
  {#if showFind}<SearchBar
      bind:query={search}
      bind:replacement
      bind:matchCase={searchCase}
      bind:whole={searchWhole}
      wholeLabel="Match entire cell contents"
      findLabel="Find cell value"
      placeholder={searchWildcards ? 'Find text (*, ?, ~ wildcards)' : 'Find literal text'}
      status={searchStatus}
      onfind={findCell}
      onreplace={replaceCells}
      onclose={() => {
        setShowFind(false);
        grid.current?.focus();
      }}
    >
      <label
        >Within <select aria-label="Search within" bind:value={searchScope}
          ><option value="sheet">Sheet</option><option value="workbook">Workbook</option><option
            value="selection">Selection</option
          ></select
        ></label
      >
      <label
        >Look in <select aria-label="Look in" bind:value={searchLookIn}
          ><option value="formulas">Formulas</option><option value="values">Values</option><option
            value="notes">Notes</option
          ></select
        ></label
      >
      <label><input type="checkbox" bind:checked={searchColumns} />Search by columns</label>
      <label><input type="checkbox" bind:checked={searchWildcards} />Use wildcards</label>
      <span>Visible cells only</span>
    </SearchBar>{/if}
  <!-- svelte-ignore a11y_no_noninteractive_tabindex (The grid region receives navigation and clipboard shortcuts.) -->
  <div
    class="grid-scroll"
    bind:this={grid.current}
    tabindex={0}
    role="region"
    aria-label="Spreadsheet grid"
    onscroll={(e) => setScroll(e.currentTarget.scrollTop)}
    onpaste={(e) => {
      if ((e.target as HTMLElement).tagName === 'INPUT') return;
      e.preventDefault();
      paste(e.clipboardData!.getData('text/plain'));
    }}
    oncopy={(e) => {
      if ((e.target as HTMLElement).tagName === 'INPUT') return;
      e.preventDefault();
      e.clipboardData!.setData('text/plain', copyText());
    }}
  >
    <table
      class="spreadsheet"
      class:no-gridlines={sheet.gridLines === false}
      style:width={`${layout.columnOffsets.at(-1)}px`}
      role="grid"
      aria-rowcount={sheet.rows}
      aria-colcount={sheet.cols}
    >
      <colgroup
        ><col style={css({ width: 46 })} />{#each layout.columns as c}<col
            style={css({ width: layout.width(c) })}
          />{/each}</colgroup
      ><thead
        ><tr
          ><th class="corner"></th>{#each layout.columns as c}<th
              scope="col"
              style={css({
                position: c < (sheet.frozenColumns || 0) ? 'sticky' : undefined,
                left:
                  c < (sheet.frozenColumns || 0)
                    ? layout.columnOffsets[layout.columns.indexOf(c)]
                    : undefined,
                zIndex: c < (sheet.frozenColumns || 0) ? 2 : undefined,
              })}>{colName(c)}</th
            >{/each}</tr
        ></thead
      ><tbody
        >{#each rowWindow.rows as r (r)}{#if r === rowWindow.bodyStart && rowWindow.top > 0}<tr
              aria-hidden="true"
              ><td colspan={sheet.cols + 1} style={css({ height: rowWindow.top, padding: 0 })}
              ></td></tr
            >{/if}<tr aria-rowindex={r + 1} style:height={`${layout.height(r)}px`}
            ><th
              scope="row"
              style={css({
                top:
                  r < (sheet.frozenRows || 0)
                    ? 27 + layout.offsets[layout.rows.indexOf(r)]
                    : undefined,
                zIndex: r < (sheet.frozenRows || 0) ? 5 : undefined,
              })}>{r + 1}</th
            >{#each layout.columns as c}{@const ref = address(r, c)}{@const cell =
                sheet.cells[ref]}{@const tableStyle = tableCellStyle(sheet, ref)}{@const value =
                calc(sheet, ref)}{@const active = selected === ref}{@const geometry = layout.cell(
                r,
                c,
              )}{#if !geometry.hidden}<td
                  rowspan={geometry.rowspan}
                  colspan={geometry.colspan}
                  role="gridcell"
                  aria-label={ref}
                  aria-selected={selectedSet.has(ref)}
                  class={`${active ? 'active-cell' : ''} ${selectedSet.has(ref) ? 'in-selection' : ''} ${typeof value === 'string' && value.startsWith('#') ? 'cell-error' : ''}`}
                  style={css({
                    fontWeight: (cell?.bold ?? tableStyle.bold) ? 650 : undefined,
                    fontStyle: cell?.italic ? 'italic' : undefined,
                    fontFamily: fontStack(cell?.fontFamily),
                    fontSize: (cell?.fontSize || 11) + 'pt',
                    textDecoration: cell?.underline ? 'underline' : undefined,
                    height: geometry.height,
                    whiteSpace: cell?.wrap ? 'pre-wrap' : 'pre',
                    verticalAlign: cell?.vertical || 'bottom',
                    borderTop: cell?.borders?.top,
                    borderBottom: cell?.borders?.bottom,
                    borderLeft: cell?.borders?.left,
                    borderRight: cell?.borders?.right,
                    position:
                      r < (sheet.frozenRows || 0) || c < (sheet.frozenColumns || 0)
                        ? 'sticky'
                        : undefined,
                    top:
                      r < (sheet.frozenRows || 0)
                        ? 27 + layout.offsets[layout.rows.indexOf(r)]
                        : undefined,
                    left:
                      c < (sheet.frozenColumns || 0)
                        ? layout.columnOffsets[layout.columns.indexOf(c)]
                        : undefined,
                    zIndex:
                      r < (sheet.frozenRows || 0)
                        ? c < (sheet.frozenColumns || 0)
                          ? 4
                          : 2
                        : c < (sheet.frozenColumns || 0)
                          ? 1
                          : undefined,
                    color: cell?.color || tableStyle.color,
                    backgroundColor:
                      cell?.fill ||
                      tableStyle.fill ||
                      (r < (sheet.frozenRows || 0) || c < (sheet.frozenColumns || 0)
                        ? '#fff'
                        : undefined),
                    textAlign: cell?.align || (typeof value === 'number' ? 'right' : 'left'),
                  })}
                  onmousedown={(e) => {
                    if (e.button !== 0 || (e.target as HTMLElement).tagName === 'INPUT') return;
                    e.preventDefault();
                    choose(ref, e.shiftKey);
                    grid.current?.focus({ preventScroll: true });
                  }}
                  onmouseenter={(e) => {
                    if (e.buttons === 1 && !editing) setSelected(ref);
                  }}
                  ondblclick={() => {
                    if (sheet.protected && cell?.locked !== false) {
                      notify('This cell is locked by worksheet protection.');
                      return;
                    }
                    setDraft(cell?.value || '');
                    setEditing(true);
                  }}
                >
                  {#if active && editing && document.activeElement?.getAttribute('aria-label') !== 'Formula bar'}<input
                      use:focusOnMount
                      aria-label={`Edit ${ref}`}
                      value={draft}
                      oninput={(e) => setDraft(e.currentTarget.value)}
                      onblur={() => {
                        if (editing) saveDraft();
                      }}
                      onkeydown={(e) => {
                        if (e.key === 'Enter' || e.key === 'Tab') {
                          e.preventDefault();
                          const arrayEntry =
                            e.key === 'Enter' && (e.ctrlKey || e.metaKey) && e.shiftKey;
                          saveDraft(draft, arrayEntry);
                          setSelected(
                            address(
                              Math.min(
                                sheet.rows - 1,
                                r + (e.key === 'Enter' && !arrayEntry ? 1 : 0),
                              ),
                              Math.min(sheet.cols - 1, c + (e.key === 'Tab' ? 1 : 0)),
                            ),
                          );
                          setAnchor(
                            address(
                              Math.min(
                                sheet.rows - 1,
                                r + (e.key === 'Enter' && !arrayEntry ? 1 : 0),
                              ),
                              Math.min(sheet.cols - 1, c + (e.key === 'Tab' ? 1 : 0)),
                            ),
                          );
                          grid.current?.focus();
                        }
                        if (e.key === 'Escape') {
                          setEditing(false);
                          grid.current?.focus();
                        }
                      }}
                    />{:else}<div
                      class="cell-display"
                      style:max-height={`${geometry.height - 1}px`}
                      title={value === '#UNSUPPORTED!'
                        ? 'This formula is not supported by the browser calculator. The original result is preserved in exports.'
                        : cell?.hyperlinkTooltip || cell?.hyperlink || undefined}
                    >
                      {displayValue(value, cell)}
                    </div>{/if}
                </td>{/if}{/each}</tr
          >{/each}{#if rowWindow.bottom > 0}<tr aria-hidden="true"
            ><td colspan={sheet.cols + 1} style={css({ height: rowWindow.bottom, padding: 0 })}
            ></td></tr
          >{/if}</tbody
      >
    </table>
    {#each sheet.images || [] as image}<img
        class="sheet-image"
        src={image.src}
        alt="Imported worksheet drawing"
        style={css({ left: image.x + 46, top: image.y + 27, width: image.w, height: image.h })}
      />{/each}
  </div>
  <div class="sheet-tabs">
    {#each content.sheets.filter((s) => !s.state || s.state === 'visible') as s}<button
        class={s.id === sheet.id ? 'selected' : ''}
        onclick={() => {
          if (editing) saveDraft();
          setSheetId(s.id);
          setSelected('A1');
          setAnchor('A1');
          setScroll(0);
          if (grid.current) grid.current.scrollTop = 0;
        }}
        ondblclick={() => {
          setSheetName(s.name);
          setRename(true);
        }}>{s.name}</button
      >{/each}<Tool
      label="Add sheet"
      disabled={!!sheet.sourcePath}
      onclick={() => {
        let n = content.sheets.length + 1;
        while (content.sheets.some((s) => s.name === `Sheet ${n}`)) n++;
        const id = uid();
        commit({
          ...content,
          sheets: [...content.sheets, { id, name: `Sheet ${n}`, cells: {}, rows: 100, cols: 26 }],
        });
        setSheetId(id);
        setSelected('A1');
        setAnchor('A1');
      }}><Plus /></Tool
    ><span class="ribbon-spacer"></span><button
      onclick={() => {
        setSheetName(sheet.name);
        setRename(true);
      }}>Sheet settings</button
    ><button
      onclick={() =>
        commit({
          ...content,
          sheets: content.sheets.map((s) =>
            s.id === sheet.id ? { ...s, rows: Math.min(10000, s.rows + 100) } : s,
          ),
        })}>+ 100 rows</button
    >
  </div>
  <div class="editor-status">
    <span>{refs.length === 1 ? 'Ready' : `${refs.length} cells selected`}</span><span
      class="ribbon-spacer"
    ></span>{#if selectedNumbers.length > 0}<span
        >Average: {displayValue(
          selectedNumbers.reduce((a, b) => a + b, 0) / selectedNumbers.length,
        )}</span
      ><span>Count: {selectedNumbers.length}</span><span
        >Sum: {displayValue(selectedNumbers.reduce((a, b) => a + b, 0))}</span
      >{/if}
  </div>
  {#if validationOpen}<ValidationDialog
      rule={cellValidation(sheet, refs[0])}
      count={refs.length}
      onclose={() => (validationOpen = false)}
      onapply={changeValidation}
    />{/if}
  {#if linkOpen}<SheetLinkDialog
      cell={current}
      onclose={() => (linkOpen = false)}
      onapply={applyLink}
    />{/if}
  {#if tableOpen}<SheetTableDialog
      table={activeTable}
      columnIndex={activeTable ? coordinates(selected)[1] - tableBounds(activeTable.ref).c : 0}
      initialRange={`${address(Math.min(coordinates(anchor)[0], coordinates(selected)[0]), Math.min(coordinates(anchor)[1], coordinates(selected)[1]))}:${address(Math.max(coordinates(anchor)[0], coordinates(selected)[0]), Math.max(coordinates(anchor)[1], coordinates(selected)[1]))}`}
      initialName={nextTableName(content.sheets)}
      onclose={closeTable}
      onapply={applyTable}
    />{/if}
  {#if numberFormatOpen}<NumberFormatDialog
      initialCode={current.numFmt ||
        { general: 'General', number: '0.00', currency: '$#,##0.00', percent: '0.00%' }[
          current.format || 'general'
        ]}
      value={calc(sheet, selected)}
      date1904={sheet.date1904}
      count={refs.length}
      onclose={() => (numberFormatOpen = false)}
      onapply={(numFmt) => {
        patchStyle({ format: 'general', numFmt });
        numberFormatOpen = false;
      }}
    />{/if}
  {#if editingFilter}<SheetFilterDialog
      {sheet}
      sheets={content.sheets}
      filter={editingFilter}
      onclose={() => (editingFilter = null)}
      onapply={(filter) =>
        changeFilters([
          ...(sheet.autoFilters || []).filter((f) => f.sourcePath !== filter.sourcePath),
          filter,
        ])}
    />{/if}
  {#if rename}<Modal title="Sheet settings" onclose={() => setRename(false)}
      ><label class="field"
        ><span>Sheet name</span><input
          disabled={!!sheet.sourcePath}
          value={sheetName}
          maxLength={31}
          oninput={(e) => setSheetName(e.currentTarget.value)}
        /></label
      >
      <div class="inspector-row">
        <label class="field"
          ><span>Selected column width (px)</span><input
            aria-label="Column width"
            type="number"
            min="8"
            max="1000"
            value={layout.width(coordinates(selected)[1])}
            onchange={(e) =>
              changeGeometry(
                'columnWidths',
                coordinates(selected)[1],
                e.currentTarget.valueAsNumber,
              )}
          /></label
        >
        <label class="field"
          ><span>Selected row height (px)</span><input
            aria-label="Row height"
            type="number"
            min="8"
            max="1000"
            value={layout.height(coordinates(selected)[0])}
            onchange={(e) =>
              changeGeometry('rowHeights', coordinates(selected)[0], e.currentTarget.valueAsNumber)}
          /></label
        >
      </div>
      {#if sheet.sourcePath}<p class="muted">
          Original sheet names and order are retained to preserve workbook references. Cell edits
          and row/column sizing export into the original package.
        </p>{/if}
      <h3>Worksheet visibility</h3>
      <div class="modal-actions">
        <button
          disabled={sheet.protected || sheet.filterMode}
          onclick={() => changeRowVisibility(true)}>Hide selected rows</button
        >
        <button
          disabled={!sheet.hiddenRows?.length || sheet.protected || sheet.filterMode}
          onclick={() => changeRowVisibility(false)}>Unhide all rows</button
        >
      </div>
      {#each content.sheets as s}<label class="field"
          ><span>{s.name}</span><select
            aria-label={`Visibility of ${s.name}`}
            value={s.state || 'visible'}
            disabled={s.state === 'veryHidden'}
            onchange={(e) => {
              const state = e.currentTarget.value as 'visible' | 'hidden';
              if (
                state === 'hidden' &&
                content.sheets.filter((s) => !s.state || s.state === 'visible').length === 1
              ) {
                notify('Keep at least one worksheet visible.');
                return;
              }
              commit({
                ...content,
                sheets: content.sheets.map((other) =>
                  other.id === s.id ? { ...other, state } : other,
                ),
              });
              if (s.id === sheet.id && state === 'hidden')
                setSheetId(
                  content.sheets.find(
                    (other) => other.id !== s.id && (!other.state || other.state === 'visible'),
                  )!.id,
                );
            }}
            ><option value="visible">Visible</option><option value="hidden">Hidden</option
            >{#if s.state === 'veryHidden'}<option value="veryHidden">Very hidden</option
              >{/if}</select
          ></label
        >{/each}
      <div class="modal-actions">
        <button
          class="danger"
          disabled={content.sheets.length === 1 || !!sheet.sourcePath}
          onclick={() => {
            commit({ ...content, sheets: content.sheets.filter((s) => s.id !== sheet.id) });
            setSheetId(content.sheets.find((s) => s.id !== sheet.id)!.id);
            setRename(false);
          }}><Trash2 size={15} />Delete sheet</button
        ><button
          class="primary"
          onclick={() => {
            if (
              !sheetName.trim() ||
              /[\\/*?:\[\]]/.test(sheetName) ||
              content.sheets.some(
                (s) => s.id !== sheet.id && s.name.toLowerCase() === sheetName.trim().toLowerCase(),
              )
            ) {
              notify('Use a unique sheet name without \\ / * ? : [ ].');
              return;
            }
            if (
              content.sheets.some((s) =>
                Object.values(s.cells).some(
                  (c) => c.value.startsWith('=') && c.value.includes('!'),
                ),
              )
            ) {
              notify('Renaming sheets with cross-sheet formulas is not supported yet.');
              return;
            }
            commit({
              ...content,
              sheets: content.sheets.map((s) =>
                s.id === sheet.id ? { ...s, name: sheetName.trim() } : s,
              ),
            });
            setRename(false);
          }}>Save name</button
        >
      </div></Modal
    >{/if}{#if chart}<Modal title="Chart your selection" onclose={() => setChart(false)} wide
      >{#if chartData.length}<div class="bar-chart">
          {#each chartData as d}<div>
              <span>{d.ref}</span><i
                style={css({
                  width: `${Math.max(1, (Math.abs(d.value) / chartMax) * 75)}%`,
                  background: d.value < 0 ? '#bd7354' : undefined,
                })}
              ></i><strong>{displayValue(d.value)}</strong>
            </div>{/each}
        </div>{:else}<p>Select cells containing numbers to see a bar chart.</p>{/if}
      <p class="muted">
        A preview of up to 30 selected numeric cells. Chart previews are not embedded in the
        workbook.
      </p></Modal
    >{/if}
</div>
{#if headerBusy}<div class="muted" role="status">Updating table headings and references…</div>{/if}
