<script lang="ts">
  import { untrack } from 'svelte';
  import Modal from './Modal.svelte';
  import {
    defaultTableStyle,
    tableStyles,
    type SheetTable,
    type TableStyle,
  } from '../sheet-tables';
  let {
    table,
    columnIndex = 0,
    initialRange,
    initialName,
    onclose,
    onapply,
  }: {
    table?: SheetTable;
    columnIndex?: number;
    initialRange: string;
    initialName: string;
    onclose: () => void;
    onapply: (
      name: string,
      ref: string,
      style: TableStyle,
      column?: { index: number; name: string },
    ) => void | Promise<void>;
  } = $props();
  const initial = untrack(() => ({ table, initialRange, initialName, columnIndex }));
  let name = $state(initial.table?.name || initial.initialName),
    range = $state(initial.table?.ref || initial.initialRange);
  let style = $state({ ...(initial.table?.style || defaultTableStyle) }),
    error = $state('');
  let selectedColumn = $state(initial.columnIndex),
    columnName = $state(initial.table?.columns[initial.columnIndex] || ''),
    busy = $state(false);
  async function apply() {
    if (busy) return;
    busy = true;
    error = '';
    try {
      await onapply(
        name,
        range,
        style,
        table ? { index: selectedColumn, name: columnName } : undefined,
      );
    } catch (e) {
      error = (e as Error).message;
    } finally {
      busy = false;
    }
  }
</script>

<Modal title={table ? 'Table design' : 'Create table'} {onclose}>
  <p class="muted">
    {table
      ? 'Change the name, columns or appearance. Renaming updates structured references throughout the workbook.'
      : 'The first row becomes the table header. Blank and duplicate headings receive unique names.'}
  </p>
  <div class="table-fields">
    <label class="field"
      ><span>Table name</span><input
        aria-label="Table name"
        bind:value={name}
        maxlength="255"
      /></label
    >
    <label class="field"
      ><span>Table range</span><input aria-label="Table range" bind:value={range} /></label
    >
  </div>
  {#if table}<div class="table-fields">
      <label class="field"
        ><span>Column</span><select
          aria-label="Table column"
          bind:value={selectedColumn}
          onchange={() => (columnName = table!.columns[selectedColumn])}
        >
          {#each table.columns as heading, index}<option value={index}>{heading}</option>{/each}
        </select></label
      >
      <label class="field"
        ><span>Column name</span><input
          aria-label="Column name"
          bind:value={columnName}
          maxlength="255"
        /></label
      >
    </div>{/if}
  <label class="field"
    ><span>Table style</span><select aria-label="Table style" bind:value={style.name}>
      {#if !tableStyles.includes(style.name)}<option value={style.name}
          >{style.name} (original)</option
        >{/if}
      {#each tableStyles as choice, i}<option value={choice}
          >{choice ? `Medium ${i + 2} · Accent ${i + 1}` : 'None'}</option
        >{/each}
    </select></label
  >
  <div class="table-options">
    <label><input type="checkbox" bind:checked={style.rowStripes} />Banded rows</label>
    <label><input type="checkbox" bind:checked={style.columnStripes} />Banded columns</label>
    <label><input type="checkbox" bind:checked={style.firstColumn} />First column</label>
    <label><input type="checkbox" bind:checked={style.lastColumn} />Last column</label>
  </div>
  {#if table?.resizeBlocked}<p class="muted">{table.resizeBlocked}</p>{/if}
  {#if error}<p role="alert">{error}</p>{/if}
  <div class="dialog-actions">
    <button onclick={onclose} disabled={busy}>Cancel</button><button
      class="primary"
      onclick={apply}
      disabled={busy}>{busy ? 'Applying…' : table ? 'Apply' : 'Create table'}</button
    >
  </div>
</Modal>

<style>
  .dialog-actions {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: 20px;
  }
  .table-fields,
  .table-options {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 12px;
    margin: 16px 0;
  }
  .table-options label {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .table-options input {
    width: 16px;
    height: 16px;
    flex: none;
  }
  .field input {
    min-width: 0;
    width: 100%;
  }
  @media (max-width: 480px) {
    .table-fields {
      grid-template-columns: 1fr;
    }
  }
</style>
