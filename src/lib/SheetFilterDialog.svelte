<script lang="ts">
  import { untrack } from 'svelte';
  import Modal from './Modal.svelte';
  import type { Sheet } from '../model';
  import type { SheetFilter } from '../sheet-filters';
  import { address, calculator, colName, coordinates, displayValue } from '../formulas';
  let {
    sheet,
    sheets,
    filter,
    onapply,
    onclose,
  }: {
    sheet: Sheet;
    sheets: Sheet[];
    filter: SheetFilter;
    onapply: (filter: SheetFilter) => void;
    onclose: () => void;
  } = $props();
  const setup = untrack(() => {
    const [a, b = a] = filter.ref.split(':');
    const [r, c] = coordinates(a),
      [rr, cc] = coordinates(b);
    return { r, c, rr, cc, calc: calculator(sheets) };
  });
  let col = $state(0),
    mode = $state('values'),
    value = $state(''),
    selected = $state<string[]>([]),
    search = $state('');
  let values = $derived(
    [
      ...new Set(
        Array.from({ length: setup.rr - setup.r }, (_, i) => {
          const ref = address(setup.r + i + 1, setup.c + col);
          return displayValue(setup.calc(sheet, ref), sheet.cells[ref]);
        }),
      ),
    ].sort((a, b) => a.localeCompare(b)),
  );
  let visible = $derived(
    values.filter((v) => v.toLowerCase().includes(search.toLowerCase())).slice(0, 200),
  );
  function selectColumn(index: number) {
    col = index;
    search = '';
    const current = filter.columns.find((c) => c.col === index);
    if (current?.custom?.length === 1) {
      mode = current.custom[0].operator;
      value = current.custom[0].value;
    } else {
      mode = 'values';
      value = '';
    }
    const all = Array.from({ length: setup.rr - setup.r }, (_, i) => {
      const ref = address(setup.r + i + 1, setup.c + index);
      return displayValue(setup.calc(sheet, ref), sheet.cells[ref]);
    });
    selected = current?.values
      ? [...current.values, ...(current.blank ? [''] : [])]
      : [...new Set(all)];
  }
  untrack(() => selectColumn(0));
  function apply(clear = false) {
    let columns = filter.columns.filter((c) => c.col !== col);
    if (!clear && !(mode === 'values' && values.every((v) => selected.includes(v))))
      columns.push(
        mode === 'values'
          ? { col, values: selected.filter((v) => v !== ''), blank: selected.includes('') }
          : { col, custom: [{ operator: mode, value }] },
      );
    onapply({ ...filter, columns: columns.sort((a, b) => a.col - b.col) });
  }
</script>

<Modal title="Filter range" {onclose}>
  <p class="muted">{filter.ref} · The first row contains column headings.</p>
  <label class="field"
    ><span>Column</span><select
      aria-label="Filter column"
      value={col}
      onchange={(e) => selectColumn(Number(e.currentTarget.value))}
    >
      {#each Array.from({ length: setup.cc - setup.c + 1 }, (_, i) => i) as i}<option value={i}
          >{sheet.cells[address(setup.r, setup.c + i)]?.value || colName(setup.c + i)}</option
        >{/each}
    </select></label
  >
  <label class="field"
    ><span>Match</span><select aria-label="Filter match" bind:value={mode}>
      <option value="values">Selected values</option><option value="equal">Equals</option><option
        value="notEqual">Does not equal</option
      ><option value="greaterThan">Greater than</option><option value="greaterThanOrEqual"
        >Greater than or equal</option
      ><option value="lessThan">Less than</option><option value="lessThanOrEqual"
        >Less than or equal</option
      >
    </select></label
  >
  {#if mode === 'values'}
    <input aria-label="Search filter values" placeholder="Search values" bind:value={search} />
    <div class="modal-actions">
      <button onclick={() => (selected = [...values])}>Select all values</button><button
        onclick={() => (selected = [])}>Clear selection</button
      >
    </div>
    <div class="filter-values">
      {#each visible as item}<label
          ><input
            type="checkbox"
            checked={selected.includes(item)}
            onchange={(e) =>
              (selected = e.currentTarget.checked
                ? [...selected, item]
                : selected.filter((v) => v !== item))}
          /><span>{item || '(Blanks)'}</span></label
        >{/each}
    </div>
    {#if values.length > 200}<p class="muted">
        Showing up to 200 values. Search to find more.
      </p>{/if}
  {:else}<label class="field"
      ><span>Value</span><input aria-label="Filter value" bind:value maxlength="1000" /></label
    >
    <p class="muted">Text comparisons support * and ? wildcards. Use ~ to escape them.</p>{/if}
  <div class="modal-actions">
    <button onclick={() => apply(true)}>Clear this column</button><button onclick={onclose}
      >Cancel</button
    ><button class="primary" onclick={() => apply()}>Apply filter</button>
  </div>
</Modal>

<style>
  .filter-values {
    max-height: 220px;
    overflow: auto;
    border: 1px solid var(--border, #e6e9e2);
    padding: 8px;
    border-radius: 6px;
  }
  .filter-values label {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 5px;
  }
  .filter-values span {
    overflow-wrap: anywhere;
  }
</style>
