<script lang="ts">
  import { untrack } from 'svelte';
  import Modal from './Modal.svelte';
  import { formatNumber } from '../number-format';
  import type { Value } from '../formulas';
  let {
    initialCode,
    value,
    date1904 = false,
    count,
    onapply,
    onclose,
  }: {
    initialCode: string;
    value: Value;
    date1904?: boolean;
    count: number;
    onapply: (code: string) => void;
    onclose: () => void;
  } = $props();
  let code = $state(untrack(() => initialCode));
  const presets = [
    ['General', 'General'],
    ['Number', '#,##0.00'],
    ['Currency', '$#,##0.00'],
    ['Accounting', '#,##0.00;(#,##0.00);"—"'],
    ['Percentage', '0.00%'],
    ['Date', 'dd.mm.yyyy'],
    ['Time', 'hh:mm:ss'],
    ['Elapsed time', '[h]:mm:ss'],
    ['Fraction', '# ??/??'],
    ['Scientific', '0.00E+00'],
    ['Text', '@'],
  ];
  let preview = $derived.by(() => {
    try {
      formatNumber(code, 1234.56, date1904);
      return { text: formatNumber(code, value, date1904), error: '' };
    } catch {
      return { text: '', error: 'This format is invalid or exceeds the supported format limits.' };
    }
  });
</script>

<Modal title="Format cells" {onclose}>
  <p class="muted">
    Number formatting for {count === 1 ? 'the selected cell' : `${count} selected cells`}.
  </p>
  <label class="field"
    ><span>Category</span>
    <select
      aria-label="Format category"
      value={presets.some((p) => p[1] === code) ? code : 'custom'}
      onchange={(e) => {
        if (e.currentTarget.value !== 'custom') code = e.currentTarget.value;
      }}
    >
      {#each presets as [name, pattern]}<option value={pattern}>{name}</option>{/each}
      <option value="custom">Custom</option>
    </select>
  </label>
  <label class="field"
    ><span>Format code</span><input
      aria-label="Format code"
      bind:value={code}
      maxlength="2048"
      spellcheck="false"
    /></label
  >
  <p class="muted">
    Use 0 for required digits, # for optional digits and ? for fraction spacing. Put labels in
    quotation marks.
  </p>
  <div class="format-preview">
    <span>Preview</span><output aria-label="Format preview"
      >{preview.text || (preview.error ? '' : '(empty)')}</output
    >
  </div>
  {#if preview.error}<p role="alert">{preview.error}</p>{/if}
  <div class="modal-actions">
    <button onclick={onclose}>Cancel</button><button
      class="primary"
      disabled={!!preview.error}
      onclick={() => onapply(code)}>Apply format</button
    >
  </div>
</Modal>

<style>
  .format-preview {
    display: grid;
    gap: 8px;
    border: 1px solid var(--border, #e6e9e2);
    border-radius: 8px;
    padding: 16px;
    margin-top: 16px;
  }
  .format-preview span {
    font-size: 12px;
    color: #6c7e73;
  }
  output {
    white-space: pre-wrap;
    overflow-wrap: anywhere;
    font-variant-numeric: tabular-nums;
    min-height: 24px;
  }
</style>
