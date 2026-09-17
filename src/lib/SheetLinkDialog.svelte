<script lang="ts">
  import Modal from './Modal.svelte';
  import { untrack } from 'svelte';
  import type { Cell } from '../model';
  let {
    cell,
    onclose,
    onapply,
  }: {
    cell: Cell;
    onclose: () => void;
    onapply: (href: string | null, text: string, tooltip: string) => void;
  } = $props();
  let href = $state(untrack(() => cell.hyperlink || '')),
    text = $state(untrack(() => cell.value)),
    tooltip = $state(untrack(() => cell.hyperlinkTooltip || '')),
    error = $state('');
  function apply(remove = false) {
    try {
      onapply(remove ? null : href, text, tooltip);
    } catch (e) {
      error = (e as Error).message;
    }
  }
</script>

<Modal title="Cell hyperlink" {onclose}>
  <label class="field"
    ><span>Text to display</span><input
      aria-label="Link text"
      bind:value={text}
      maxlength="32767"
      disabled={cell.value.startsWith('=') && cell.dataType !== 'text'}
    /></label
  >
  <label class="field"
    ><span>Address</span><input
      aria-label="Link address"
      bind:value={href}
      maxlength="8192"
      placeholder="https://example.com or #'Sheet 2'!A1"
    /></label
  >
  <label class="field"
    ><span>ScreenTip</span><input
      aria-label="Link ScreenTip"
      bind:value={tooltip}
      maxlength="255"
    /></label
  >
  <p class="muted">
    Web, email and workbook locations are supported. Formula results keep their existing display
    text.
  </p>
  {#if error}<p role="alert">{error}</p>{/if}
  <div class="modal-actions">
    {#if cell.hyperlink}<button onclick={() => apply(true)}>Remove link</button>{/if}
    <button onclick={onclose}>Cancel</button><button class="primary" onclick={() => apply()}
      >Apply link</button
    >
  </div>
</Modal>
