<script lang="ts">
  import { tick } from 'svelte';
  import type { Editor } from '@tiptap/core';
  import Modal from './Modal.svelte';
  import { wordLineSpacing, wordLineSpacingInput, type WordLineRule } from '../word-line-spacing';
  import { wordSelectionLineSpacing } from '../word-line-spacing-commands';
  let { editor, transaction, inline = false }: { editor: Editor | undefined; transaction: number; inline?: boolean } = $props();
  let open = $state(false), rule = $state<WordLineRule>('auto'), amount = $state<number | undefined>(1);
  const selection = $derived.by(() => { transaction; return editor ? wordSelectionLineSpacing(editor.state) : null; });
  const value = $derived(selection ? selection.rule === 'auto' ? String(selection.line / 240) : `${selection.line / 20}pt` : '');
  const valid = $derived(wordLineSpacingInput(amount, rule));
  const presets = ['1', '1.15', '1.5', '1.65', '2', '3'];
  function showOptions() {
    rule = selection?.rule || 'auto';
    amount = selection ? selection.line / (selection.rule === 'auto' ? 240 : 20) : 1;
    open = true;
  }
  async function close() { open = false; await tick(); editor?.view.focus(); }
  async function apply() {
    if (!valid || !editor?.commands.setWordLineSpacing(valid)) return;
    await close();
  }
</script>

<select aria-label="Line spacing" {value} onchange={event => {
  if (event.currentTarget.value === 'custom') { event.currentTarget.value = value; showOptions(); }
  else { const spacing = wordLineSpacing(event.currentTarget.value); if (spacing) editor?.commands.setWordLineSpacing(spacing); }
}}>
  {#if !selection}<option value="" disabled>Mixed spacing</option>
  {:else if selection.rule !== 'auto'}<option {value}>{selection.rule === 'atLeast' ? 'At least' : 'Exactly'} {value}</option>
  {:else if !presets.includes(value)}<option {value}>{value} lines</option>{/if}
  <option value="1">Single</option><option value="1.15">1.15 lines</option><option value="1.5">1.5 lines</option>
  <option value="1.65">Default spacing</option><option value="2">Double</option><option value="3">Triple</option>
  <option value="custom">Line spacing options...</option>
</select>

{#snippet options()}
  <form onsubmit={event => { event.preventDefault(); void apply(); }}>
    <label class="field"><span>Spacing rule</span><select aria-label="Spacing rule" value={rule}
      onchange={event => {
        const next = event.currentTarget.value as WordLineRule;
        if ((next === 'auto') !== (rule === 'auto')) amount = next === 'auto' ? 1 : 12;
        rule = next;
      }}>
      <option value="auto">Multiple</option><option value="exact">Exactly</option><option value="atLeast">At least</option>
    </select></label>
    <label class="field"><span>{rule === 'auto' ? 'Lines' : 'Points'}</span><input aria-label="Spacing amount" type="number"
      step="any" bind:value={amount} aria-invalid={!valid} aria-describedby={!valid ? 'line-spacing-error' : undefined} /></label>
    {#if !valid}<p id="line-spacing-error" role="status">Enter a valid spacing amount. Word permits values that round to 0.7–1584 points{rule === 'auto' ? ' (about 0.06–132 lines)' : ''}.</p>{/if}
    <div class="modal-actions"><button type="button" onclick={close}>{inline ? 'Cancel line spacing' : 'Cancel'}</button>
      <button type="submit" class="primary" disabled={!valid}>{inline ? 'Apply line spacing' : 'Apply'}</button></div>
  </form>
{/snippet}
{#if open}
  {#if inline}<div class="line-options" role="region" aria-label="Line spacing options">{@render options()}</div>
  {:else}<Modal title="Line spacing" onclose={close}>{@render options()}</Modal>{/if}
{/if}

<style>
  .line-options { border: 1px solid var(--border); border-radius: .4rem; padding: .75rem; margin: .5rem 0; }
  .field { display: flex; justify-content: space-between; align-items: center; gap: .75rem; margin: .5rem 0; }
  .field input { width: 6rem; }
</style>
