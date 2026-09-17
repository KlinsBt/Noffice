<script lang="ts">
  import type { Editor } from '@tiptap/core';
  import { paragraphIndents, paragraphBreaks } from '../paragraph-layout';
  import { wordSelectionParagraphLayout, wordParagraphIndentInput } from '../word-paragraph-layout-commands';
  import { AlignLeft, AlignCenter, AlignRight, AlignJustify } from '@lucide/svelte';
  import Tool from './Tool.svelte';
  let { editor, transaction, mode }: {
    editor: Editor | undefined; transaction: number; mode: 'alignment' | 'indents' | 'breaks';
  } = $props();
  let invalid = $state('');
  const selected = $derived.by(() => { transaction; return editor ? wordSelectionParagraphLayout(editor.state) : null; });
  const errorId = $props.id();
  function indent(input: HTMLInputElement, key: typeof paragraphIndents[number][0], focus = false) {
    const value = wordParagraphIndentInput(input.valueAsNumber);
    if (value === null) { invalid = key; return; }
    if (editor?.commands.setWordParagraphLayout({ [key]: value })) {
      invalid = ''; input.value = String(value); if (focus) editor.view.focus();
    }
  }
</script>

{#if mode === 'alignment'}
  {#each ['left', 'center', 'right', 'justify'] as const as align, index}
    {@const Icon = [AlignLeft, AlignCenter, AlignRight, AlignJustify][index]}
    <Tool label={`Align ${align}`} active={selected?.textAlign === align}
      onclick={() => editor?.chain().focus().setWordParagraphLayout({ textAlign: align }).run()}><Icon /></Tool>
  {/each}
{:else if mode === 'indents'}
  {#each paragraphIndents as [key, , label]}
    <label class="field"><span>{key === 'indentStart' ? 'Before text' : key === 'indentEnd' ? 'After text' : 'First line'} (pt)</span>
      <input type="number" step="any" aria-label={label}
        value={selected?.[key] ?? ''} placeholder={selected?.[key] === null ? 'Mixed' : undefined}
        aria-invalid={invalid === key ? 'true' : undefined} aria-describedby={invalid === key ? errorId : undefined}
        onchange={event => indent(event.currentTarget, key)}
        onkeydown={event => {
          if (event.key === 'Enter' && !event.isComposing && !event.altKey && !event.ctrlKey && !event.metaKey) {
            event.preventDefault(); indent(event.currentTarget, key, true);
          } else if (event.key === 'Escape') {
            event.preventDefault(); invalid = ''; event.currentTarget.value = String(selected?.[key] ?? ''); editor?.view.focus();
          }
        }} />
    </label>
  {/each}
  {#if invalid}<p id={errorId} role="status">Enter an indent that rounds to −1,584 through 1,584 points.</p>{/if}
{:else}
  {#each paragraphBreaks as [key, , , label]}
    <label class="ribbon-check"><input type="checkbox" checked={selected?.[key] ?? false}
      indeterminate={selected?.[key] === null}
      onchange={event => editor?.chain().focus().setWordParagraphLayout({ [key]: event.currentTarget.checked }).run()} />{label}</label>
  {/each}
{/if}
