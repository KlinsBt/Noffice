<script lang="ts">
  import type { Editor } from '@tiptap/core';
  import { wordSelectionParagraphSpacing } from '../word-paragraph-spacing-commands';
  import { wordParagraphSpaceInput, wordParagraphSpaceTwips, type WordParagraphSpace, type WordSpaceSide } from '../word-paragraph-spacing';
  let { editor, transaction }: { editor: Editor | undefined; transaction: number } = $props();
  let error = $state('');
  const selection = $derived.by(() => { transaction; return editor ? wordSelectionParagraphSpacing(editor.state) : null; });
  const amount = (space?: WordParagraphSpace | null) => !space || space.unit === 'auto' ? '' : space.value / (space.unit === 'points' ? 20 : 100);
  function change(side: WordSpaceSide, unit: WordParagraphSpace['unit'], value: unknown) {
    const space = wordParagraphSpaceInput(unit, value);
    if (!space) { error = unit === 'lines' ? 'Enter a spacing from 0 to 327.67 lines.' : 'Enter a spacing from 0 to 1584 points.'; return false; }
    const applied = editor?.commands.setWordParagraphSpacing({ [side]: space }) || false;
    if (applied) error = '';
    return applied;
  }
</script>

<div class="paragraph-spacing">
  {#each ['before', 'after'] as side}
    {@const key = side as WordSpaceSide}
    {@const space = selection?.[key]}
    {@const label = side === 'before' ? 'Before' : 'After'}
    <label class="field"><span>{label}</span>
      <input aria-label={`${label} paragraph`} type="number" min="0" max={space?.unit === 'lines' ? 327.67 : 1584} step="any"
        value={amount(space)} placeholder={space ? undefined : 'Mixed'} disabled={space?.unit === 'auto'}
        onchange={event => change(key, space?.unit || 'points', event.currentTarget.value)}
        onkeydown={event => { if (event.key === 'Enter' && !event.isComposing) { event.preventDefault(); if (change(key, space?.unit || 'points', event.currentTarget.value)) editor?.view.focus(); } }} />
    </label>
    <select aria-label={`${label} paragraph unit`} value={space?.unit || ''}
      onchange={event => {
        const unit = event.currentTarget.value as WordParagraphSpace['unit'];
        const points = space ? wordParagraphSpaceTwips(space) / 20 : 0;
        change(key, unit, unit === 'lines' ? points / 12 : points);
      }}>
      {#if !space}<option value="" disabled>Mixed</option>{/if}
      <option value="points">Points</option><option value="lines">Lines</option><option value="auto">Auto</option>
    </select>
  {/each}
  <label class="contextual"><input type="checkbox" aria-label="Don't add space between paragraphs of the same style"
    checked={selection?.contextual === true} indeterminate={selection?.contextual === null}
    onchange={event => { editor?.commands.setWordParagraphSpacing({ contextual: event.currentTarget.checked }); error = ''; }} />
    Don't add space between paragraphs of the same style</label>
  {#if error}<span role="alert">{error}</span>{/if}
</div>

<style>
  .paragraph-spacing { display: grid; grid-template-columns: 1fr auto; gap: .25rem .5rem; align-items: center; }
  .field { display: flex; gap: .5rem; align-items: center; justify-content: space-between; }
  .field input { width: 5.5rem; }
  .contextual, [role='alert'] { grid-column: 1 / -1; max-width: 21rem; font-size: .75rem; }
  [role='alert'] { color: var(--danger, #b91c1c); }
</style>
