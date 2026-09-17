<script lang="ts">
  import type { Editor } from '@tiptap/core';
  import type { WordContent } from '../model';
  import { resolveWordSections, singleSectionPage } from '../word-section-layout';
  import { wordSectionMap } from '../word-editor-sections';
  import { selectedWordSections } from '../word-section-ranges';
  import { changeWordPageLayout } from '../word-page-layout';
  import {
    changeWordSectionLayout,
    type WordSectionLayoutChange,
  } from '../word-section-scoped-layout';

  let {
    editor,
    content,
    transaction,
    notify,
  }: {
    editor: Editor | undefined;
    content: WordContent;
    transaction: number;
    notify: (message: string) => void;
  } = $props();
  let scope = $state('document');
  const sections = $derived(resolveWordSections(content));
  const selected = $derived.by(() => {
    transaction;
    if (!editor) return null;
    const map = wordSectionMap(editor.state);
    if (!map?.ranges.length) return sections.length === 1 ? sections[0] : null;
    const { from, to } = editor.state.selection;
    const ids = selectedWordSections(map, from, to);
    return ids.length === 1 ? sections.find((s) => s.id === ids[0]) || null : null;
  });
  const page = $derived(scope === 'section' ? selected : singleSectionPage(content));
  const paper = $derived.by(() => {
    if (!page) return content.paper;
    const size = [page.width, page.height].sort((a, b) => a! - b!).join(',');
    return size === '11906,16838' || size === '11907,16839'
      ? 'a4'
      : size === '12240,15840'
        ? 'letter'
        : 'source';
  });
  const margin = $derived.by(() => {
    if (!page) return content.margin;
    const values = ['left', 'right', 'top', 'bottom'].map((side) => page.margins[side as 'left']);
    if (!values.every((value) => value === values[0])) return 'source';
    return values[0] === 720
      ? 'narrow'
      : values[0] === 1440
        ? 'normal'
        : values[0] === 2160
          ? 'wide'
          : 'source';
  });
  const disabled = $derived(!editor || (scope === 'section' && !selected));
  function apply(change: WordSectionLayoutChange) {
    if (!editor) return;
    try {
      if (scope === 'section') changeWordSectionLayout(editor, content, change);
      else if (change.key !== 'start')
        changeWordPageLayout(editor, change.key, change.value, content);
    } catch (error) {
      notify((error as Error).message);
    }
  }
</script>

<label class="field"
  ><span>Apply to</span><select aria-label="Apply page setup to" bind:value={scope}>
    <option value="document">Whole document</option>
    <option value="section">Current section</option>
  </select></label
>
<label class="field"
  ><span>Paper size</span><select
    aria-label="Paper size"
    value={paper}
    {disabled}
    onchange={(e) => apply({ key: 'paper', value: e.currentTarget.value as WordContent['paper'] })}
  >
    {#if paper === 'source'}<option value="source" disabled>Custom (from document)</option>{/if}
    <option value="a4">A4</option><option value="letter">US Letter</option>
  </select></label
>
<label class="field"
  ><span>Margins</span><select
    aria-label="Margins"
    value={margin}
    {disabled}
    onchange={(e) =>
      apply({ key: 'margin', value: e.currentTarget.value as WordContent['margin'] })}
  >
    {#if margin === 'source'}<option value="source" disabled>Custom (from document)</option>{/if}
    <option value="normal">Normal · 1 inch</option><option value="narrow">Narrow · ½ inch</option>
    <option value="wide">Wide · 1½ inches</option>
  </select></label
>
<label class="field"
  ><span>Orientation</span><select
    aria-label="Page orientation"
    value={page?.orientation || content.orientation || 'portrait'}
    {disabled}
    onchange={(e) =>
      apply({ key: 'orientation', value: e.currentTarget.value as 'portrait' | 'landscape' })}
  >
    <option value="portrait">Portrait</option><option value="landscape">Landscape</option>
  </select></label
>
{#if scope === 'section'}
  <label class="field"
    ><span>Section start</span><select
      aria-label="Section start"
      value={selected?.start || 'nextPage'}
      {disabled}
      onchange={(e) =>
        apply({
          key: 'start',
          value: e.currentTarget.value as 'nextPage' | 'continuous' | 'evenPage' | 'oddPage',
        })}
    >
      <option value="nextPage">Next Page</option><option value="continuous">Continuous</option>
      <option value="evenPage">Even Page</option><option value="oddPage">Odd Page</option>
    </select></label
  >
  {#if !selected}<span role="status">Select text within one section.</span>{/if}
{/if}
