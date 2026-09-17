<script lang="ts">
  import { untrack } from 'svelte';
  import Modal from './Modal.svelte';
  import { focusOnMount } from './css';
  import type { WordContent } from '../model';
  import type { WordStoryPageOptions } from '../word-story-options';
  import { wordStoryDistancePoints as distance } from '../word-story-options';
  import { resolveWordSections } from '../word-section-layout';

  let {
    content,
    onclose,
    onapply,
  }: {
    content: WordContent;
    onclose: () => void;
    onapply: (sectionId: string, options: WordStoryPageOptions) => void;
  } = $props();
  const sections = $derived(resolveWordSections(content));
  const initial = untrack(() => sections[0]);
  let sectionId = $state(initial?.id || '');
  let first = $state(initial?.differentFirstPage || false);
  let even = $state(untrack(() => !!content.stories?.evenAndOddHeaders));
  let header = $state(String((initial?.margins.header ?? 0) / 20));
  let footer = $state(String((initial?.margins.footer ?? 0) / 20));
  let error = $state('');
  const valid = $derived(
    !!sectionId && Number.isFinite(distance(header)) && Number.isFinite(distance(footer)),
  );
  function selectSection(id: string) {
    const section = sections.find((s) => s.id === id);
    if (!section) return;
    sectionId = id;
    first = section.differentFirstPage;
    header = String((section.margins.header ?? 0) / 20);
    footer = String((section.margins.footer ?? 0) / 20);
    error = '';
  }
</script>

<Modal title="Header and footer page options" {onclose}>
  <form
    onsubmit={(event) => {
      event.preventDefault();
      if (!valid) return;
      try {
        onapply(sectionId, {
          differentFirstPage: first,
          evenAndOddHeaders: even,
          headerDistance: distance(header),
          footerDistance: distance(footer),
        });
      } catch (cause) {
        error = cause instanceof Error ? cause.message : 'The page options could not be applied.';
      }
    }}
  >
    <label class="field"
      ><span>Section</span>
      <select
        aria-label="Header/footer section"
        value={sectionId}
        onchange={(event) => selectSection(event.currentTarget.value)}
        use:focusOnMount
      >
        {#each sections as section, index}<option value={section.id}>Section {index + 1}</option
          >{/each}
      </select>
    </label>
    <label class="field"
      ><span><input type="checkbox" bind:checked={first} /> Different first page</span></label
    >
    <label class="field"
      ><span><input type="checkbox" bind:checked={even} /> Different odd and even pages</span
      ></label
    >
    <p>
      Odd/even headers apply to the whole document. Other options apply to the selected section.
    </p>
    <label class="field"
      ><span>Header from top (pt)</span><input
        aria-label="Header from top (pt)"
        type="text"
        inputmode="decimal"
        bind:value={header}
      /></label
    >
    <label class="field"
      ><span>Footer from bottom (pt)</span><input
        aria-label="Footer from bottom (pt)"
        type="text"
        inputmode="decimal"
        bind:value={footer}
      /></label
    >
    <p>Distances are rounded to the nearest 0.05 pt.</p>
    {#if !valid}<p role="alert">Enter distances that round to between 0 and 1,584 pt.</p>{/if}
    {#if error}<p role="alert">{error}</p>{/if}
    <div class="modal-actions">
      <button type="button" onclick={onclose}>Cancel</button><button
        class="primary"
        type="submit"
        disabled={!valid}>Apply page options</button
      >
    </div>
  </form>
</Modal>
