<script lang="ts">
  import { untrack } from 'svelte';
  import Modal from './Modal.svelte';
  import { focusOnMount } from './css';
  import type { WordContent } from '../model';
  import type { WordStoryReference } from '../word-stories';
  import { wordStoryIsLinked } from '../word-story-links';
  import { resolveWordSections } from '../word-section-layout';

  let {
    content,
    onclose,
    onapply,
  }: {
    content: WordContent;
    onclose: () => void;
    onapply: (
      sectionId: string,
      kind: WordStoryReference['kind'],
      slot: WordStoryReference['slot'],
      linked: boolean,
    ) => void;
  } = $props();
  const sections = $derived(resolveWordSections(content));
  const initial = untrack(() => sections[1] || sections[0]);
  let sectionId = $state(initial?.id || '');
  let kind = $state<WordStoryReference['kind']>('header');
  let slot = $state<WordStoryReference['slot']>('default');
  let linked = $state(untrack(() => wordStoryIsLinked(content, sectionId, kind, slot)));
  let error = $state('');
  const first = $derived(sections.findIndex((s) => s.id === sectionId) < 1);
  function select(id: string, nextKind = kind, nextSlot = slot) {
    sectionId = id;
    kind = nextKind;
    slot = nextSlot;
    linked = wordStoryIsLinked(content, sectionId, kind, slot);
    error = '';
  }
</script>

<Modal title="Header and footer links" {onclose}>
  <form
    onsubmit={(event) => {
      event.preventDefault();
      if (first) return;
      try {
        onapply(sectionId, kind, slot, linked);
      } catch (cause) {
        error = cause instanceof Error ? cause.message : 'The link could not be changed.';
      }
    }}
  >
    <label class="field"
      ><span>Section</span>
      <select
        aria-label="Link section"
        value={sectionId}
        onchange={(event) => select(event.currentTarget.value)}
        use:focusOnMount
      >
        {#each sections as section, index}<option value={section.id}>Section {index + 1}</option
          >{/each}
      </select>
    </label>
    <label class="field"
      ><span>Header or footer</span>
      <select
        aria-label="Header or footer"
        value={kind}
        onchange={(event) => select(sectionId, event.currentTarget.value as typeof kind)}
      >
        <option value="header">Header</option><option value="footer">Footer</option>
      </select>
    </label>
    <label class="field"
      ><span>Page type</span>
      <select
        aria-label="Header/footer page type"
        value={slot}
        onchange={(event) => select(sectionId, kind, event.currentTarget.value as typeof slot)}
      >
        <option value="default">Default</option><option value="first">First page</option><option
          value="even">Even pages</option
        >
      </select>
    </label>
    <label class="field"
      ><span><input type="checkbox" bind:checked={linked} disabled={first} /> Link to previous</span
      ></label
    >
    {#if first}<p>The first section has no previous section.</p>
    {:else}<p>
        Linking uses the previous section's corresponding header or footer. Unlinking keeps a
        separate copy that you can edit independently.
      </p>{/if}
    <p>First-page and even-page content is displayed when enabled in Page options.</p>
    {#if error}<p role="alert">{error}</p>{/if}
    <div class="modal-actions">
      <button type="button" onclick={onclose}>Cancel</button><button
        type="submit"
        class="primary"
        disabled={first}>Apply header/footer link</button
      >
    </div>
  </form>
</Modal>
