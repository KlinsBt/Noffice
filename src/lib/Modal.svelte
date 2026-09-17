<script lang="ts">
  import { onMount, type Snippet } from 'svelte';
  import { X } from '@lucide/svelte';
  import Tool from './Tool.svelte';
  let {
    title,
    children,
    onclose,
    wide = false,
  }: { title: string; children: Snippet; onclose: () => void; wide?: boolean } = $props();
  let dialog: HTMLDialogElement;
  onMount(() => {
    dialog.showModal();
  });
</script>

<dialog
  bind:this={dialog}
  class:wide
  class="modal"
  aria-label={title}
  oncancel={onclose}
  onclick={(e) => {
    if (e.target === dialog) onclose();
  }}
>
  <div class="modal-heading">
    <h2>{title}</h2>
    <Tool label="Close dialog" onclick={onclose}><X size={19} /></Tool>
  </div>
  {@render children()}
</dialog>
