<script lang="ts">
  import type { Snippet } from 'svelte';
  let {
    label,
    kind = '',
    children,
  }: { label: string; kind?: string; children: Snippet } = $props();
</script>

<div class={`ribbon-group ${kind}`} role="group" aria-label={label}>
  <div class="ribbon-group-content">{@render children()}</div>
  <span class="ribbon-group-caption" aria-hidden="true">{label}</span>
</div>
