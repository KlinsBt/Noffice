<script lang="ts">
  import type { Snippet } from 'svelte';
  import type { HTMLButtonAttributes } from 'svelte/elements';
  let {
    label,
    active,
    children,
    ...rest
  }: HTMLButtonAttributes & { label: string; active?: boolean; children?: Snippet } = $props();
</script>

<button
  type="button"
  class:active
  class="tool"
  title={label}
  aria-label={label}
  aria-pressed={active}
  {...rest}>{@render children?.()}</button
>
