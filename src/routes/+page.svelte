<script lang="ts">
  import Workspace from '../lib/Workspace.svelte';
</script>

<svelte:head><title>Noffice — A little space for big ideas</title></svelte:head>
<svelte:boundary>
  <Workspace />
  {#snippet failed(error, reset)}
    <div class="fatal">
      <h1>Something interrupted your workspace.</h1>
      <p>Previously saved files remain on this device. Reload to open them again.</p>
      <button onclick={() => location.reload()}>Reload workspace</button><button onclick={reset}
        >Try again</button
      >
    </div>
  {/snippet}
</svelte:boundary>
