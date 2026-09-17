<script lang="ts">
  import { onMount } from 'svelte';
  import Modal from './Modal.svelte';
  import { bundledFonts, systemFonts, defaultFont, initializeFonts, importFont } from '../fonts';
  let { value = defaultFont, mixed = false, onchange }: { value?: string; mixed?: boolean; onchange: (family: string) => void } =
    $props();
  let custom = $state<string[]>([]);
  let open = $state(false),
    family = $state(''),
    error = $state(''),
    busy = $state(false);
  let file = $state<File>();
  onMount(() => {
    void initializeFonts()
      .then((names) => (custom = names))
      .catch((e) => (error = e.message));
  });
  async function add() {
    if (!file) {
      error = 'Choose a font file first.';
      return;
    }
    busy = true;
    error = '';
    try {
      custom = await importFont(file, family);
      onchange(family.trim());
      open = false;
    } catch (e) {
      error = (e as Error).message;
    } finally {
      busy = false;
    }
  }
</script>

<div class="font-picker">
  <select aria-label="Font family" value={mixed ? '' : value} onchange={(e) => onchange(e.currentTarget.value)}>
    {#if mixed}<option value="" disabled>Mixed fonts</option>{/if}
    <optgroup label="Included · works offline"
      >{#each bundledFonts as name}<option>{name}</option>{/each}</optgroup
    >
    {#if custom.length}<optgroup label="Imported on this device"
        >{#each custom as name}<option>{name}</option>{/each}</optgroup
      >{/if}
    <optgroup label="If installed on this device"
      >{#each systemFonts as name}<option>{name}</option>{/each}</optgroup
    >
    {#if !mixed && ![...bundledFonts, ...systemFonts, ...custom].includes(value)}<option {value}
        >{value} (document font)</option
      >{/if}
  </select>
  <button
    class="font-add"
    title="Add a font from this device"
    aria-label="Add font"
    onclick={() => {
      open = true;
    }}>+</button
  >
</div>
{#if open}<Modal title="Add a font" onclose={() => (open = false)}>
    <p>
      Load a font file from your device. It will be available in all three modes, including after a
      reload.
    </p>
    <label class="field"
      ><span>Family name</span><input
        aria-label="Imported font family name"
        bind:value={family}
        placeholder="My font"
        maxlength={80}
      /></label
    >
    <label class="field"
      ><span>Font file</span><input
        aria-label="Font file"
        type="file"
        accept=".ttf,.otf,.woff,.woff2"
        onchange={(e) => (file = e.currentTarget.files?.[0])}
      /></label
    >
    <p class="muted">
      TTF, OTF, WOFF or WOFF2, up to 10 MB. Fonts are stored separately from documents. Exports
      preserve the family name; font files are not embedded. Recipients need the same font
      installed.
    </p>
    {#if error}<p role="alert">{error}</p>{/if}
    <button class="primary" disabled={busy} onclick={add}
      >{busy ? 'Loading font…' : 'Add font to this device'}</button
    >
  </Modal>{/if}
