<script lang="ts">
  import type { Snippet } from 'svelte';
  import { X } from '@lucide/svelte';
  import { focusOnMount } from './css';
  let {
    query = $bindable(''),
    replacement = $bindable(''),
    matchCase = $bindable(false),
    whole = $bindable(false),
    wholeLabel = 'Whole words only',
    findLabel = 'Find text',
    placeholder = 'Find literal text',
    status = '',
    onfind,
    onreplace,
    onclose,
    children,
  }: {
    query?: string;
    replacement?: string;
    matchCase?: boolean;
    whole?: boolean;
    wholeLabel?: string;
    findLabel?: string;
    placeholder?: string;
    status?: string;
    onfind: (previous: boolean) => void;
    onreplace: (all: boolean) => void;
    onclose: () => void;
    children?: Snippet;
  } = $props();
</script>

<!-- svelte-ignore a11y_no_noninteractive_element_interactions (Enter and Escape from the contained controls operate the search panel.) -->
<section
  class="search-panel"
  aria-label="Find and replace"
  onkeydown={(e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      onclose();
    }
    if (e.key === 'Enter' && e.target instanceof HTMLInputElement && e.target.type === 'text') {
      e.preventDefault();
      onfind(e.shiftKey);
    }
  }}
>
  <div class="search-fields">
    <input
      use:focusOnMount
      aria-label={findLabel}
      {placeholder}
      bind:value={query}
      maxlength="256"
    />
    <input
      aria-label="Replacement text"
      placeholder="Replace with"
      bind:value={replacement}
      maxlength="32767"
    />
    <button disabled={!query} onmousedown={(e) => e.preventDefault()} onclick={() => onfind(true)}
      >Find previous</button
    >
    <button disabled={!query} onmousedown={(e) => e.preventDefault()} onclick={() => onfind(false)}
      >Find next</button
    >
    <button
      disabled={!query}
      onmousedown={(e) => e.preventDefault()}
      onclick={() => onreplace(false)}>Replace</button
    >
    <button
      disabled={!query}
      onmousedown={(e) => e.preventDefault()}
      onclick={() => onreplace(true)}>Replace all</button
    >
    <button class="search-close" aria-label="Close find and replace" onclick={onclose}
      ><X size={16} /></button
    >
  </div>
  <div class="search-options">
    <label><input type="checkbox" bind:checked={matchCase} />Match case</label>
    <label><input type="checkbox" bind:checked={whole} />{wholeLabel}</label>
    {#if children}{@render children()}{/if}
    <span role="status">{status}</span>
  </div>
</section>

<style>
  .search-panel {
    background: #f8fafc;
    border-bottom: 1px solid #dce2ea;
    padding: 10px 20px;
    flex-shrink: 0;
    max-height: 35vh;
    overflow: auto;
  }
  .search-fields,
  .search-options {
    display: flex;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
  }
  .search-fields > input {
    width: 190px;
    min-width: 100px;
    max-width: 100%;
    padding: 7px 9px;
    font-size: 12px;
  }
  .search-fields button {
    padding: 7px 9px;
    font-size: 12px;
    border: 1px solid #d1d9e4;
    border-radius: 4px;
    background: white;
  }
  .search-options {
    margin-top: 8px;
    gap: 12px;
    font-size: 12px;
  }
  .search-options label {
    display: flex;
    align-items: center;
    gap: 6px;
  }
  .search-options input {
    margin: 0;
  }
  .search-options span {
    color: #526071;
  }
  .search-close {
    margin-left: auto;
  }
</style>
