<script lang="ts">
  import { onMount } from 'svelte';
  import { Editor } from '@tiptap/core';
  import { wordExtensions } from '../word-extensions';
  import { sanitizeHTML } from '../formats';
  import { syncWordSelection, syncWordNavigation } from '../word-links';
  import Modal from './Modal.svelte';
  import Tool from './Tool.svelte';
  import FontPicker from './FontPicker.svelte';
  import WordParagraphSpacing from './WordParagraphSpacing.svelte';
  import WordLineSpacing from './WordLineSpacing.svelte';
  import WordParagraphLayout from './WordParagraphLayout.svelte';
  import { defaultFont } from '../fonts';
  import { wordDefaults } from '../word-defaults';
  import { wordSelectionFont } from '../word-selection-font';
  import { wordSelectionScript } from '../word-script';
  import { WordFontLineMetrics } from '../word-font-line-metrics';
  import { WordJustification } from '../word-justification';
  import type { WordCompatibility } from '../word-compatibility';
  import { Bold, Italic, Underline, Subscript, Superscript, Undo2, Redo2, AArrowUp, AArrowDown } from '@lucide/svelte';
  let {
    title,
    html,
    pageLeft = null,
    compatibility,
    onapply,
    onclose,
  }: {
    title: string;
    html: string;
    pageLeft?: number | null;
    compatibility?: WordCompatibility;
    onapply: (html: string) => void;
    onclose: () => void;
  } = $props();
  let editor = $state.raw<Editor>(),
    transaction = $state(0),
    error = $state('');
  let host: HTMLDivElement;
  onMount(() => {
    const instance = new Editor({
      element: host,
      extensions: [WordJustification.configure({ compatibility: () => compatibility }), WordFontLineMetrics, ...wordExtensions({ retainedEditRuns: true, tabPageLeft: () => pageLeft })],
      content: html,
      parseOptions: { preserveWhitespace: true },
      editorProps: {
        handleDOMEvents: {
          keyup: (_view, event) => {
            syncWordNavigation(instance, event);
            return false;
          },
          mouseup: () => {
            syncWordSelection(instance);
            return false;
          },
        },
        transformPastedHTML: sanitizeHTML,
        attributes: {
          class: 'document-page',
          style: 'padding:0;min-height:6rem;width:100%;height:auto',
          role: 'textbox',
          'aria-label': 'Header or footer text',
          'aria-multiline': 'true',
        },
      },
      onTransaction: ({ transaction: tr }) => {
        if (tr.docChanged || tr.selectionSet || tr.storedMarksSet) transaction++;
      },
    });
    editor = instance;
    return () => instance.destroy();
  });
  const available = (command: 'undo' | 'redo') => {
    transaction;
    return editor?.can()[command]() || false;
  };
  const active = (mark: string) => {
    transaction;
    if (mark === 'superscript' || mark === 'subscript')
      return editor ? wordSelectionScript(editor.state) === mark : false;
    return editor?.isActive(mark) || false;
  };
  const selectionFont = $derived.by(() => {
    transaction;
    return editor ? wordSelectionFont(editor.state) : null;
  });
  const fontPoints = $derived(selectionFont ? selectionFont.size ?? '' : wordDefaults.fontSize);
  function applyFontSize(input: HTMLInputElement) {
    const size = input.valueAsNumber;
    if (input.validity.valid && Number.isFinite(size)) editor?.chain().focus().setWordFontSize(size + 'pt').run();
    else input.value = String(fontPoints);
  }
  function apply() {
    try {
      if (editor) onapply(editor.getHTML());
    } catch (e) {
      error = (e as Error).message;
    }
  }
</script>

<Modal {title} {onclose} wide>
  <div
    class="story-toolbar"
    role="toolbar"
    tabindex={-1}
    aria-label="Header or footer formatting"
    onmousedown={(e) => {
      if (editor) syncWordSelection(editor);
      if ((e.target as HTMLElement).closest('button')) e.preventDefault();
    }}
  >
    <FontPicker
      value={selectionFont?.family || defaultFont}
      mixed={selectionFont?.family === null}
      onchange={(family) => editor?.chain().focus().setWordFontFamily(JSON.stringify(family)).run()}
    />
    <input
      class="story-font-size"
      aria-label="Font size"
      title="Font size in points"
      type="number"
      min="1"
      max="1638"
      step="0.5"
      value={fontPoints}
      placeholder={selectionFont?.size === null ? 'Mixed' : undefined}
      onchange={(event) => applyFontSize(event.currentTarget)}
      onkeydown={(event) => {
        if (event.key === 'Enter' && !event.isComposing && !event.altKey && !event.ctrlKey && !event.metaKey) {
          event.preventDefault();
          applyFontSize(event.currentTarget);
        }
      }}
    />
    <Tool label="Grow font" onclick={() => editor?.chain().focus()
      .stepWordFontSize(true).run()}><AArrowUp /></Tool>
    <Tool label="Shrink font" onclick={() => editor?.chain().focus()
      .stepWordFontSize(false).run()}><AArrowDown /></Tool>
    <Tool
      label="Bold"
      active={active('bold')}
      onclick={() => editor?.chain().focus().toggleBold().run()}><Bold /></Tool
    >
    <Tool
      label="Italic"
      active={active('italic')}
      onclick={() => editor?.chain().focus().toggleItalic().run()}><Italic /></Tool
    >
    <Tool
      label="Underline"
      active={active('underline')}
      onclick={() => editor?.chain().focus().toggleUnderline().run()}><Underline /></Tool
    >
    <Tool
      label="Subscript"
      active={active('subscript')}
      onclick={() => editor?.chain().focus().toggleWordScript('subscript').run()}><Subscript /></Tool
    >
    <Tool
      label="Superscript"
      active={active('superscript')}
      onclick={() => editor?.chain().focus().toggleWordScript('superscript').run()}><Superscript /></Tool
    >
    <Tool
      label="Undo"
      disabled={!available('undo')}
      onclick={() => editor?.chain().focus().undo().run()}><Undo2 /></Tool
    >
    <Tool
      label="Redo"
      disabled={!available('redo')}
      onclick={() => editor?.chain().focus().redo().run()}><Redo2 /></Tool
    >
    <WordParagraphLayout {editor} {transaction} mode="alignment" />
  </div>
  <WordParagraphSpacing {editor} {transaction} />
  <WordLineSpacing {editor} {transaction} inline />
  <div class="story-paragraph-indents"><WordParagraphLayout {editor} {transaction} mode="indents" /></div>
  <div class="story-paragraph-breaks"><WordParagraphLayout {editor} {transaction} mode="breaks" /></div>
  <div class="story-editor" bind:this={host}></div>
  {#if error}<p role="alert">{error}</p>{/if}
  <div class="story-actions">
    <button onclick={onclose}>Cancel</button>
    <button class="primary" onclick={apply}>Apply header or footer</button>
  </div>
</Modal>

<style>
  .story-toolbar,
  .story-actions {
    display: flex;
    gap: 0.5rem;
    margin: 0.75rem 0;
  }
  .story-toolbar {
    align-items: center;
    flex-wrap: wrap;
  }
  .story-paragraph-indents,
  .story-paragraph-breaks {
    display: flex;
    flex-wrap: wrap;
    gap: 0.5rem;
    margin: 0.75rem 0;
  }
  .story-font-size {
    width: 5.5rem;
  }
  .story-actions {
    justify-content: flex-end;
  }
  .story-editor {
    border: 1px solid var(--border);
    background: white;
    color: black;
    max-height: 45vh;
    overflow: auto;
    padding: 1rem;
  }
  .story-editor :global(.tiptap) {
    min-height: 6rem;
    outline: none;
  }
  .story-editor :global(p) {
    color: inherit;
  }
</style>
