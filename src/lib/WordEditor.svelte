<script lang="ts">
  import SearchBar from './SearchBar.svelte';
  import WordStoryEditor from './WordStoryEditor.svelte';
  import WordStoryOptions from './WordStoryOptions.svelte';
  import WordStoryLinks from './WordStoryLinks.svelte';
  import { wordStoryGridHtml } from '../word-story-grid';
  import { changeWordStoryLink } from '../word-story-links';
  import { createWordStory } from '../word-story-create';
  import { changeWordStoryPageOptions } from '../word-story-options';
  import { authoredWordStoryTemplates } from '../word-authored-stories';
  import type { WordStories } from '../word-stories';
  import { resolveWordSections } from '../word-section-layout';
  import {
    WordStoryState,
    initializeWordStories,
    changeWordStory,
    wordStoryChoices,
  } from '../word-story-edit';
  import { wordMatches, replaceWordMatches } from '../word-search';
  import {
    selectedWordLink,
    applyWordLink,
    syncWordSelection,
    syncWordNavigation,
  } from '../word-links';
  let searchCase = $state(false),
    searchWhole = $state(false),
    searchStatus = $state('');
  import RibbonGroup from './RibbonGroup.svelte';
  import { onMount, tick } from 'svelte';
  import { Editor } from '@tiptap/core';
  import { closeHistory } from '@tiptap/pm/history';
  import FontPicker from './FontPicker.svelte';
  import WordParagraphSpacing from './WordParagraphSpacing.svelte';
  import WordPageSetup from './WordPageSetup.svelte';
  import { wordExtensions, wordJSON } from '../word-extensions';
  import WordParagraphLayout from './WordParagraphLayout.svelte';
  import WordLineSpacing from './WordLineSpacing.svelte';
  import { defaultFont } from '../fonts';
  import { wordSelectionFont } from '../word-selection-font';
  import { wordSelectionScript } from '../word-script';
  import {
    changeCase,
    captureTextFormat,
    applyTextFormat,
    indentParagraphs,
    type CaseMode,
    type TextFormat,
  } from '../word-commands';
  import { showFormattingMarks } from '../word-marks';
  import { singleSectionPage } from '../word-section-layout';
  import { authoredWordSection } from '../word-authored-section';
  import { checkpointWordEditSession, initializeWordEditSession } from '../word-edit-run';
  import {
    WordSurfaceView,
    wordColumnExportParagraphs,
    wordPdfSnapshot,
    type SurfaceViewState,
  } from '../word-surface-view';
  import type { WordExportLayout } from '../word-export-layout';
  import { WordFontLineMetrics, wordLayoutChangeEvent } from '../word-font-line-metrics';
  import { WordJustification } from '../word-justification';
  import { prepareWordPrint } from '../word-print-ready';
  import {
    WordEditorSections,
    wordSectionMap,
    updateWordSectionSource,
  } from '../word-editor-sections';
  import { selectedWordSections } from '../word-section-ranges';
  import { insertWordSectionBreak } from '../word-section-insert';
  import type { WordSectionStart } from '../word-section-identity';
  import { WordPageLayout, initializeWordPageLayout } from '../word-page-layout';
  import {
    Bold,
    Subscript,
    Superscript,
    Italic,
    Underline,
    Strikethrough,
    List,
    ListOrdered,
    Undo2,
    Redo2,
    ImagePlus,
    Table2,
    Link,
    Highlighter,
    Search,
    Minus,
    Plus,
    Quote,
    RemoveFormatting,
    Paintbrush,
    Pilcrow,
    IndentIncrease,
    IndentDecrease,
    AArrowUp,
    AArrowDown,
    FilePlus2,
    WholeWord,
  } from '@lucide/svelte';
  import type { WordContent } from '../model';
  import { imageData, sanitizeHTML } from '../formats';
  import Tool from './Tool.svelte';
  import Modal from './Modal.svelte';
  import { css, focusOnMount } from './css';
  let {
    content,
    onChange,
    notify,
  }: {
    content: WordContent;
    onChange: (value: WordContent) => void;
    notify: (text: string) => void;
  } = $props();
  let editor = $state.raw<Editor>();
  let storiesOpen = $state(false);
  let storyOptionsOpen = $state(false);
  let storyLinksOpen = $state(false);
  let storyOptionsExpected = $state.raw<WordStories>();
  let editingStory = $state<
    (ReturnType<typeof wordStoryChoices>[number] & { html: string }) | null
  >(null);
  const storyChoices = $derived(wordStoryChoices(content));
  const tabSections = $derived(resolveWordSections(content));
  const storyTemplates = $derived(
    content.stories?.emptyTemplates || authoredWordStoryTemplates(content),
  );
  async function closeStory() {
    editingStory = null;
    storiesOpen = false;
    storyOptionsOpen = false;
    storyLinksOpen = false;
    await tick();
    editor?.view.focus();
  }
  export async function preparePdf(value: WordContent) {
    if (!editor) throw Error('The document is still opening.');
    await prepareWordPrint(editor.view);
    if (
      JSON.stringify(wordJSON(sanitizeHTML(editor.getHTML())).content) !==
      JSON.stringify(wordJSON(sanitizeHTML(value.html)).content)
    )
      throw Error('Document changed before export. Try exporting again.');
    return wordPdfSnapshot(editor.view, JSON.stringify(value));
  }
  export async function prepareExport(value: WordContent): Promise<WordExportLayout | undefined> {
    if (!editor) throw Error('The document is still opening.');
    await prepareWordPrint(editor.view);
    const paragraphs = wordColumnExportParagraphs(editor.view);
    if (!paragraphs) return undefined;
    // Compare through the same HTML normalization on both sides. Editor marks
    // can use unquoted font names while their serialized CSS adds quotes.
    if (
      JSON.stringify(wordJSON(sanitizeHTML(editor.getHTML())).content) !==
      JSON.stringify(wordJSON(sanitizeHTML(value.html)).content)
    )
      throw Error('Document changed before export. Try exporting again.');
    return { content: JSON.stringify(value), paragraphs };
  }
  export async function preparePrint() {
    if (!editor) throw Error('The document is still opening.');
    await prepareWordPrint(editor.view);
    // Closing the Export menu moves focus to its former trigger. Printing and
    // cancellation must retain the document's model selection for later input.
    editor.view.focus();
  }
  export function checkpointExport() {
    if (editor && !editor.isDestroyed) checkpointWordEditSession(editor.view);
  }
  let transaction = $state(0);
  let surfaces = $state.raw<SurfaceViewState>({ plan: null, printRules: '', reason: '' });
  let sectionLabel = $derived.by(() => {
    transaction;
    if (!editor) return '';
    const mapped = wordSectionMap(editor.state);
    if (!mapped?.ranges.length) return '';
    const selection = editor.state.selection;
    const selected = selectedWordSections(mapped, selection.from, selection.to);
    const indexes = new Map(mapped.ranges.map((range, index) => [range.id, index + 1]));
    if (selected.length === 1)
      return `Section ${indexes.get(selected[0])} of ${mapped.ranges.length}`;
    if (selected.length > 1)
      return `Sections ${indexes.get(selected[0])}–${indexes.get(selected.at(-1)!)} of ${mapped.ranges.length}`;
    return '';
  });
  $effect(() => {
    if (editor) {
      try {
        updateWordSectionSource(editor, content.docxStructure, content.sectionState);
      } catch (error) {
        notify((error as Error).message);
      }
    }
  });
  let sourcePage = $derived.by(() => {
    try {
      return singleSectionPage(content);
    } catch {
      return null;
    }
  });
  let sourcePageStyle = $derived(
    sourcePage
      ? {
          '--source-width': `${sourcePage.width! / 15}px`,
          '--source-height': `${sourcePage.height! / 15}px`,
          '--source-top': `${sourcePage.margins.top! / 15}px`,
          '--source-right': `${sourcePage.margins.right! / 15}px`,
          '--source-bottom': `${sourcePage.margins.bottom! / 15}px`,
          '--source-left': `${sourcePage.margins.left! / 15}px`,
        }
      : {},
  );
  let copiedFormat = $state.raw<TextFormat | null>(null),
    painterArmed = $state(false),
    marksVisible = $state(false),
    rulerVisible = $state(true),
    navigationVisible = $state(false),
    countOpen = $state(false),
    spellcheck = $state(true);
  function paintSelection() {
    if (!editor || !painterArmed || !copiedFormat || editor.state.selection.empty) return;
    const selection = window.getSelection();
    if (
      !editor.view.hasFocus() ||
      !selection ||
      selection.isCollapsed ||
      !editor.view.dom.contains(selection.anchorNode) ||
      !editor.view.dom.contains(selection.focusNode)
    )
      return;
    painterArmed = false;
    applyTextFormat(editor, copiedFormat);
  }
  function toggleMarks() {
    if (editor) {
      marksVisible = !marksVisible;
      showFormattingMarks(editor, marksVisible);
    }
  }
  let tab = $state('Home'),
    zoom = $state(100),
    findOpen = $state(false),
    find = $state(''),
    replace = $state(''),
    linkOpen = $state(false),
    url = $state('');
  let linkText = $state(''),
    originalLinkText = '';
  function openLink() {
    if (!editor) return;
    const selection = selectedWordLink(editor);
    url = selection.href;
    linkText = selection.text;
    originalLinkText = selection.text;
    linkOpen = true;
  }
  async function applyLink(remove = false) {
    try {
      applyWordLink(editor!, remove ? null : url, linkText, originalLinkText);
      linkOpen = false;
      await tick();
      editor!.view.focus();
    } catch (error) {
      notify((error as Error).message);
    }
  }
  const imageInput: { current: HTMLInputElement | null } = { current: null };
  const setTab = (v: string) => (tab = v);
  const setZoom = (v: number) => (zoom = v);
  $effect(() => {
    zoom;
    // CSS zoom changes the glyph rasterization grid without necessarily
    // resizing the content box observed by ResizeObserver.
    editor?.view.dom.dispatchEvent(new Event(wordLayoutChangeEvent));
  });
  const setFindOpen = (v: boolean) => (findOpen = v);
  const setFind = (v: string) => (find = v);
  const setReplace = (v: string) => (replace = v);
  const setLinkOpen = (v: boolean) => (linkOpen = v);
  const setUrl = (v: string) => (url = v);
  onMount(() => {
    const instance: Editor = new Editor({
      extensions: [
        WordJustification.configure({ compatibility: () => content.docxStructure?.compatibility }),
        WordFontLineMetrics.configure({ authoredDefaults: () => !!authoredWordSection(content) }),
        ...wordExtensions({
          retainedEditRuns: true,
          numbering: () => content.numbering,
          tabPageLeft: () => {
            const section = tabSections[0];
            return tabSections.length === 1 &&
              (!section.columns || section.columns.widths.length === 1) &&
              section.margins.left !== null
              ? section.margins.left / 15
              : null;
          },
        }),
        WordPageLayout,
        WordStoryState,
        WordEditorSections.configure({ source: content.docxStructure }),
        WordSurfaceView.configure({
          content: () => content,
          change: (value) => {
            surfaces = value;
          },
        }),
      ],
      parseOptions: { preserveWhitespace: true },
      content: content.html,
      editorProps: {
        handleDOMEvents: {
          mousedown: (view, event) => {
            // Table resize handles can consume the event before the browser focuses
            // the editable surface. Keep keyboard navigation in the document.
            if (event.button === 0 && !view.hasFocus()) view.focus();
            return false;
          },
          mouseup: () => {
            syncWordSelection(instance);
            setTimeout(paintSelection, 0);
            return false;
          },
          keyup: (_view, event) => {
            syncWordNavigation(instance, event);
            setTimeout(paintSelection, 0);
            return false;
          },
        },
        handleKeyDown: (_view, event) => {
          if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') {
            openLink();
            return true;
          }
          if ((event.ctrlKey || event.metaKey) && ['f', 'h'].includes(event.key.toLowerCase())) {
            setFindOpen(true);
            return true;
          }
          if (event.key === 'Escape') {
            painterArmed = false;
            return false;
          }
          if ((event.ctrlKey || event.metaKey) && event.altKey && event.key.toLowerCase() === 'c') {
            copiedFormat = captureTextFormat(instance);
            return true;
          }
          if (
            (event.ctrlKey || event.metaKey) &&
            event.altKey &&
            event.key.toLowerCase() === 'v' &&
            copiedFormat
          )
            return applyTextFormat(instance, copiedFormat);
          if ((event.ctrlKey || event.metaKey) && event.shiftKey && event.code === 'Digit8') {
            toggleMarks();
            return true;
          }
          if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'm')
            return indentParagraphs(instance, event.shiftKey);
          if (event.shiftKey && event.key === 'F3') {
            const { from, to } = instance.state.selection,
              text = instance.state.doc.textBetween(from, to);
            return changeCase(
              instance,
              text === text.toLowerCase()
                ? 'upper'
                : text === text.toUpperCase()
                  ? 'title'
                  : 'lower',
            );
          }
          return false;
        },
        transformPastedHTML: sanitizeHTML,
        attributes: {
          class: 'document-page',
          role: 'textbox',
          'aria-multiline': 'true',
          'aria-label': 'Document text',
          spellcheck: 'true',
        },
      },
      onUpdate: ({ editor: instance }) =>
        onChange({
          ...content,
          initialEditSession: instance.state.doc.attrs.wordInitialEditSession || undefined,
          ...instance.state.doc.attrs.wordPageLayout,
          sectionState: instance.state.doc.attrs.wordSectionState || undefined,
          stories: instance.state.doc.attrs.wordStories || undefined,
          html: instance.getHTML(),
        }),
      onTransaction: ({ transaction: change }) => {
        // Removing a focused editor during a keyed render emits a native blur.
        // Pure focus notifications change no ribbon value and must not write
        // Svelte state from inside that render.
        if (
          (change.getMeta('focus') || change.getMeta('blur')) &&
          !change.docChanged &&
          !change.selectionSet &&
          !change.storedMarksSet
        )
          return;
        transaction += 1;
      },
    });
    initializeWordPageLayout(instance, content);
    initializeWordEditSession(instance, content.initialEditSession);
    initializeWordStories(instance, content);
    editor = instance;
    return () => instance.destroy();
  });
  function attachEditor(node: HTMLElement) {
    if (editor) node.appendChild(editor.view.dom);
  }
  const chain = () => editor!.chain().focus();
  function isActive(name: string | Record<string, unknown>, attrs?: Record<string, unknown>) {
    transaction;
    if (name === 'superscript' || name === 'subscript')
      return editor ? wordSelectionScript(editor.state) === name : false;
    return typeof name === 'string' ? editor!.isActive(name, attrs) : editor!.isActive(name);
  }
  let canUndo = $derived.by(() => {
    transaction;
    return editor?.can().undo() || false;
  });
  let canRedo = $derived.by(() => {
    transaction;
    return editor?.can().redo() || false;
  });
  let canMergeCells = $derived.by(() => {
    transaction;
    return editor?.can().mergeCells() || false;
  });
  let canSplitCell = $derived.by(() => {
    transaction;
    return editor?.can().splitCell() || false;
  });
  let textStyle = $derived.by(() => {
    transaction;
    return editor?.getAttributes('textStyle') || {};
  });
  let paragraphStyle = $derived.by(() => {
    transaction;
    return editor?.getAttributes(editor.isActive('heading') ? 'heading' : 'paragraph') || {};
  });
  function paragraphLayout(attrs: Record<string, string | boolean | null>) {
    chain()
      .command(({ tr }) => {
        closeHistory(tr);
        return true;
      })
      .updateAttributes('paragraph', attrs)
      .updateAttributes('heading', attrs)
      .run();
  }
  $effect(() => {
    const style = document.createElement('style');
    const margin =
      content.margin === 'narrow' ? '12.7mm' : content.margin === 'wide' ? '38.1mm' : '25.4mm';
    const size = sourcePage
      ? `${sourcePage.width! / 20}pt ${sourcePage.height! / 20}pt`
      : `${content.paper === 'a4' ? 'A4' : 'letter'} ${content.orientation || 'portrait'}`;
    const margins = sourcePage
      ? ['top', 'right', 'bottom', 'left']
          .map((side) => `${sourcePage!.margins[side as 'top']! / 20}pt`)
          .join(' ')
      : margin;
    style.textContent = `@media print { @page { size: ${size}; margin: ${margins}; } ${surfaces.printRules} }`;
    document.head.appendChild(style);
    return () => style.remove();
  });
  const selectionFont = $derived.by(() => {
    transaction;
    return editor ? wordSelectionFont(editor.state) : null;
  });
  const fontPoints = $derived(selectionFont ? (selectionFont.size ?? '') : 12);
  function applyFontSize(input: HTMLInputElement) {
    const size = input.valueAsNumber;
    if (input.validity.valid && Number.isFinite(size))
      chain()
        .setWordFontSize(size + 'pt')
        .run();
    else input.value = String(fontPoints);
  }
  let words = $derived.by(() => {
    transaction;
    return editor?.getText().trim().split(/\s+/).filter(Boolean).length || 0;
  });
  let selectedWords = $derived.by(() => {
    transaction;
    if (!editor) return 0;
    const { from, to } = editor.state.selection;
    return editor.state.doc.textBetween(from, to, ' ').trim().split(/\s+/).filter(Boolean).length;
  });
  let countedSelection = $state(0);
  let headings = $derived.by(() => {
    transaction;
    const items: { pos: number; text: string }[] = [];
    editor?.state.doc.descendants((node, pos) => {
      if (node.type.name === 'heading') items.push({ pos, text: node.textContent });
    });
    return items;
  });
  function findNext(previous = false) {
    if (!editor || !find) return;
    try {
      const matches = wordMatches(editor, find, { matchCase: searchCase, wholeWord: searchWhole });
      const next = previous
        ? matches.filter((m) => m.to <= editor!.state.selection.from).at(-1) || matches.at(-1)
        : matches.find((m) => m.from >= editor!.state.selection.to) || matches[0];
      if (next) {
        editor.chain().focus().setTextSelection(next).scrollIntoView().run();
        searchStatus = `${matches.indexOf(next) + 1} of ${matches.length} matches`;
      } else searchStatus = 'No matches found.';
    } catch (error) {
      searchStatus = (error as Error).message;
    }
  }
  function replaceFound(all: boolean) {
    if (!editor || !find) return;
    try {
      const matches = wordMatches(editor, find, { matchCase: searchCase, wholeWord: searchWhole });
      const chosen = all
        ? matches
        : matches.filter(
            (m) => m.from === editor!.state.selection.from && m.to === editor!.state.selection.to,
          );
      if (!all && !chosen.length) {
        findNext();
        return;
      }
      replaceWordMatches(editor, chosen, replace);
      searchStatus = `Replaced ${chosen.length} matches.`;
      if (!all && chosen.length) findNext();
    } catch (error) {
      searchStatus = (error as Error).message;
    }
  }
</script>

<svelte:window
  onkeydown={async (event) => {
    if (
      !event.defaultPrevented &&
      (event.ctrlKey || event.metaKey) &&
      event.key.toLowerCase() === 'k' &&
      !(event.target as HTMLElement)?.closest('dialog')
    ) {
      event.preventDefault();
      openLink();
      return;
    }
    if (
      !event.defaultPrevented &&
      (event.ctrlKey || event.metaKey) &&
      ['f', 'h'].includes(event.key.toLowerCase()) &&
      !(event.target as HTMLElement)?.closest('dialog')
    ) {
      event.preventDefault();
      setFindOpen(true);
      await tick();
      document.querySelector<HTMLInputElement>('.word-editor .search-panel input')?.focus();
    }
  }}
/>

{#if editor}
  <div class="word-editor editor-body">
    <div class="ribbon-tabs">
      {#each ['Home', 'Insert', 'Layout', 'Review', 'View'] as t}<button
          class={tab === t ? 'selected' : ''}
          onclick={() => setTab(t)}>{t}</button
        >{/each}<span class="ribbon-spacer"></span><button
        class="ribbon-tab-action"
        onclick={() => setFindOpen(!findOpen)}><Search size={15} />Find & replace</button
      >
    </div>
    <div
      class="ribbon grouped-ribbon"
      role="toolbar"
      tabindex={-1}
      aria-label="Document formatting"
      onmousedown={(e) => {
        syncWordSelection(editor!);
        if ((e.target as HTMLElement).closest('button')) e.preventDefault();
      }}
    >
      <RibbonGroup label="History" kind="history">
        <Tool label="Undo" disabled={!canUndo} onclick={() => chain().undo().run()}><Undo2 /></Tool
        ><Tool label="Redo" disabled={!canRedo} onclick={() => chain().redo().run()}><Redo2 /></Tool
        >
      </RibbonGroup>
      {#if tab === 'Home'}
        <RibbonGroup label="Editing" kind="editing">
          <Tool
            label="Format painter"
            active={painterArmed}
            onclick={() => {
              if (painterArmed) {
                painterArmed = false;
                return;
              }
              copiedFormat = captureTextFormat(editor!);
              painterArmed = true;
            }}><Paintbrush /></Tool
          >
          <button class="ribbon-action" onclick={() => chain().selectAll().run()}>Select all</button
          >
        </RibbonGroup>
        <RibbonGroup label="Font" kind="font">
          <div class="ribbon-row">
            <FontPicker
              value={selectionFont?.family || defaultFont}
              mixed={selectionFont?.family === null}
              onchange={(family) => chain().setWordFontFamily(JSON.stringify(family)).run()}
            /><input
              class="font-size-input"
              aria-label="Font size"
              title="Font size in points"
              type="number"
              min="1"
              max="1638"
              step="0.5"
              value={fontPoints}
              placeholder={selectionFont?.size === null ? 'Mixed' : undefined}
              onchange={(e) => applyFontSize(e.currentTarget)}
              onkeydown={(e) => {
                if (e.key === 'Enter' && !e.isComposing && !e.altKey && !e.ctrlKey && !e.metaKey) {
                  e.preventDefault();
                  applyFontSize(e.currentTarget);
                }
              }}
            />

            <Tool label="Grow font" onclick={() => chain().stepWordFontSize(true).run()}
              ><AArrowUp /></Tool
            >
            <Tool label="Shrink font" onclick={() => chain().stepWordFontSize(false).run()}
              ><AArrowDown /></Tool
            >
          </div>
          <div class="ribbon-row">
            <Tool label="Bold" active={isActive('bold')} onclick={() => chain().toggleBold().run()}
              ><Bold /></Tool
            ><Tool
              label="Italic"
              active={isActive('italic')}
              onclick={() => chain().toggleItalic().run()}><Italic /></Tool
            ><Tool
              label="Underline"
              active={isActive('underline')}
              onclick={() => chain().toggleUnderline().run()}><Underline /></Tool
            ><Tool
              label="Strikethrough"
              active={isActive('strike')}
              onclick={() => chain().toggleStrike().run()}><Strikethrough /></Tool
            ><Tool
              label="Subscript"
              active={isActive('subscript')}
              onclick={() => chain().toggleWordScript('subscript').run()}><Subscript /></Tool
            >
            <Tool
              label="Superscript"
              active={isActive('superscript')}
              onclick={() => chain().toggleWordScript('superscript').run()}><Superscript /></Tool
            >
            <label class="color-tool" title="Text color"
              >A<input
                type="color"
                aria-label="Text color"
                value={textStyle.color || '#000000'}
                oninput={(e) => chain().setColor(e.currentTarget.value).run()}
              /></label
            ><Tool
              label="Highlight"
              active={isActive('highlight')}
              onclick={() => chain().toggleHighlight({ color: '#fff0a3' }).run()}
              ><Highlighter /></Tool
            >
          </div>
        </RibbonGroup>
        <RibbonGroup label="Paragraph" kind="paragraph">
          <div class="ribbon-row">
            <WordParagraphLayout {editor} {transaction} mode="alignment" />
            <WordLineSpacing {editor} {transaction} />
          </div>
          <div class="ribbon-row">
            <Tool
              label="Bullet list"
              active={isActive('bulletList')}
              onclick={() => chain().toggleBulletList().run()}><List /></Tool
            ><Tool
              label="Numbered list"
              active={isActive('orderedList')}
              onclick={() => chain().toggleOrderedList().run()}><ListOrdered /></Tool
            ><Tool
              label="Clear formatting"
              onclick={() => chain().unsetAllMarks().clearNodes().run()}><RemoveFormatting /></Tool
            >
            <Tool label="Increase indent" onclick={() => indentParagraphs(editor!)}
              ><IndentIncrease /></Tool
            >
            <Tool label="Decrease indent" onclick={() => indentParagraphs(editor!, true)}
              ><IndentDecrease /></Tool
            >
            <Tool label="Show formatting marks" active={marksVisible} onclick={toggleMarks}
              ><Pilcrow /></Tool
            >
          </div>
        </RibbonGroup>
        <RibbonGroup label="Styles & case" kind="styles">
          <select
            aria-label="Paragraph style"
            value={isActive('heading', { level: 1 })
              ? 'h1'
              : isActive('heading', { level: 2 })
                ? 'h2'
                : isActive('heading', { level: 3 })
                  ? 'h3'
                  : 'p'}
            onchange={(e) =>
              e.currentTarget.value === 'p'
                ? chain().setParagraph().run()
                : chain()
                    .setHeading({ level: Number(e.currentTarget.value.slice(1)) as 1 | 2 | 3 })
                    .run()}
            ><option value="p">Normal text</option><option value="h1">Heading 1</option><option
              value="h2">Heading 2</option
            ><option value="h3">Heading 3</option></select
          ><select
            aria-label="Change case"
            value=""
            onchange={(e) => {
              changeCase(editor!, e.currentTarget.value as CaseMode);
              e.currentTarget.value = '';
            }}
          >
            <option value="" disabled>Change case</option><option value="sentence"
              >Sentence case</option
            ><option value="lower">lowercase</option><option value="upper">UPPERCASE</option><option
              value="title">Capitalize Each Word</option
            ><option value="toggle">tOGGLE cASE</option>
          </select>
        </RibbonGroup>
      {:else if tab === 'Insert'}
        <RibbonGroup label="Headers and footers" kind="stories">
          <button
            class="ribbon-action"
            disabled={!storyChoices.length}
            onclick={() => (storiesOpen = true)}>Headers and footers</button
          >
        </RibbonGroup>
        <RibbonGroup label="Insert" kind="insert">
          <button
            class="ribbon-action"
            onclick={() => chain().insertTable({ rows: 3, cols: 3, withHeaderRow: true }).run()}
            ><Table2 />Table</button
          >
          <button class="ribbon-action" onclick={() => imageInput.current?.click()}
            ><ImagePlus />Image</button
          >
          <button class="ribbon-action" onclick={openLink}><Link />Link</button>
          <button class="ribbon-action" onclick={() => chain().toggleBlockquote().run()}
            ><Quote />Quote</button
          >
          <button class="ribbon-action" onclick={() => chain().setHorizontalRule().run()}
            ><Minus />Divider</button
          >
          <button class="ribbon-action" onclick={() => chain().insertWordFlowBreak('page').run()}
            ><FilePlus2 />Page break</button
          >
          <button class="ribbon-action" onclick={() => chain().insertWordFlowBreak('column').run()}
            ><FilePlus2 />Column break</button
          >
          <select
            aria-label="Insert section break"
            value=""
            onchange={(event) => {
              const kind = event.currentTarget.value as WordSectionStart;
              event.currentTarget.value = '';
              if (!editor || !kind) return;
              try {
                if (!insertWordSectionBreak(editor, content, kind))
                  notify('Place the selection in a body paragraph to insert a section break.');
              } catch (error) {
                notify((error as Error).message);
              }
            }}
          >
            <option value="" disabled>Section break</option>
            <option value="nextPage">Next page</option>
            <option value="continuous">Continuous</option>
            <option value="evenPage">Even page</option>
            <option value="oddPage">Odd page</option>
          </select>
        </RibbonGroup>
        {#if isActive('table')}
          <RibbonGroup label="Table editing" kind="table-tools">
            <button onclick={() => chain().addRowAfter().run()}>Add row</button><button
              onclick={() => chain().addColumnAfter().run()}>Add column</button
            ><button onclick={() => chain().deleteRow().run()}>Delete row</button><button
              onclick={() => chain().deleteColumn().run()}>Delete column</button
            >
            <button disabled={!canMergeCells} onclick={() => chain().mergeCells().run()}
              >Merge cells</button
            >
            <button disabled={!canSplitCell} onclick={() => chain().splitCell().run()}
              >Split cell</button
            >
            <button onclick={() => chain().toggleHeaderRow().run()}>Header row</button><button
              onclick={() => chain().deleteTable().run()}>Delete table</button
            >
          </RibbonGroup>
        {/if}
      {:else if tab === 'Layout'}
        <RibbonGroup label="Page setup" kind="page-setup">
          <WordPageSetup {editor} {content} {transaction} {notify} />
        </RibbonGroup>
        <RibbonGroup label="Spacing" kind="spacing">
          <WordParagraphSpacing {editor} {transaction} />
        </RibbonGroup>
        <RibbonGroup label="Indentation" kind="indentation">
          <WordParagraphLayout {editor} {transaction} mode="indents" />
        </RibbonGroup>
        <RibbonGroup label="Text direction" kind="direction">
          <label class="field"
            ><span>Text direction</span><select
              aria-label="Text direction"
              value={paragraphStyle.direction || 'ltr'}
              onchange={(event) => paragraphLayout({ direction: event.currentTarget.value })}
            >
              <option value="ltr">Left to right</option><option value="rtl">Right to left</option>
            </select></label
          >
        </RibbonGroup>
        <RibbonGroup label="Pagination" kind="pagination">
          <WordParagraphLayout {editor} {transaction} mode="breaks" />
        </RibbonGroup>
      {:else if tab === 'Review'}
        <RibbonGroup label="Statistics" kind="statistics">
          <button
            class="ribbon-action"
            onmousedown={(e) => e.preventDefault()}
            onclick={() => {
              countedSelection = selectedWords;
              countOpen = true;
            }}><WholeWord />Word count</button
          >
        </RibbonGroup>
        <RibbonGroup label="Proofing" kind="proofing">
          <label class="ribbon-check"
            ><input
              type="checkbox"
              checked={spellcheck}
              onchange={(e) => {
                spellcheck = e.currentTarget.checked;
                editor!.view.dom.setAttribute('spellcheck', String(spellcheck));
              }}
            /> Browser spellcheck</label
          >
        </RibbonGroup>
      {:else if tab === 'View'}
        <RibbonGroup label="Show" kind="visibility">
          <label class="ribbon-check"
            ><input
              type="checkbox"
              checked={rulerVisible}
              onchange={(e) => (rulerVisible = e.currentTarget.checked)}
            /> Ruler</label
          >
          <label class="ribbon-check"
            ><input
              type="checkbox"
              checked={navigationVisible}
              onchange={(e) => (navigationVisible = e.currentTarget.checked)}
            /> Navigation pane</label
          >
        </RibbonGroup>
        <RibbonGroup label="Zoom" kind="zoom">
          <label class="ribbon-zoom"
            >Zoom <input
              class="font-size-input"
              aria-label="Document zoom"
              type="number"
              min="50"
              max="150"
              value={zoom}
              onchange={(e) =>
                setZoom(Math.max(50, Math.min(150, e.currentTarget.valueAsNumber || 100)))}
            />%</label
          >
          <button class="ribbon-action" onclick={() => setZoom(100)}>100%</button>
        </RibbonGroup>
      {/if}
    </div>
    {#if findOpen}<SearchBar
        bind:query={find}
        bind:replacement={replace}
        bind:matchCase={searchCase}
        bind:whole={searchWhole}
        status={searchStatus}
        onfind={findNext}
        onreplace={replaceFound}
        onclose={() => {
          setFindOpen(false);
          editor?.commands.focus();
        }}
      />{/if}
    {#if navigationVisible}<nav class="word-navigation" aria-label="Document headings">
        {#each headings as heading}<button
            onclick={() =>
              editor!
                .chain()
                .focus()
                .setTextSelection(heading.pos + 1)
                .scrollIntoView()
                .run()}>{heading.text || 'Empty heading'}</button
          >{:else}<span>No headings in this document.</span>{/each}
      </nav>{/if}
    {#if painterArmed}<div class="find-bar" role="status">
        Select text to apply copied character formatting. Escape cancels.
      </div>{/if}
    <div class="document-scroll">
      <div
        class="paper-wrap"
        class:source-layout={!!sourcePage && !surfaces.plan}
        class:mixed-layout={!!surfaces.plan}
        style={css({
          zoom: zoom / 100,
          ...sourcePageStyle,
          '--surface-canvas-width': surfaces.plan ? `${surfaces.plan.width}px` : undefined,
          '--surface-canvas-height': surfaces.plan ? `${surfaces.plan.height}px` : undefined,
        })}
      >
        {#if rulerVisible}<div class="ruler">
            <span>1</span><span>2</span><span>3</span><span>4</span><span>5</span><span>6</span>
          </div>{/if}
        <div
          class={`paper ${content.paper} margin-${content.margin} ${content.orientation || 'portrait'}`}
        >
          {#if surfaces.plan}<div class="section-pages" aria-hidden="true">
              {#each surfaces.plan.pages as surface, pageIndex}
                <div
                  class="section-page"
                  data-section-id={surface.sectionId}
                  style={css({
                    left: `${surface.left}px`,
                    top: `${surface.top}px`,
                    width: `${surface.width}px`,
                    height: `${surface.height}px`,
                  })}
                >
                  {#each (surfaces.plan.tabBars || []).filter((rule) => rule.page === pageIndex) as rule}
                    <div
                      class="word-tab-bar"
                      style={css({
                        position: 'absolute',
                        background: '#000',
                        width: `${rule.width}px`,
                        left: `${rule.left - surface.left}px`,
                        top: `${rule.top - surface.top}px`,
                        height: `${rule.height}px`,
                      })}
                    ></div>
                  {/each}
                  {#each (surfaces.plan.separators || []).filter((rule) => rule.page === pageIndex) as rule}
                    <div
                      class="word-column-separator"
                      style={css({
                        position: 'absolute',
                        background: '#000',
                        width: '1px',
                        left: `${rule.left - surface.left - 0.5}px`,
                        top: `${rule.top - surface.top}px`,
                        height: `${rule.height}px`,
                      })}
                    ></div>
                  {/each}
                  {#each surface.stories || [] as story}
                    <div
                      class="document-page word-page-story"
                      data-word-story={story.path}
                      style={css({
                        position: 'absolute',
                        left: `${story.left}px`,
                        top: `${story.top}px`,
                        width: `${story.width}px`,
                        height: `${story.height}px`,
                        minHeight: '0',
                        padding: '0',
                      })}
                    >
                      {@html wordStoryGridHtml(
                        surfaces.storyHtml?.[story.renderKey || story.path] || '',
                        story.top,
                      )}
                    </div>
                  {/each}
                </div>
              {/each}
            </div>{/if}
          <div use:attachEditor></div>
        </div>
      </div>
    </div>
    <div class="editor-status">
      {#if surfaces.reason}<span class="section-layout-status">{surfaces.reason}</span>{/if}
      {#if sectionLabel}<span aria-label="Selected sections">{sectionLabel}</span>{/if}
      <span>{words} words</span><span>{editor.getText().length} characters</span><span
        class="ribbon-spacer"
      ></span><span>Spellcheck {spellcheck ? 'on' : 'off'}</span><Tool
        label="Zoom out"
        onclick={() => setZoom(Math.max(50, zoom - 10))}><Minus /></Tool
      ><span>{zoom}%</span><Tool label="Zoom in" onclick={() => setZoom(Math.min(150, zoom + 10))}
        ><Plus /></Tool
      >
    </div>

    {#if editingStory}
      <WordStoryEditor
        title={editingStory.label}
        html={editingStory.html}
        compatibility={content.docxStructure?.compatibility}
        pageLeft={(() => {
          const left = tabSections.find((section) => section.id === editingStory?.sectionId)
            ?.margins.left;
          return left == null ? null : left / 15;
        })()}
        onclose={closeStory}
        onapply={(html) => {
          if (editingStory!.path)
            changeWordStory(editor!, editingStory!.path, editingStory!.html, html);
          else
            createWordStory(
              editor!,
              content,
              storyOptionsExpected,
              editingStory!.sectionId,
              editingStory!.kind,
              editingStory!.slot,
              html,
            );
          void closeStory();
        }}
      />
    {:else if storyLinksOpen}
      <WordStoryLinks
        {content}
        onclose={() => (storyLinksOpen = false)}
        onapply={(sectionId, kind, slot, linked) => {
          changeWordStoryLink(
            editor!,
            content,
            storyOptionsExpected,
            sectionId,
            kind,
            slot,
            linked,
          );
          void closeStory();
        }}
      />
    {:else if storyOptionsOpen}
      <WordStoryOptions
        {content}
        onclose={() => (storyOptionsOpen = false)}
        onapply={(sectionId, options) => {
          changeWordStoryPageOptions(editor!, content, storyOptionsExpected, sectionId, options);
          void closeStory();
        }}
      />
    {:else if storiesOpen}
      <Modal title="Headers and footers" onclose={closeStory}>
        <p>
          Select a header or footer to edit or create. Linked sections share the same text.
          First-page and even-page stories appear when enabled in Page options.
        </p>
        <button
          class="story-choice"
          onclick={() => {
            storyOptionsExpected = content.stories;
            storyOptionsOpen = true;
          }}>Page options</button
        >
        <button
          class="story-choice"
          onclick={() => {
            storyOptionsExpected = content.stories;
            storyLinksOpen = true;
          }}>Section links</button
        >
        {#each storyChoices as choice}
          <button
            class="story-choice"
            disabled={!choice.path && !storyTemplates?.[choice.kind]}
            onclick={() => {
              const part = content.stories?.parts.find((p) => p.path === choice.path);
              storyOptionsExpected = content.stories;
              editingStory = {
                ...choice,
                html: part?.html || storyTemplates![choice.kind]!,
              };
            }}>{choice.label}</button
          >
        {/each}
      </Modal>
    {/if}
    {#if countOpen}<Modal title="Word count" onclose={() => (countOpen = false)}>
        <dl>
          <dt>Words</dt>
          <dd>{words}</dd>
          <dt>Characters (with spaces)</dt>
          <dd>{editor.getText().length}</dd>
          <dt>Characters (without whitespace)</dt>
          <dd>{editor.getText().replace(/\s/g, '').length}</dd>
          <dt>Selected words</dt>
          <dd>{countedSelection}</dd>
        </dl>
        <p>Words are counted by spaces and line breaks.</p>
      </Modal>{/if}
    <input
      hidden
      bind:this={imageInput.current}
      type="file"
      accept="image/png,image/jpeg,image/gif,image/webp"
      oninput={async (e) => {
        const imagePicker = e.currentTarget;
        const file = imagePicker.files?.[0];
        if (file) {
          try {
            chain()
              .setImage({ src: await imageData(file), alt: file.name })
              .run();
          } catch (err) {
            notify((err as Error).message);
          }
        }
        imagePicker.value = '';
      }}
    />{#if linkOpen}<Modal title="Insert a link" onclose={() => setLinkOpen(false)}
        ><label class="field"
          ><span>Text to display</span><input
            aria-label="Text to display"
            bind:value={linkText}
            placeholder="Use address as link text"
          /></label
        ><label class="field"
          ><span>Web or email address</span><input
            use:focusOnMount
            aria-label="Link address"
            value={url}
            oninput={(e) => setUrl(e.currentTarget.value)}
            placeholder="https://example.com"
          /></label
        >
        <div class="modal-actions">
          <button disabled={!isActive('link')} onclick={() => applyLink(true)}>Remove link</button
          ><button class="primary" onclick={() => applyLink()}>Apply link</button>
        </div></Modal
      >{/if}
  </div>
{/if}
