<script lang="ts">
  import SearchBar from './SearchBar.svelte';
  import { paintCSS } from '../drawing-colors';
  import { stackSlide, type StackCommand } from '../slide-stack';
  import { tick } from 'svelte';
  import {
    presentationMatches,
    replacePresentationMatches,
    nextPresentationMatch,
    type SlideMatch,
  } from '../office-search';
  let showSearch = $state(false),
    search = $state(''),
    replacement = $state(''),
    searchCase = $state(false),
    searchWhole = $state(false),
    searchStatus = $state('');
  let lastMatch: SlideMatch | undefined;
  async function findSlide(previous: boolean) {
    try {
      const matches = presentationMatches(content, search, {
        matchCase: searchCase,
        wholeWord: searchWhole,
      });
      const next = nextPresentationMatch(content, matches, slide.id, selected, lastMatch, previous);
      if (!next) {
        searchStatus = 'No matches found.';
        return;
      }
      lastMatch = next;
      setIndex(content.slides.findIndex((s) => s.id === next.slideId));
      setSelected(next.elementId);
      setEditing(null);
      await tick();
      const input = editorRoot?.querySelector<HTMLTextAreaElement>(
        'textarea[aria-label="Object text"]',
      );
      input?.focus();
      input?.setSelectionRange(next.from, next.to);
      searchStatus = `${matches.indexOf(next) + 1} of ${matches.length} matches · slide ${index + 1}`;
    } catch (error) {
      searchStatus = (error as Error).message;
    }
  }
  function replaceSlideText(all: boolean) {
    try {
      const matches = presentationMatches(content, search, {
        matchCase: searchCase,
        wholeWord: searchWhole,
      });
      const chosen = all
        ? matches
        : matches.filter(
            (m) =>
              m.slideId === slide.id &&
              m.elementId === selected &&
              m.from === lastMatch?.from &&
              m.to === lastMatch?.to,
          );
      if (!all && !chosen.length) {
        void findSlide(false);
        return;
      }
      const next = replacePresentationMatches(content, chosen, replacement);
      if (chosen.length) commit(next);
      searchStatus = `Replaced ${chosen.length} matches.`;
      lastMatch = undefined;
    } catch (error) {
      searchStatus = (error as Error).message;
    }
  }
  import RibbonGroup from './RibbonGroup.svelte';
  import { onMount } from 'svelte';
  import FontPicker from './FontPicker.svelte';
  import { fontStack, defaultFont } from '../fonts';
  import Tool from './Tool.svelte';
  import Modal from './Modal.svelte';
  import { css, focusOnMount } from './css';
  import {
    Plus,
    Type,
    Square,
    Circle,
    ImagePlus,
    Undo2,
    Redo2,
    Copy,
    Trash2,
    ChevronUp,
    ChevronDown,
    Play,
    X,
    ChevronLeft,
    ChevronRight,
    Bold,
    Italic,
    Underline,
    AlignLeft,
    AlignCenter,
    AlignRight,
    Layers,
  } from '@lucide/svelte';
  import {
    newSlide,
    textElement,
    uid,
    type DeckContent,
    type Slide,
    type SlideElement,
  } from '../model';
  import { imageData } from '../formats';
  import SlidePreview from './SlidePreview.svelte';
  import SlideText from './SlideText.svelte';
  import SlideOutline from './SlideOutline.svelte';
  import { editOutline, outlineDashes } from '../slide-outline';
  import { resizeSlideElement } from '../slide-geometry';
  import { groupGeometryError } from '../pptx-group-transform';
  import { coarseGroupGrid, quantizeGroupEdit, groupSizeField } from '../pptx-group-edit';
  import { arrangeSlide, moveSlideSelection, type ArrangeCommand } from '../slide-arrange';
  function elementStyle(el: SlideElement) {
    return {
      left: el.x,
      top: el.y,
      width: el.w,
      height: el.h,
      color: el.color,
      background: paintCSS(el.fill, el.fillOpacity),
      fontSize: el.fontSize,
      fontFamily: fontStack(el.fontFamily),
      fontStyle: el.italic ? 'italic' : 'normal',
      textDecoration: el.sourceText ? 'none' : el.underline ? 'underline' : 'none',
      fontWeight: el.bold ? 700 : 400,
      textAlign: el.align,
      borderRadius: el.type === 'ellipse' ? '50%' : undefined,
      transform: el.rotation ? `rotate(${el.rotation}deg)` : undefined,
    };
  }
  let {
    content,
    onChange,
    notify,
  }: {
    content: DeckContent;
    onChange: (content: DeckContent) => void;
    notify: (message: string) => void;
  } = $props();
  let index = $state.raw(0);
  function setIndex(value: typeof index | ((previous: typeof index) => typeof index)) {
    const next = typeof value === 'function' ? value(index) : value;
    if (next !== index) {
      setSelected(null);
      setEditing(null);
    }
    index = next;
  }
  let selected = $state.raw<string | null>(null);
  let selectedIds = $state.raw<string[]>([]);
  let alignToSlide = $state(false);
  let showSelection = $state(false);
  let editorRoot: HTMLDivElement;
  function setSelected(value: typeof selected | ((previous: typeof selected) => typeof selected)) {
    selected = typeof value === 'function' ? value(selected) : value;
    selectedIds = selected ? [selected] : [];
  }
  function toggleSelected(id: string) {
    selectedIds = selectedIds.includes(id)
      ? selectedIds.filter((value) => value !== id)
      : [...selectedIds, id];
    selected = selectedIds.at(-1) || null;
    setEditing(null);
  }
  function arrange(command: ArrangeCommand) {
    try {
      const next = arrangeSlide(
        slide,
        selectedIds,
        command,
        alignToSlide || selectedIds.length === 1,
        aspectRatio,
      );
      if (
        next.elements.some(
          (el, i) =>
            Math.abs(el.x - slide.elements[i].x) > 0.000001 ||
            Math.abs(el.y - slide.elements[i].y) > 0.000001,
        )
      )
        changeSlide(next);
    } catch (error) {
      notify((error as Error).message);
    }
  }
  function stack(command: StackCommand) {
    try {
      const next = stackSlide(slide, selectedIds, command);
      if (next !== slide) changeSlide(next);
    } catch (error) {
      notify((error as Error).message);
    }
  }
  let presenting = $state.raw(false);
  function setPresenting(
    value: typeof presenting | ((previous: typeof presenting) => typeof presenting),
  ) {
    presenting = typeof value === 'function' ? value(presenting) : value;
  }
  let editing = $state.raw<string | null>(null);
  function setEditing(value: typeof editing | ((previous: typeof editing) => typeof editing)) {
    editing = typeof value === 'function' ? value(editing) : value;
  }
  let preview = $state.raw<Slide | null>(null);
  function setPreview(value: typeof preview | ((previous: typeof preview) => typeof preview)) {
    preview = typeof value === 'function' ? value(preview) : value;
  }
  let scale = $state.raw(0.7);
  let aspectRatio = $derived(content.aspectRatio || 16 / 9);
  let canvasHeight = $derived(960 / aspectRatio);
  let verticalScale = $derived(canvasHeight / 540);
  function setScale(value: typeof scale | ((previous: typeof scale) => typeof scale)) {
    scale = typeof value === 'function' ? value(scale) : value;
  }
  const stageWrap: { current: HTMLDivElement | null } = { current: null };
  const imageInput: { current: HTMLInputElement | null } = { current: null };
  const undo: { current: DeckContent[] } = { current: [] };
  const redo: { current: DeckContent[] } = { current: [] };
  let slide = $derived(content.slides[Math.min(index, content.slides.length - 1)]);
  let element = $derived(
    selectedIds.length === 1 ? slide.elements.find((e) => e.id === selected) : undefined,
  );
  const drag: {
    current: {
      startX: number;
      startY: number;
      element: SlideElement;
      slide: Slide;
      ids: string[];
      resize: boolean;
    } | null | null;
  } = { current: null };
  function commit(next: DeckContent) {
    undo.current.push(content);
    if (undo.current.length > 50) undo.current.shift();
    redo.current = [];
    onChange(next);
  }
  function changeSlide(next: Slide) {
    try {
      next = {
        ...next,
        elements: next.elements.map((after) => {
          const before = slide.elements.find((el) => el.id === after.id);
          return before ? quantizeGroupEdit(before, after) : after;
        }),
      };
    } catch (error) {
      notify((error as Error).message);
      return;
    }
    for (const after of next.elements) {
      const before = slide.elements.find((el) => el.id === after.id);
      const error = before && groupGeometryError(before, after);
      if (error) {
        notify(error);
        return;
      }
    }
    if (JSON.stringify(next) !== JSON.stringify(slide))
      commit({ ...content, slides: content.slides.map((s) => (s.id === slide.id ? next : s)) });
  }
  function changeElement(patch: Partial<SlideElement>) {
    if (element)
      changeSlide({
        ...slide,
        elements: slide.elements.map((el) => (el.id === selected ? { ...el, ...patch } : el)),
      });
  }
  async function positionInput(
    input: HTMLInputElement,
    key: 'x' | 'y' | 'w' | 'h',
    commitField = false,
  ) {
    if (!element) return;
    const coarse = !!coarseGroupGrid(element);
    if (coarse !== commitField) return;
    if (!input.value.trim() || !Number.isFinite(Number(input.value))) {
      notify('Enter a finite size or position.');
    } else {
      const value = Math.max(
        key === 'w' || key === 'h' ? 10 : 0,
        Math.min(2000, Number(input.value)),
      );
      changeElement(
        key === 'w' || key === 'h' ? groupSizeField(element, key, value) : { [key]: value },
      );
    }
    if (coarse) {
      await tick();
      input.value = String(Math.round(element[key]));
    }
  }
  function addElement(type: SlideElement['type'], src?: string) {
    const el = textElement(type === 'text' ? 'Your text here' : '', {
      type,
      x: 100,
      y: 100,
      w: type === 'text' ? 600 : 280,
      h: type === 'text' ? 100 : 200,
      fontSize: 36,
      fill: type === 'rect' || type === 'ellipse' ? '#a8beab' : 'transparent',
      src,
    });
    changeSlide({ ...slide, elements: [...slide.elements, el] });
    setSelected(el.id);
  }
  function addSlide(layout: 'title' | 'content' | 'blank' = 'content') {
    const slides = [...content.slides];
    slides.splice(index + 1, 0, newSlide(layout));
    commit({ ...content, slides });
    setIndex(index + 1);
    setSelected(null);
  }
  function duplicateSlide() {
    const copy = structuredClone(slide);
    copy.id = uid();
    copy.elements.forEach((el) => (el.id = uid()));
    const slides = [...content.slides];
    slides.splice(index + 1, 0, copy);
    commit({ ...content, slides });
    setIndex(index + 1);
    setSelected(null);
  }
  function reorder(delta: number) {
    const next = index + delta;
    if (next < 0 || next >= content.slides.length) return;
    const slides = [...content.slides];
    [slides[index], slides[next]] = [slides[next], slides[index]];
    commit({ ...content, slides });
    setIndex(next);
  }
  function history(forward: boolean) {
    const source = forward ? redo : undo,
      dest = forward ? undo : redo,
      previous = source.current.pop();
    if (previous) {
      dest.current.push(content);
      onChange(previous);
      const restoredIndex = Math.min(index, previous.slides.length - 1);
      setIndex(restoredIndex);
      selectedIds = selectedIds.filter((id) =>
        previous.slides[restoredIndex].elements.some((el) => el.id === id),
      );
      selected = selectedIds.at(-1) || null;
      // An inspector control may disappear on undo. Keep subsequent shortcuts in the editor.
      editorRoot?.focus();
    }
  }
  onMount(() => {
    if (!stageWrap.current) return;
    const resize = new ResizeObserver(([entry]) =>
      setScale(
        Math.min(
          1,
          (entry.contentRect.width - 64) / 960,
          (entry.contentRect.height - 48) / canvasHeight,
        ),
      ),
    );
    resize.observe(stageWrap.current);
    return () => resize.disconnect();
  });
  $effect(() => {
    if (!presenting) return;
    const key = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setPresenting(false);
      if (['ArrowRight', 'ArrowDown', ' '].includes(e.key)) {
        e.preventDefault();
        setIndex((i) => Math.min(content.slides.length - 1, i + 1));
      }
      if (['ArrowLeft', 'ArrowUp'].includes(e.key)) {
        e.preventDefault();
        setIndex((i) => Math.max(0, i - 1));
      }
      if (e.key === 'Home') setIndex(0);
      if (e.key === 'End') setIndex(content.slides.length - 1);
    };
    window.addEventListener('keydown', key);
    return () => window.removeEventListener('keydown', key);
  });
</script>

<!-- svelte-ignore a11y_no_noninteractive_tabindex, a11y_no_noninteractive_element_interactions (Presentation editing has application-level object keyboard controls.) -->
<div
  bind:this={editorRoot}
  class="slide-editor editor-body"
  role="application"
  aria-label="Presentation editor"
  tabindex={-1}
  onkeydown={(e) => {
    if (
      !presenting &&
      (e.ctrlKey || e.metaKey) &&
      ['f', 'h'].includes(e.key.toLowerCase()) &&
      !(e.target as HTMLElement).closest('dialog')
    ) {
      e.preventDefault();
      showSearch = true;
      return;
    }
    if (presenting || ['INPUT', 'TEXTAREA', 'SELECT'].includes((e.target as HTMLElement).tagName))
      return;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      selectedIds = slide.elements.map((el) => el.id);
      selected = selectedIds.at(-1) || null;
      setEditing(null);
    }
    if (e.key === 'Escape') {
      setSelected(null);
      setEditing(null);
    }
    if ((e.ctrlKey || e.metaKey) && ['z', 'y'].includes(e.key.toLowerCase())) {
      e.preventDefault();
      history(e.shiftKey || e.key.toLowerCase() === 'y');
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && selectedIds.length) {
      e.preventDefault();
      if (slide.elements.some((el) => selectedIds.includes(el.id) && el.sourceShapeId)) {
        notify(
          'Deleting imported objects requires relationship-aware editing and is not supported yet.',
        );
        return;
      }
      changeSlide({
        ...slide,
        elements: slide.elements.filter((el) => !selectedIds.includes(el.id)),
      });
      setSelected(null);
    }
    if (e.key.startsWith('Arrow') && selectedIds.length) {
      e.preventDefault();
      const step = e.shiftKey ? 10 : 1;
      changeSlide(
        moveSlideSelection(
          slide,
          selectedIds,
          e.key === 'ArrowLeft' ? -step : e.key === 'ArrowRight' ? step : 0,
          e.key === 'ArrowUp' ? -step : e.key === 'ArrowDown' ? step : 0,
        ),
      );
    }
  }}
>
  <div class="ribbon-tabs">
    <button class="selected">Home</button><button
      class="ribbon-tab-action"
      onclick={() => (showSearch = !showSearch)}>Find & replace</button
    ><span class="ribbon-spacer"></span><button
      class="ribbon-tab-action"
      onclick={() => setPresenting(true)}><Play size={15} />Present from this slide</button
    >
  </div>
  <div class="ribbon grouped-ribbon" role="toolbar" aria-label="Presentation tools">
    <RibbonGroup label="History" kind="history">
      <Tool label="Undo" disabled={!content || !undo.current.length} onclick={() => history(false)}
        ><Undo2 /></Tool
      ><Tool label="Redo" disabled={!content || !redo.current.length} onclick={() => history(true)}
        ><Redo2 /></Tool
      >
    </RibbonGroup>
    <RibbonGroup label="Slides" kind="slides">
      <button class="ribbon-action" onclick={() => addSlide()}><Plus />New slide</button><select
        aria-label="New slide layout"
        value=""
        onchange={(e) => {
          addSlide(e.currentTarget.value as 'title' | 'content' | 'blank');
          e.currentTarget.value = '';
        }}
        ><option value="" disabled>Layouts</option><option value="title">Title slide</option><option
          value="content">Title and content</option
        ><option value="blank">Blank slide</option></select
      >
    </RibbonGroup>
    <RibbonGroup label="Insert" kind="slide-insert">
      <button class="ribbon-action" onclick={() => addElement('text')}><Type />Text box</button
      ><button class="ribbon-action" aria-label="Add rectangle" onclick={() => addElement('rect')}
        ><Square />Rectangle</button
      ><button class="ribbon-action" aria-label="Add ellipse" onclick={() => addElement('ellipse')}
        ><Circle />Ellipse</button
      ><button class="ribbon-action" onclick={() => imageInput.current?.click()}
        ><ImagePlus />Image</button
      >
    </RibbonGroup>
    <RibbonGroup label="Organize" kind="organize">
      <Tool label="Duplicate slide" onclick={duplicateSlide}><Copy /></Tool><Tool
        label="Delete slide"
        disabled={content.slides.length === 1}
        onclick={() => {
          commit({ ...content, slides: content.slides.filter((s) => s.id !== slide.id) });
          setIndex(Math.max(0, index - 1));
          setSelected(null);
        }}><Trash2 /></Tool
      >
    </RibbonGroup>
    <RibbonGroup label="Arrange" kind="arrange">
      <select
        aria-label="Order objects"
        value=""
        disabled={!selectedIds.length}
        onchange={(e) => {
          stack(e.currentTarget.value as StackCommand);
          e.currentTarget.value = '';
        }}
      >
        <option value="" disabled>Order objects</option>
        <option value="front">Bring to front</option>
        <option value="forward">Bring forward</option>
        <option value="backward">Send backward</option>
        <option value="back">Send to back</option>
      </select>
      <select
        aria-label="Align relative to"
        disabled={!selectedIds.length}
        value={alignToSlide || selectedIds.length === 1 ? 'slide' : 'selection'}
        onchange={(e) => (alignToSlide = e.currentTarget.value === 'slide')}
      >
        <option value="selection" disabled={selectedIds.length < 2}>Selected objects</option>
        <option value="slide">Slide</option>
      </select>
      <select
        aria-label="Arrange objects"
        value=""
        disabled={!selectedIds.length}
        onchange={(e) => {
          arrange(e.currentTarget.value as ArrangeCommand);
          e.currentTarget.value = '';
        }}
      >
        <option value="" disabled>Align / distribute</option>
        {#each [['left', 'Align left'], ['center', 'Align center'], ['right', 'Align right'], ['top', 'Align top'], ['middle', 'Align middle'], ['bottom', 'Align bottom'], ['horizontal', 'Distribute horizontally'], ['vertical', 'Distribute vertically']] as [value, label]}
          <option
            {value}
            disabled={['horizontal', 'vertical'].includes(value) &&
              !alignToSlide &&
              selectedIds.length === 2}>{label}</option
          >
        {/each}
      </select>
      <button
        class="ribbon-action"
        aria-pressed={showSelection}
        onclick={() => (showSelection = !showSelection)}>Selection pane</button
      >
    </RibbonGroup>
    <RibbonGroup label="Slide show" kind="slideshow"
      ><button class="primary small" onclick={() => setPresenting(true)}
        ><Play size={15} />Present</button
      >
    </RibbonGroup>
  </div>
  {#if showSearch}<SearchBar
      bind:query={search}
      bind:replacement
      bind:matchCase={searchCase}
      bind:whole={searchWhole}
      status={searchStatus}
      onfind={findSlide}
      onreplace={replaceSlideText}
      onclose={() => {
        showSearch = false;
        editorRoot?.focus();
      }}
    />{/if}
  <div class="slide-workspace">
    <aside class="slide-filmstrip" aria-label="Slides">
      {#each content.slides as s, i}<button
          class={`slide-thumb ${i === index ? 'selected' : ''}`}
          aria-label={`Slide ${i + 1}`}
          onclick={() => {
            setIndex(i);
            setSelected(null);
            setEditing(null);
          }}><span>{i + 1}</span><SlidePreview slide={s} {aspectRatio} /></button
        >{/each}<button class="add-slide" onclick={() => addSlide()}
        ><Plus size={17} />Add slide</button
      >
      <div class="slide-order">
        <Tool label="Move slide up" disabled={index === 0} onclick={() => reorder(-1)}
          ><ChevronUp /></Tool
        ><Tool
          label="Move slide down"
          disabled={index === content.slides.length - 1}
          onclick={() => reorder(1)}><ChevronDown /></Tool
        >
      </div>
    </aside>
    <div class="slide-main">
      <div class="stage-wrap" bind:this={stageWrap.current}>
        <div class="stage-space" style={css({ width: 960 * scale, height: canvasHeight * scale })}>
          <!-- svelte-ignore a11y_no_noninteractive_tabindex (The canvas receives focus for arrow-key object movement.) -->
          <div
            class="slide-stage"
            role="application"
            style={css({
              height: canvasHeight,
              background: paintCSS(slide.background, slide.backgroundOpacity),
              transform: `scale(${scale})`,
            })}
            tabindex={0}
            aria-label="Slide canvas"
            onpointerdown={(e) => {
              if (e.target === e.currentTarget) {
                setSelected(null);
                setEditing(null);
              }
            }}
            onpointermove={(e) => {
              if (!drag.current) return;
              const d = drag.current,
                dx = (e.clientX - d.startX) / scale,
                dy = (e.clientY - d.startY) / scale / verticalScale;
              setPreview(
                d.resize
                  ? {
                      ...d.slide,
                      elements: d.slide.elements.map((el) =>
                        el.id === d.element.id
                          ? {
                              ...el,
                              ...resizeSlideElement(
                                d.element,
                                dx,
                                dy * verticalScale,
                                verticalScale,
                              ),
                            }
                          : el,
                      ),
                    }
                  : moveSlideSelection(d.slide, d.ids, Math.round(dx), Math.round(dy)),
              );
            }}
            onpointerup={() => {
              if (drag.current && preview) {
                if (
                  preview.elements.some((el, i) =>
                    (['x', 'y', 'w', 'h'] as const).some(
                      (key) => el[key] !== slide.elements[i][key],
                    ),
                  )
                )
                  changeSlide(preview);
              }
              drag.current = null;
              setPreview(null);
            }}
            onpointercancel={() => {
              drag.current = null;
              setPreview(null);
            }}
          >
            {#each slide.elements as original}{@const el =
                preview?.elements.find((el) => el.id === original.id) || original}
              <div
                class={`slide-element ${selectedIds.includes(el.id) ? 'selected' : ''} ${el.type}`}
                role="group"
                style={css({
                  ...elementStyle(el),
                  top: el.y * verticalScale,
                  height: el.h * verticalScale,
                })}
                aria-label={el.type === 'text' ? el.text : `${el.type} object`}
                onpointerdown={(e) => {
                  if (editing === el.id) return;
                  e.stopPropagation();
                  e.currentTarget.parentElement?.focus();
                  if (e.ctrlKey || e.metaKey || e.shiftKey) {
                    toggleSelected(el.id);
                    return;
                  }
                  if (!selectedIds.includes(el.id)) setSelected(el.id);
                  const resize = (e.target as HTMLElement).classList.contains('resize-handle');
                  drag.current = {
                    startX: e.clientX,
                    startY: e.clientY,
                    element: el,
                    resize,
                    slide,
                    ids: [...selectedIds],
                  };
                  e.currentTarget.setPointerCapture(e.pointerId);
                }}
                ondblclick={() => {
                  if (el.type === 'text') {
                    setSelected(el.id);
                    setEditing(el.id);
                  }
                }}
              >
                {#if el.type === 'image'}<img
                    src={el.src}
                    alt={el.text || 'Slide image'}
                    draggable={false}
                    style:transform={`scale(${el.flipH ? -1 : 1}, ${el.flipV ? -1 : 1})`}
                  />{:else}{#if el.type === 'text'}{#if editing === el.id}<textarea
                        use:focusOnMount
                        aria-label="Edit slide text"
                        value={el.text}
                        style={css({
                          color: el.color,
                          fontSize: el.fontSize,
                          fontFamily: fontStack(el.fontFamily),
                          fontStyle: el.italic ? 'italic' : 'normal',
                          textDecoration: el.underline ? 'underline' : 'none',
                          fontWeight: el.bold ? 700 : 400,
                          textAlign: el.align,
                        })}
                        onblur={(e) => {
                          changeElement({ text: e.currentTarget.value });
                          setEditing(null);
                        }}></textarea>{:else}<SlideText
                        element={el}
                      />{/if}{:else}{/if}{/if}{#if selectedIds.length === 1 && el.id === selected && !editing}<i
                    class="selection-handle tl"
                  ></i><i class="selection-handle tr"></i><i class="selection-handle bl"></i><i
                    class="resize-handle"
                  ></i>{/if}
                <svg
                  aria-hidden="true"
                  width={el.w}
                  height={el.h * verticalScale}
                  style="position:absolute;inset:0;overflow:visible;pointer-events:none"
                >
                  <SlideOutline
                    outline={el.outline}
                    ellipse={el.type === 'ellipse'}
                    width={el.w}
                    height={el.h * verticalScale}
                  />
                </svg>
              </div>{/each}
          </div>
        </div>
      </div>
      <div class="speaker-notes">
        <label for="speaker-notes">Speaker notes</label><textarea
          id="speaker-notes"
          value={slide.notes}
          placeholder="A little reminder for when it’s your time to talk…"
          oninput={(e) => changeSlide({ ...slide, notes: e.currentTarget.value })}></textarea>
      </div>
    </div>
    <aside class="slide-inspector">
      {#if showSelection}
        <h3>Selection</h3>
        <div class="object-selection" aria-label="Slide objects">
          {#each slide.elements as el, i}
            <button
              class="inspector-button"
              aria-pressed={selectedIds.includes(el.id)}
              onclick={() => toggleSelected(el.id)}
              >{i + 1}. {el.text.slice(0, 45) || `${el.type} object`}</button
            >
          {/each}
        </div>
      {/if}
      <h3>
        {element ? 'Format object' : selectedIds.length > 1 ? 'Format selection' : 'Slide design'}
      </h3>
      {#if element}{#if element.type === 'text'}<label class="field"
            ><span>Text</span><textarea
              aria-label="Object text"
              value={element.text}
              oninput={(e) => changeElement({ text: e.currentTarget.value })}
              rows={5}></textarea></label
          >
          <label class="field"
            ><span>Font</span><FontPicker
              value={element.fontFamily || defaultFont}
              onchange={(fontFamily) => changeElement({ fontFamily })}
            /></label
          >
          <div class="inspector-row">
            <label class="field"
              ><span>Size</span><input
                aria-label="Object font size"
                type="number"
                min={8}
                max={160}
                value={element.fontSize}
                oninput={(e) =>
                  changeElement({
                    fontSize: Math.max(8, Math.min(160, Number(e.currentTarget.value))),
                  })}
              /></label
            ><label class="field"
              ><span>Color</span><input
                aria-label="Object text color"
                type="color"
                value={element.color}
                oninput={(e) => changeElement({ color: e.currentTarget.value })}
              /></label
            >
          </div>
          <div class="tool-group">
            <Tool
              label="Bold text"
              active={element.bold}
              onclick={() => changeElement({ bold: !element.bold })}><Bold /></Tool
            ><Tool
              label="Italic text"
              active={element.italic || false}
              onclick={() => changeElement({ italic: !element.italic })}><Italic /></Tool
            >
            <Tool
              label="Underline text"
              active={element.underline || false}
              onclick={() => changeElement({ underline: !element.underline })}><Underline /></Tool
            >
            {#each ['left', 'center', 'right'] as const as align, i}{@const Icon = [
                AlignLeft,
                AlignCenter,
                AlignRight,
              ][i]}<Tool
                label={`Align ${align}`}
                active={element.align === align}
                onclick={() => changeElement({ align })}><Icon /></Tool
              >{/each}
          </div>{/if}{#if element.type !== 'image'}<label class="field"
            ><span>Fill color</span><input
              aria-label="Shape fill"
              type="color"
              value={element.fill === 'transparent' ? '#ffffff' : element.fill}
              oninput={(e) =>
                changeElement({
                  fill: e.currentTarget.value,
                  fillOpacity: element.fill === 'transparent' ? 1 : element.fillOpacity,
                })}
            /></label
          >
          <button
            class="inspector-button"
            onclick={() => changeElement({ fill: 'transparent', fillOpacity: 1 })}>No fill</button
          >
          <label class="field"
            ><span>Fill transparency (%)</span><input
              aria-label="Fill transparency"
              type="number"
              min="0"
              max="100"
              value={100 * (1 - (element.fillOpacity ?? 1))}
              onchange={(e) =>
                changeElement({
                  fillOpacity:
                    1 - Math.max(0, Math.min(100, Number(e.currentTarget.value) || 0)) / 100,
                })}
            /></label
          >
        {/if}
        {#if element.type !== 'image'}
          <h4>Outline</h4>
          <label class="field"
            ><span>Outline color</span><input
              aria-label="Outline color"
              type="color"
              value={element.outline?.color && element.outline.color !== 'transparent'
                ? element.outline.color
                : '#000000'}
              oninput={(e) =>
                changeElement({
                  outline: editOutline(element, {
                    color: e.currentTarget.value,
                    width: element.outline?.width || 1,
                    opacity:
                      element.outline?.color === 'transparent'
                        ? 1
                        : (element.outline?.opacity ?? 1),
                  }),
                })}
            /></label
          >
          <button
            class="inspector-button"
            onclick={() =>
              changeElement({ outline: editOutline(element, { color: 'transparent' }) })}
            >No outline</button
          >
          <label class="field"
            ><span>Outline width</span><input
              aria-label="Outline width"
              type="number"
              min="0"
              max="100"
              step="0.25"
              value={element.outline?.width ?? 1}
              onchange={(e) =>
                changeElement({
                  outline: editOutline(element, {
                    width: Math.max(0, Math.min(100, Number(e.currentTarget.value) || 0)),
                  }),
                })}
            /></label
          >
          <label class="field"
            ><span>Outline transparency (%)</span><input
              aria-label="Outline transparency"
              type="number"
              min="0"
              max="100"
              value={100 * (1 - (element.outline?.opacity ?? 1))}
              onchange={(e) =>
                changeElement({
                  outline: editOutline(element, {
                    opacity:
                      1 - Math.max(0, Math.min(100, Number(e.currentTarget.value) || 0)) / 100,
                  }),
                })}
            /></label
          >
          <label class="field"
            ><span>Outline dashes</span><select
              aria-label="Outline dashes"
              value={element.outline?.dash ?? 'solid'}
              onchange={(e) =>
                changeElement({
                  outline: editOutline(element, {
                    dash: e.currentTarget.value as (typeof outlineDashes)[number],
                  }),
                })}
            >
              {#each outlineDashes as dash}<option value={dash}
                  >{{
                    solid: 'Solid',
                    dot: 'Dot',
                    dash: 'Dash',
                    lgDash: 'Long dash',
                    dashDot: 'Dash dot',
                    lgDashDot: 'Long dash dot',
                    lgDashDotDot: 'Long dash dot dot',
                    sysDash: 'Short dash',
                    sysDot: 'Short dot',
                    sysDashDot: 'Short dash dot',
                    sysDashDotDot: 'Short dash dot dot',
                  }[dash]}</option
                >{/each}
            </select></label
          >
        {/if}
        <h4>Position & size</h4>
        <div class="position-fields">
          {#each ['x', 'y', 'w', 'h'] as const as key}<label class="field"
              ><span>{{ x: 'X', y: 'Y', w: 'Width', h: 'Height' }[key]}</span><input
                aria-label={`Object ${key}`}
                type="number"
                value={Math.round(element[key])}
                oninput={(e) => positionInput(e.currentTarget, key)}
                onchange={(e) => positionInput(e.currentTarget, key, true)}
              /></label
            >{/each}
        </div>
        <label class="field"
          ><span>Rotation (degrees)</span><input
            type="number"
            aria-label="Object rotation"
            min="0"
            max="360"
            step="1"
            value={element.rotation || 0}
            onchange={(e) =>
              changeElement({
                rotation: Math.max(0, Math.min(360, Number(e.currentTarget.value) || 0)),
              })}
          /></label
        >
        {#if element.type === 'image'}
          <button
            class="inspector-button"
            aria-pressed={!!element.flipH}
            onclick={() => changeElement({ flipH: !element!.flipH })}>Flip horizontally</button
          >
          <button
            class="inspector-button"
            aria-pressed={!!element.flipV}
            onclick={() => changeElement({ flipV: !element!.flipV })}>Flip vertically</button
          >
        {/if}
        <button class="inspector-button" onclick={() => stack('front')}
          ><Layers size={15} />Bring to front</button
        ><button
          class="inspector-button"
          onclick={() => {
            const copy = { ...element, id: uid(), x: element.x + 20, y: element.y + 20 };
            changeSlide({ ...slide, elements: [...slide.elements, copy] });
            setSelected(copy.id);
          }}><Copy size={15} />Duplicate object</button
        ><button
          class="inspector-button danger"
          onclick={() => {
            changeSlide({ ...slide, elements: slide.elements.filter((el) => el.id !== selected) });
            setSelected(null);
          }}><Trash2 size={15} />Delete object</button
        >{:else if selectedIds.length > 1}
        <p>{selectedIds.length} objects selected</p>
        <p class="muted">
          Use Arrange to align or distribute the selection. Drag a selected object to move them
          together. Arrow keys move the selection; Shift uses larger steps.
        </p>
        <button class="inspector-button" onclick={() => setSelected(null)}>Clear selection</button>
      {:else}<label class="field"
          ><span>Background</span><input
            aria-label="Slide background"
            type="color"
            value={slide.background}
            oninput={(e) => changeSlide({ ...slide, background: e.currentTarget.value })}
          /></label
        >
        <h4>Color palette</h4>
        <div class="palette">
          {#each ['#f5f3ea', '#ffffff', '#263d34', '#e6edf4', '#f4e8df', '#e5e9dd'] as color}<button
              aria-label={`Background ${color}`}
              style={css({ background: color })}
              onclick={() => changeSlide({ ...slide, background: color })}
            ></button>{/each}
        </div>
        <p class="muted">
          Select an object to edit its appearance. Ctrl-click or Shift-click to select several
          objects. Drag to move them, or use the corner handle to resize one object.
        </p>
        <p class="muted">Double-click a text box to write directly on the slide.</p>{/if}
    </aside>
  </div>
  <div class="editor-status">
    <span>Slide {index + 1} of {content.slides.length}</span><span class="ribbon-spacer"
    ></span><span
      >{Math.abs(aspectRatio - 16 / 9) < 0.001
        ? 'Widescreen · 16:9'
        : Math.abs(aspectRatio - 4 / 3) < 0.001
          ? 'Standard · 4:3'
          : `Custom · ${aspectRatio.toFixed(3)}:1`}</span
    ><span>{Math.round(scale * 100)}%</span>
  </div>
  <input
    bind:this={imageInput.current}
    hidden
    type="file"
    accept="image/png,image/jpeg,image/webp,image/gif"
    oninput={async (e) => {
      const imagePicker = e.currentTarget;
      const file = imagePicker.files?.[0];
      if (file)
        try {
          addElement('image', await imageData(file));
        } catch (error) {
          notify((error as Error).message);
        }
      imagePicker.value = '';
    }}
  />{#if presenting}<div
      class="presentation-mode"
      role="dialog"
      aria-modal="true"
      aria-label="Slideshow"
    >
      <SlidePreview {slide} {aspectRatio} />
      <div class="presentation-controls">
        <Tool label="Previous slide" disabled={index === 0} onclick={() => setIndex(index - 1)}
          ><ChevronLeft /></Tool
        ><span>{index + 1} / {content.slides.length}</span><Tool
          label="Next slide"
          disabled={index === content.slides.length - 1}
          onclick={() => setIndex(index + 1)}><ChevronRight /></Tool
        ><Tool label="Exit presentation" onclick={() => setPresenting(false)}><X /></Tool>
      </div>
    </div>{/if}
</div>

<style>
  .object-selection {
    display: flex;
    flex-direction: column;
    gap: 4px;
    max-height: 200px;
    overflow-y: auto;
    margin-bottom: 20px;
  }
  .object-selection button {
    flex-shrink: 0;
    padding: 7px 8px;
    min-height: 32px;
    text-align: left;
    font-size: 12px;
    color: #475569;
    border: 1px solid #dce2eb;
    overflow-wrap: anywhere;
  }
  .object-selection button[aria-pressed='true'] {
    background: #fbede8;
    color: #974d32;
    border-color: #ca967f;
  }
</style>
