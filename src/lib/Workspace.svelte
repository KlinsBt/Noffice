<script lang="ts">
  import { onMount, tick } from 'svelte';
  import { initializeFonts } from '../fonts';
  import Tool from './Tool.svelte';
  import Modal from './Modal.svelte';
  import { css } from './css';
  import {
    ArrowDownToLine,
    ArrowLeft,
    ArrowRight,
    Check,
    CheckCheck,
    ChevronDown,
    ChevronRight,
    CircleHelp,
    Clock3,
    Copy,
    Download,
    FileText,
    FolderOpen,
    Grid2X2,
    HardDrive,
    History,
    List,
    LoaderCircle,
    LockKeyhole,
    MoreHorizontal,
    Plus,
    Search,
    Settings2,
    ShieldCheck,
    Star,
    Table2,
    Trash2,
    Upload,
    X,
    Printer,
    RotateCcw,
  } from '@lucide/svelte';
  import { appInfo, newFile, samples, uid, type AppKind, type OfficeFile } from '../model';
  import { getVersions, listFiles, saveFile, type Version } from '../storage';
  import { download, exportOffice, importFile, nativeBackup, unchangedOfficeOriginal, writeCSV } from '../formats';
  import JSZip from 'jszip';
  type View = 'home' | 'all' | 'favorites' | 'trash' | AppKind | 'templates';
  const kinds: AppKind[] = ['word', 'excel', 'powerpoint'];
  function relativeDate(time: number) {
    const mins = Math.floor((Date.now() - time) / 60000);
    return mins < 1
      ? 'Just now'
      : mins < 60
        ? `${mins} min ago`
        : mins < 1440
          ? `${Math.floor(mins / 60)} hr ago`
          : new Date(time).toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
  }
  import WordEditor from './WordEditor.svelte';
  import SheetEditor from './SheetEditor.svelte';
  import { editWorkbookStructure } from '../workbook-structure';
  import { renameWorkbookTable } from '../workbook-table-rename';
  import { needsTableMetadata, hydrateTableMetadata } from '../workbook-tables';
  import { needsWordStructure, hydrateWordStructure } from '../word-structure';
  import { needsPresentationText, hydratePresentationText } from '../presentation-migration';
  import SlideEditor from './SlideEditor.svelte';
  import AppIcon from './AppIcon.svelte';
  let files = $state.raw<OfficeFile[]>([]);
  function setFiles(value: typeof files | ((previous: typeof files) => typeof files)) {
    files = typeof value === 'function' ? value(files) : value;
    fileRef.current = files;
  }
  let loading = $state.raw(true);
  function setLoading(value: typeof loading | ((previous: typeof loading) => typeof loading)) {
    loading = typeof value === 'function' ? value(loading) : value;
  }
  let activeId = $state.raw<string | null>(null);
  let editorBusy = $state(false);
  let wordEditor = $state.raw<{
    preparePdf: (
      content: import('../model').WordContent,
    ) => Promise<import('../word-pdf-model').WordPdfSnapshot>;
    preparePrint: () => Promise<void>;
    checkpointExport: () => void;
    prepareExport: (
      content: import('../model').WordContent,
    ) => Promise<import('../word-export-layout').WordExportLayout | undefined>;
  }>();
  let printing = false;
  async function printWord() {
    if (active && wordHydrationFailure?.id === active.id) {
      setModal(null);
      setError(`Print unavailable: ${wordHydrationFailure.message}`);
      return;
    }
    if (printing || !wordEditor) return;
    const current = wordEditor;
    printing = true;
    try {
      setModal(null);
      await tick();
      await current.preparePrint();
      if (wordEditor === current) window.print();
    } catch (error) {
      notify((error as Error).message);
    } finally {
      printing = false;
    }
  }
  function setEditorBusy(value: boolean) {
    editorBusy = value;
  }
  function setActiveId(value: typeof activeId | ((previous: typeof activeId) => typeof activeId)) {
    activeId = typeof value === 'function' ? value(activeId) : value;
    if (activeId) setModal(null);
  }
  let view = $state.raw<View>('home');
  function setView(value: typeof view | ((previous: typeof view) => typeof view)) {
    view = typeof value === 'function' ? value(view) : value;
  }
  let query = $state.raw('');
  function setQuery(value: typeof query | ((previous: typeof query) => typeof query)) {
    query = typeof value === 'function' ? value(query) : value;
  }
  let filter = $state.raw<'all' | AppKind>('all');
  function setFilter(value: typeof filter | ((previous: typeof filter) => typeof filter)) {
    filter = typeof value === 'function' ? value(filter) : value;
  }
  let listMode = $state.raw(true);
  function setListMode(value: typeof listMode | ((previous: typeof listMode) => typeof listMode)) {
    listMode = typeof value === 'function' ? value(listMode) : value;
  }
  let sort = $state.raw('updated');
  function setSort(value: typeof sort | ((previous: typeof sort) => typeof sort)) {
    sort = typeof value === 'function' ? value(sort) : value;
  }
  let toast = $state.raw('');
  function setToast(value: typeof toast | ((previous: typeof toast) => typeof toast)) {
    toast = typeof value === 'function' ? value(toast) : value;
  }
  let error = $state.raw('');
  let errorKind = $state<'save' | 'general'>('general');
  function setError(
    value: typeof error | ((previous: typeof error) => typeof error),
    kind: 'save' | 'general' = 'general',
  ) {
    error = typeof value === 'function' ? value(error) : value;
    errorKind = kind;
  }
  let saveStatus = $state.raw('Saved on this device');
  function setSaveStatus(
    value: typeof saveStatus | ((previous: typeof saveStatus) => typeof saveStatus),
  ) {
    saveStatus = typeof value === 'function' ? value(saveStatus) : value;
  }
  let modal = $state.raw<
    'files' | 'new' | 'help' | 'settings' | 'export' | 'versions' | 'compatibility' | null
  >(null);
  function setModal(value: typeof modal | ((previous: typeof modal) => typeof modal)) {
    modal = typeof value === 'function' ? value(modal) : value;
  }
  let menu = $state.raw<string | null>(null);
  function setMenu(value: typeof menu | ((previous: typeof menu) => typeof menu)) {
    menu = typeof value === 'function' ? value(menu) : value;
  }
  let busy = $state.raw(false);
  let pdfBusy = $state(false);
  function setBusy(value: typeof busy | ((previous: typeof busy) => typeof busy)) {
    busy = typeof value === 'function' ? value(busy) : value;
  }
  let versions = $state.raw<Version[]>([]);
  function setVersions(value: typeof versions | ((previous: typeof versions) => typeof versions)) {
    versions = typeof value === 'function' ? value(versions) : value;
  }
  let restoreKey = $state.raw(0);
  function setRestoreKey(
    value: typeof restoreKey | ((previous: typeof restoreKey) => typeof restoreKey),
  ) {
    restoreKey = typeof value === 'function' ? value(restoreKey) : value;
  }
  let storage = $state.raw('');
  function setStorage(value: typeof storage | ((previous: typeof storage) => typeof storage)) {
    storage = typeof value === 'function' ? value(storage) : value;
  }
  let offline = $state.raw(!navigator.onLine);
  function setOffline(value: typeof offline | ((previous: typeof offline) => typeof offline)) {
    offline = typeof value === 'function' ? value(offline) : value;
  }
  const input: { current: HTMLInputElement | null } = { current: null };
  const pending = { current: new Map<string, OfficeFile>() };
  const revisions = { current: new Map<string, number>() };
  const timer: { current: ReturnType<typeof setTimeout> | undefined } = { current: undefined };
  const saving = { current: false };
  const fileRef: { current: OfficeFile[] } = { current: [] };
  const toastTimer: { current: ReturnType<typeof setTimeout> | undefined } = { current: undefined };

  const notify = (text: string) => {
    setToast(text);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToast(''), 5500);
  };
  const flush = async () => {
    if (saving.current || !pending.current.size) return;
    saving.current = true;
    setSaveStatus('Saving…');
    try {
      while (pending.current.size) {
        const [id, file] = pending.current.entries().next().value!;
        const revision = await saveFile(file, revisions.current.get(id) || 0);
        revisions.current.set(id, revision);
        if (pending.current.get(id) === file) pending.current.delete(id);
      }
      setSaveStatus('Saved on this device');
      if (errorKind === 'save') setError('');
    } catch (err) {
      setSaveStatus('Not saved');
      setError(
        err instanceof Error
          ? err.message
          : 'Could not save locally. Export a backup before leaving.',
        'save',
      );
    } finally {
      saving.current = false;
    }
  };
  const queue = (file: OfficeFile) => {
    pending.current.set(file.id, file);
    setSaveStatus('Unsaved changes');
    clearTimeout(timer.current);
    timer.current = setTimeout(() => void flush(), 450);
  };
  onMount(() => {
    let alive = true;
    void initializeFonts().catch((err) => notify(`Local fonts could not load: ${err.message}`));
    (async () => {
      try {
        let data = await listFiles();
        if (!data.length) {
          data = samples();
          for (const file of data) {
            file.revision = await saveFile(file, 0);
          }
        }
        if (alive) {
          data.forEach((f) => {
            if (!revisions.current.has(f.id)) revisions.current.set(f.id, f.revision);
          });
          // Imports can finish while the initial IndexedDB snapshot is loading.
          // Keep files and edits already made in this mounted workspace.
          fileRef.current = [
            ...new Map([...data, ...fileRef.current].map((f) => [f.id, f])).values(),
          ];
          setFiles(fileRef.current);
        }
      } catch (err) {
        if (alive) setError(`Local storage could not open: ${(err as Error).message}`);
      } finally {
        if (alive) setLoading(false);
      }
    })();
    return () => {
      alive = false;
    };
  });
  onMount(() => {
    const beforeUnload = (e: BeforeUnloadEvent) => {
      if (pending.current.size || editorBusy) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    const hidden = () => {
      if (document.visibilityState === 'hidden') void flush();
    };
    const key = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p' && wordEditor) {
        e.preventDefault();
        void printWord();
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
        e.preventDefault();
        void flush();
      }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'o') {
        e.preventDefault();
        input.current?.click();
      }
    };
    const connection = () => setOffline(!navigator.onLine);
    window.addEventListener('beforeunload', beforeUnload);
    document.addEventListener('visibilitychange', hidden);
    window.addEventListener('keydown', key);
    window.addEventListener('online', connection);
    window.addEventListener('offline', connection);
    return () => {
      window.removeEventListener('beforeunload', beforeUnload);
      document.removeEventListener('visibilitychange', hidden);
      window.removeEventListener('keydown', key);
      window.removeEventListener('online', connection);
      window.removeEventListener('offline', connection);
    };
  });
  let active = $derived(files.find((f) => f.id === activeId));
  let wordHydrationFailure = $state<{ id: string; message: string } | null>(null);
  $effect(() => {
    const file = active;
    if (
      !file ||
      (!needsTableMetadata(file) && !needsWordStructure(file) && !needsPresentationText(file))
    )
      return;
    let alive = true;
    wordHydrationFailure = null;
    void (
      needsPresentationText(file)
        ? hydratePresentationText(file)
        : needsWordStructure(file)
          ? hydrateWordStructure(file)
          : hydrateTableMetadata(file)
    )
      .then((next) => {
        // Metadata hydration never overwrites a newer edit or writes a new storage revision.
        if (alive && fileRef.current.find((f) => f.id === file.id) === file)
          setFiles(fileRef.current.map((f) => (f === file ? next : f)));
      })
      .catch((error) => {
        if (alive) {
          if (needsWordStructure(file))
            wordHydrationFailure = { id: file.id, message: (error as Error).message };
          notify((error as Error).message);
        }
      });
    return () => {
      alive = false;
    };
  });
  function update(id: string, patch: Partial<OfficeFile>) {
    const current = fileRef.current.find((f) => f.id === id);
    if (!current) return;
    const next = { ...current, ...patch, updatedAt: Date.now() };
    fileRef.current = fileRef.current.map((f) => (f.id === id ? next : f));
    setFiles(fileRef.current);
    queue(next);
  }
  function add(file: OfficeFile) {
    fileRef.current = [...fileRef.current, file];
    setFiles(fileRef.current);
    queue(file);
    setActiveId(file.id);
    setModal(null);
  }
  function create(kind: AppKind, template = false) {
    const file = template ? samples().find((f) => f.kind === kind)! : newFile(kind);
    if (template) file.name += ' — copy';
    add(file);
  }
  function duplicate(file: OfficeFile) {
    const next = {
      ...structuredClone(file),
      id: uid(),
      revision: 0,
      name: file.name + ' — copy',
      createdAt: Date.now(),
      updatedAt: Date.now(),
      trashed: false,
    };
    add(next);
    setMenu(null);
  }
  async function importFiles(list: FileList | File[]) {
    setBusy(true);
    let count = 0;
    for (const file of Array.from(list)) {
      try {
        add(await importFile(file));
        count++;
      } catch (err) {
        setError(
          `${file.name}: ${err instanceof Error ? err.message : 'Could not open this file.'}`,
        );
      }
    }
    setBusy(false);
    if (count) notify(`Opened ${count} ${count === 1 ? 'file' : 'files'} on this device.`);
  }
  async function doExport(
    format: 'office' | 'native' | 'text' | 'html' | 'csv' | 'original' | 'pdf',
  ) {
    if (!active || busy || pdfBusy) return;
    if (editorBusy) {
      notify('Wait for the table headings to finish updating before exporting.');
      return;
    }
    if (format === 'pdf') {
      pdfBusy = true;
      setModal(null);
    } else setBusy(true);
    try {
      if ((format === 'office' || format === 'pdf') && wordHydrationFailure?.id === active.id)
        throw Error(wordHydrationFailure.message);
      if (format === 'native')
        download(nativeBackup(active), active.name + '.noffice', 'application/json');
      else if (format === 'original' && active.original)
        download(active.original.data, active.original.name);
      else if (format === 'pdf' && active.content.kind === 'word') {
        const current = active;
        if (!wordEditor || current.content.kind !== 'word')
          throw Error('The document is still opening.');
        if (current.original && current.content.docxStructure) {
          const { assertWordPdfSourceDecorations } = await import('../docx-pdf-source');
          await assertWordPdfSourceDecorations(current.original.data);
        }
        const snapshot = await wordEditor.preparePdf(current.content);
        const { exportWordPdf } = await import('../word-pdf');
        const result = await exportWordPdf(snapshot);
        if (active?.id !== current.id || JSON.stringify(active.content) !== snapshot.content)
          throw Error('Document changed before export. Try exporting again.');
        download(result, current.name + '.pdf');
      } else if (format === 'office') {
        const current = active;
        const exportedEditor = wordEditor;
        const original = current.content.kind === 'word' ? await unchangedOfficeOriginal(current) : undefined;
        const layout =
          current.content.kind === 'word' && !original
            ? await wordEditor?.prepareExport(current.content)
            : undefined;
        // A completed IndexedDB save can replace the file record without
        // changing its contents. Revision bookkeeping does not stale layout.
        if (
          active?.id !== current.id ||
          JSON.stringify(active.content) !== JSON.stringify(current.content)
        )
          throw Error('Document changed before export. Try exporting again.');
        const result = original ?? await exportOffice(current, layout);
        if (
          active?.id !== current.id ||
          JSON.stringify(active.content) !== JSON.stringify(current.content)
        )
          throw Error('Document changed before export. Try exporting again.');
        download(result, current.name + '.' + appInfo[current.kind].extension);
        if (current.content.kind === 'word') exportedEditor?.checkpointExport();
      } else if (format === 'csv' && active.content.kind === 'excel')
        download(
          '\uFEFF' + writeCSV(active.content.sheets[0], active.content.sheets),
          active.name + '.csv',
          'text/csv;charset=utf-8',
        );
      else if (active.content.kind === 'word') {
        const dom = new DOMParser().parseFromString(active.content.html, 'text/html');
        download(
          format === 'html'
            ? `<!doctype html><html><head><meta charset="utf-8"><title>${active.name.replaceAll('<', '&lt;')}</title></head><body>${active.content.html}</body></html>`
            : [...dom.body.children].map((el) => el.textContent).join('\n\n'),
          active.name + (format === 'html' ? '.html' : '.txt'),
          format === 'html' ? 'text/html' : 'text/plain',
        );
      }
      if (format !== 'pdf') setModal(null);
      notify('Your download is ready.');
    } catch (err) {
      if (format !== 'pdf') setModal(null);
      setError(`Export failed: ${(err as Error).message}`);
    } finally {
      if (format === 'pdf') pdfBusy = false;
      else setBusy(false);
    }
  }
  function navigate(next: View) {
    setActiveId(null);
    setView(next);
    setQuery('');
    setFilter('all');
    setMenu(null);
    void flush();
  }
  let visible = $derived(
    files
      .filter(
        (f) =>
          (view === 'trash' ? f.trashed : !f.trashed) &&
          (view !== 'favorites' || f.favorite) &&
          (!kinds.includes(view as AppKind) || f.kind === view) &&
          (filter === 'all' || f.kind === filter) &&
          f.name.toLowerCase().includes(query.toLowerCase()),
      )
      .sort((a, b) => (sort === 'name' ? a.name.localeCompare(b.name) : b.updatedAt - a.updatedAt)),
  );
</script>

<div
  class={`workspace ${active ? 'editing-workspace' : ''}`}
  role="application"
  aria-label="Noffice workspace"
  aria-busy={loading || busy}
  ondragover={(e) => {
    if (e.dataTransfer!.types.includes('Files')) e.preventDefault();
  }}
  ondrop={(e) => {
    if (e.dataTransfer!.files.length) {
      e.preventDefault();
      void importFiles(e.dataTransfer!.files);
    }
  }}
>
  <div class="main-area">
    <header class="topbar">
      {#if active}<div class="document-heading">
          <Tool
            label="Choose a mode"
            onclick={() => {
              setActiveId(null);
              void flush();
            }}><ArrowLeft /></Tool
          ><AppIcon kind={active.kind} small /><input
            aria-label="File name"
            value={active.name}
            maxLength={200}
            oninput={(e) => update(active.id, { name: e.currentTarget.value })}
            onblur={() => {
              if (!active.name.trim())
                update(active.id, { name: 'Untitled ' + appInfo[active.kind].type.toLowerCase() });
            }}
          /><Tool
            label={active.favorite ? 'Unstar file' : 'Star file'}
            active={active.favorite}
            onclick={() => update(active.id, { favorite: !active.favorite })}
            ><Star size={16} /></Tool
          >
        </div>{:else}<button
          class="brand"
          aria-label="Noffice home"
          onclick={() => navigate('home')}
          ><span class="brand-mark">n<span>·</span></span><span
            >noffice<span class="brand-dot">.</span></span
          ></button
        >{/if}
      <div class="topbar-right">
        {#if active}<span class={`save-indicator ${saveStatus === 'Not saved' ? 'failed' : ''}`}
            >{#if !editorBusy && saveStatus === 'Saved on this device'}<CheckCheck
                size={16}
              />{:else}<Clock3 size={16} />{/if}{editorBusy
              ? 'Updating table headings…'
              : saveStatus}</span
          ><Tool
            label="Version history"
            onclick={async () => {
              await flush();
              setVersions(await getVersions(active.id));
              setModal('versions');
            }}><History /></Tool
          ><button class="primary small" disabled={editorBusy} onclick={() => setModal('export')}
            ><Download size={15} />Export <ChevronDown size={13} /></button
          >{:else}<span class="private-badge"><LockKeyhole size={13} />Private by default</span
          ><Tool label="Workspace help" onclick={() => setModal('help')}
            ><CircleHelp size={19} /></Tool
          ><Tool
            label="Storage and backups"
            onclick={async () => {
              const estimate = await navigator.storage?.estimate();
              setStorage(
                estimate?.usage !== undefined
                  ? `${(estimate.usage / 1024 / 1024).toFixed(1)} MB used on this device`
                  : 'Storage information unavailable',
              );
              setModal('settings');
            }}><Settings2 size={19} /></Tool
          >{/if}
      </div>
    </header>
    {#if error}<div class="error-banner" role="alert">
        <span>{error}</span>{#if errorKind === 'save'}<button onclick={() => void flush()}
            >Retry save</button
          >{/if}<Tool label="Dismiss error" onclick={() => setError('')}><X size={16} /></Tool>
      </div>{/if}{#if active}<div class="app-switcher">
        {#each kinds as kind}<button
            class={active.kind === kind ? 'selected' : ''}
            onclick={() => {
              if (active.kind === kind) return;
              const existing = [...files]
                .filter((f) => f.kind === kind && !f.trashed)
                .sort((a, b) => b.updatedAt - a.updatedAt)[0];
              if (existing) setActiveId(existing.id);
              else create(kind);
              void flush();
            }}><AppIcon {kind} small />{appInfo[kind].name}</button
          >{/each}<span class="ribbon-spacer"></span>{#if active.warnings.length > 0}<button
            class="compatibility-button"
            onclick={() => setModal('compatibility')}
            >Conversion notes <span>{active.warnings.length}</span></button
          >{/if}<button onclick={() => input.current?.click()}><FolderOpen size={15} />Open</button
        ><button
          onclick={() => {
            setModal('new');
          }}><Plus size={15} />New</button
        >
      </div>
      {#key active.id + ':' + restoreKey}{#if needsWordStructure(active)}
          {#if wordHydrationFailure?.id === active.id}
            <section aria-label="Document recovery">
              <p role="alert">{wordHydrationFailure.message}</p>
              <button onclick={() => void doExport('native')}>Export Noffice backup</button>
              <button onclick={() => void doExport('original')}>Download original DOCX</button>
            </section>
          {:else}<p role="status">Opening document...</p>{/if}
        {:else if active.content.kind === 'word'}<WordEditor
            bind:this={wordEditor}
            content={active.content}
            onChange={(content) => update(active.id, { content })}
            {notify}
          />{:else}{#if active.content.kind === 'excel'}<SheetEditor
              content={active.content}
              onStructure={(content, edit) => editWorkbookStructure({ ...active, content }, edit)}
              onTableRename={(content, edit) => renameWorkbookTable({ ...active, content }, edit)}
              onBusy={setEditorBusy}
              onChange={(content) => update(active.id, { content })}
              {notify}
            />{:else}<SlideEditor
              content={active.content}
              onChange={(content) => update(active.id, { content })}
              {notify}
            />{/if}{/if}{/key}{:else}<main class="mode-landing">
        <div class="eyebrow">WORDS. NUMBERS. IDEAS.</div>
        <h1>What will you create today?</h1>
        <p class="landing-intro">Choose your mode. Make it yours.</p>
        <div class="mode-cards">
          {#each kinds as kind}<button
              class={`mode-card ${kind}`}
              disabled={loading}
              onclick={() => create(kind)}
              aria-label={`Start ${appInfo[kind].name}`}
            >
              <AppIcon {kind} />
              <h2>{appInfo[kind].name}</h2>
              <p>
                {{
                  word: 'A blank page for your next big idea.',
                  excel: 'Bring your numbers into focus.',
                  powerpoint: 'Give your story a stage.',
                }[kind]}
              </p>
              <span>Start creating <ArrowRight size={17} /></span>
            </button>{/each}
        </div>
        <div class="landing-actions">
          <button class="secondary" onclick={() => input.current?.click()}
            ><Upload size={16} />Open a file</button
          >
          <button
            class="text-button"
            disabled={loading}
            onclick={() => {
              setView('all');
              setModal('files');
            }}><Clock3 size={16} />Recent files</button
          >
        </div>
        <footer class="landing-footer">
          <span><LockKeyhole size={14} />Saved on your device. Works offline.</span><span
            >Open source · MIT licensed</span
          >
        </footer>
      </main>{/if}
  </div>
  <input
    bind:this={input.current}
    hidden
    multiple
    type="file"
    accept=".docx,.xlsx,.pptx,.csv,.txt,.html,.htm,.noffice"
    oninput={(e) => {
      if (e.currentTarget.files) void importFiles(e.currentTarget.files);
      e.currentTarget.value = '';
    }}
  />{#if busy}<div class="busy-overlay" role="status">
      <LoaderCircle class="spin" />Working on your file, on this device…
    </div>{/if}{#if pdfBusy}<div class="toast" role="status">
      <LoaderCircle class="spin" /><span>Preparing PDF on this device…</span>
    </div>{:else if toast}<div class="toast" role="status">
      <Check size={17} /><span>{toast}</span><Tool
        label="Dismiss notification"
        onclick={() => setToast('')}><X size={15} /></Tool
      >
    </div>{/if}{#if modal === 'files'}<Modal
      title="Open a file"
      wide
      onclose={() => setModal(null)}
    >
      <div class="file-dialog-actions">
        <button class="secondary" onclick={() => input.current?.click()}
          ><Upload size={16} />Browse this device</button
        >
        <button class="secondary" onclick={() => setModal('new')}
          ><Plus size={16} />Create new</button
        >
      </div>
      <div class="filter-tabs file-dialog-tabs">
        {#each [['all', 'Recent'], ['favorites', 'Starred'], ['trash', 'Trash']] as item}<button
            class:selected={view === item[0]}
            onclick={() => {
              setView(item[0] as View);
              setMenu(null);
            }}>{item[1]}</button
          >{/each}
      </div>
      <section class="files-section">
        <div class="section-heading">
          <h2>
            {view === 'home'
              ? 'Pick up where you left off'
              : view === 'favorites'
                ? 'Starred files'
                : view === 'trash'
                  ? 'Deleted files'
                  : 'Your files'}<span class="count-pill">{visible.length}</span>
          </h2>
          <div class="file-view-actions">
            <select
              aria-label="Sort files"
              value={sort}
              onchange={(e) => setSort(e.currentTarget.value)}
              ><option value="updated">Last modified</option><option value="name">Name A–Z</option
              ></select
            >
            <div class="view-toggle">
              <Tool label="List view" active={listMode} onclick={() => setListMode(true)}
                ><List size={16} /></Tool
              ><Tool label="Grid view" active={!listMode} onclick={() => setListMode(false)}
                ><Grid2X2 size={16} /></Tool
              >
            </div>
          </div>
        </div>
        <div class="file-filter-bar">
          <div class="filter-tabs">
            {#each ['all', ...kinds] as const as kind}<button
                class={filter === kind ? 'selected' : ''}
                onclick={() => setFilter(kind)}
                >{kind === 'all' ? 'All files' : appInfo[kind].plural}</button
              >{/each}
          </div>
          <label class="search-input"
            ><Search size={16} /><input
              aria-label="Search files"
              placeholder="Find a file…"
              value={query}
              oninput={(e) => setQuery(e.currentTarget.value)}
            /><kbd>⌕</kbd></label
          >
        </div>
        {#if loading}<div class="loading">
            <LoaderCircle class="spin" />Loading your workspace…
          </div>{:else}{#if visible.length === 0}<div class="empty-state">
              <FolderOpen size={36} />
              <h3>{query ? 'No files found' : 'A little room for something new.'}</h3>
              <p>
                {query
                  ? 'Try a different name or file type.'
                  : 'Create a file or open one from your device to get started.'}
              </p>
              <button class="secondary" onclick={() => setModal('new')}
                ><Plus size={15} />Create a file</button
              >
            </div>{:else}<div class={listMode ? 'file-list' : 'file-grid'}>
              {#if listMode}<div class="file-list-header">
                  <span>Name</span><span>Type</span><span>Last modified</span><span></span>
                </div>{/if}{#each visible as file}<div
                  class={`file-row ${menu === file.id ? 'menu-open' : ''}`}
                >
                  {#if !listMode}<button
                      class={`file-cover ${file.kind}`}
                      onclick={() => {
                        if (!file.trashed) setActiveId(file.id);
                      }}
                      >{#if file.kind === 'word'}<div class="mini-document">
                          <strong>{file.name}</strong><i></i><i></i><i></i><i></i><i></i>
                        </div>{:else}{#if file.kind === 'excel'}<div class="mini-sheet">
                            {#each Array.from({ length: 20 }) as _, i}<span
                                >{i % 4 ? (i + 1) * 125 : ''}</span
                              >{/each}
                          </div>{:else}<div class="mini-slide">
                            <strong>{file.name}</strong><span>A new perspective.</span>
                          </div>{/if}{/if}</button
                    >{/if}<button
                    class="file-name"
                    onclick={() => {
                      if (file.trashed) {
                        update(file.id, { trashed: false });
                        notify('File restored.');
                      } else setActiveId(file.id);
                    }}
                    ><AppIcon kind={file.kind} small /><span
                      ><strong>{file.name || 'Untitled'}</strong>{#if !listMode}<small
                          >Edited {relativeDate(file.updatedAt)}</small
                        >{/if}</span
                    >{#if file.favorite}<Star
                        size={13}
                        class="starred-icon"
                        fill="currentColor"
                      />{/if}</button
                  ><span class="file-type">{appInfo[file.kind].type}</span><span class="file-date"
                    >{relativeDate(file.updatedAt)}</span
                  >
                  <div class="file-menu">
                    <Tool
                      label={`More options for ${file.name}`}
                      onclick={() => setMenu(menu === file.id ? null : file.id)}
                      ><MoreHorizontal size={20} /></Tool
                    >{#if menu === file.id}<button
                        class="menu-dismiss"
                        aria-label="Close file menu"
                        onclick={() => setMenu(null)}
                      ></button>
                      <div class="dropdown-menu">
                        {#if file.trashed}<button
                            onclick={() => {
                              update(file.id, { trashed: false });
                              setMenu(null);
                              notify('File restored.');
                            }}><RotateCcw size={15} />Restore file</button
                          >{:else}<button
                            onclick={() => {
                              update(file.id, { favorite: !file.favorite });
                              setMenu(null);
                            }}
                            ><Star size={15} />{file.favorite
                              ? 'Remove star'
                              : 'Add to starred'}</button
                          ><button onclick={() => duplicate(file)}
                            ><Copy size={15} />Duplicate</button
                          ><button
                            onclick={() => {
                              setActiveId(file.id);
                              setModal('export');
                              setMenu(null);
                            }}><Download size={15} />Export file</button
                          ><button
                            class="danger"
                            onclick={() => {
                              update(file.id, { trashed: true });
                              setMenu(null);
                              notify('Moved to Trash. You can restore it anytime.');
                            }}><Trash2 size={15} />Move to Trash</button
                          >{/if}
                      </div>{/if}
                  </div>
                </div>{/each}
            </div>{/if}{/if}
      </section></Modal
    >{/if}{#if modal === 'new'}<Modal
      title="What would you like to create?"
      onclose={() => setModal(null)}
      ><p class="modal-intro">A new beginning, in whatever shape you need.</p>
      <div class="new-options">
        {#each kinds as kind}<button onclick={() => create(kind)}
            ><AppIcon {kind} /><span
              ><strong>Blank {appInfo[kind].type.toLowerCase()}</strong><small
                >Open in {appInfo[kind].name}</small
              ></span
            ><ArrowRight size={18} /></button
          >{/each}
      </div>
      <h3>Start from an example</h3>
      <div class="template-options">
        {#each kinds as kind}<button class="secondary" onclick={() => create(kind, true)}
            >{appInfo[kind].type} template</button
          >{/each}
      </div>
      <button class="import-option" onclick={() => input.current?.click()}
        ><Upload size={17} />Or open a file from your device</button
      ></Modal
    >{/if}{#if modal === 'help'}<Modal
      title="A workspace that feels like yours."
      onclose={() => setModal(null)}
      wide
      ><p>
        Word mode is for documents, Excel mode is for spreadsheets, and PowerPoint mode is for
        slides. Switch between them at the top of any open file.
      </p>
      <div class="help-grid">
        <div>
          <h3>A few useful shortcuts</h3>
          <p><kbd>Ctrl / ⌘ + S</kbd>Save locally</p>
          <p><kbd>Ctrl / ⌘ + O</kbd>Open a file</p>
          <p><kbd>Ctrl / ⌘ + Z</kbd>Undo an edit</p>
          <p><kbd>Ctrl / ⌘ + B</kbd>Bold in Word</p>
          <p><kbd>Shift + click</kbd>Select a cell range</p>
          <p><kbd>← → / Esc</kbd>Navigate / exit a slideshow</p>
        </div>
        <div>
          <h3>Your files stay with you</h3>
          <p>
            Changes save automatically in this browser. Export a .noffice backup to keep the full
            editable model and retained original file. Browser data can be cleared, so keep backups
            of important work.
          </p>
          <h3>About Office compatibility</h3>
          <p>
            This is an early independent office suite, with limited DOCX, XLSX, and PPTX conversion.
            It does not yet match Microsoft Office desktop functionality. Check conversion notes and
            review exported files.
          </p>
        </div>
      </div>
      <p class="muted">
        Noffice is an independent MIT project, not affiliated with Microsoft. Microsoft Office,
        Word, Excel, and PowerPoint are Microsoft trademarks.
      </p></Modal
    >{/if}{#if modal === 'settings'}<Modal
      title="Your local workspace"
      onclose={() => setModal(null)}
      ><div class="settings-summary">
        <HardDrive size={30} />
        <div>
          <strong>{files.length}files on this device</strong>
          <p>{storage}</p>
        </div>
      </div>
      <p>
        No account or document upload is required. Documents and editor assets run in your browser.
      </p>
      <button
        class="secondary full"
        onclick={async () => {
          try {
            const granted = await navigator.storage?.persist();
            notify(
              granted
                ? 'Persistent storage is enabled.'
                : 'Your browser manages storage automatically. Keep exported backups.',
            );
          } catch {
            notify('This browser does not offer persistent storage.');
          }
        }}><ShieldCheck size={17} />Request persistent storage</button
      ><button
        class="secondary full"
        onclick={async () => {
          const zip = new JSZip();
          files.forEach((f) =>
            zip.file(`${f.id}/${f.name.replace(/[<>:"/\\|?*]/g, '_')}.noffice`, nativeBackup(f)),
          );
          download(await zip.generateAsync({ type: 'blob' }), 'noffice-workspace-backup.zip');
          notify('Backup downloaded. Extract the ZIP, then open its .noffice files to restore.');
        }}><Download size={17} />Back up every file</button
      >
      <p class="muted">
        Workspace backups contain individual .noffice files, including files in Trash. Extract the
        ZIP and open those files to restore them.
      </p></Modal
    >{/if}{#if modal === 'export' && active}<Modal
      title="Take your work with you"
      onclose={() => setModal(null)}
      ><p class="modal-intro">Export “{active.name}” to your device.</p>
      <div class="export-options">
        <button onclick={() => void doExport('office')}
          ><AppIcon kind={active.kind} small /><span
            ><strong>{appInfo[active.kind].extension.toUpperCase()} file</strong><small
              >Editable in Microsoft {active.kind === 'word'
                ? 'Word'
                : active.kind === 'excel'
                  ? 'Excel'
                  : 'PowerPoint'}</small
            ></span
          ><Download size={17} /></button
        ><button onclick={() => void doExport('native')}
          ><ShieldCheck size={24} /><span
            ><strong>Noffice backup</strong><small
              >Full editable model and retained original · .noffice</small
            ></span
          ><Download size={17} /></button
        >{#if active.content.kind === 'word'}<button onclick={() => void doExport('html')}
            ><FileText size={22} /><span>Web page · .html</span><Download size={17} /></button
          ><button onclick={() => void doExport('text')}
            ><FileText size={22} /><span>Plain text · .txt</span><Download size={17} /></button
          ><button onclick={() => void doExport('pdf')}
            ><FileText size={22} /><span>PDF file</span><Download size={17} /></button
          ><button onclick={() => void printWord()}
            ><Printer size={22} /><span>Print / Save as PDF</span><ChevronRight size={17} /></button
          >{/if}{#if active.content.kind === 'excel'}<button onclick={() => void doExport('csv')}
            ><Table2 size={22} /><span>First sheet as CSV · values only</span><Download
              size={17}
            /></button
          >{/if}{#if active.original}<button onclick={() => void doExport('original')}
            ><ArrowDownToLine size={22} /><span>Download unchanged original</span><Download
              size={17}
            /></button
          >{/if}
      </div>
      <p class="export-note">
        Office exports include supported content only. Advanced formatting and objects may change or
        be omitted. Keep a Noffice backup and review the exported file.
      </p></Modal
    >{/if}{#if modal === 'versions' && active}<Modal
      title="Version history"
      onclose={() => setModal(null)}
      ><p class="modal-intro">The last 20 local saves. Restoring creates a new version.</p>
      <div class="version-list">
        {#each versions as version}<div>
            <History size={18} />
            <div>
              <strong>{version.name}</strong><small
                >{new Date(version.at).toLocaleString()}· v{version.revision}</small
              >
            </div>
            <button
              onclick={() => {
                update(active.id, { name: version.name, content: version.content });
                setRestoreKey((k) => k + 1);
                setModal(null);
                notify('Version restored.');
              }}>Restore</button
            >
          </div>{/each}
      </div></Modal
    >{/if}{#if modal === 'compatibility' && active}<Modal
      title="Conversion notes"
      onclose={() => setModal(null)}
      ><p>This file was converted into Noffice’s editable format.</p>
      {#each active.warnings as warning, i}<p class="compatibility-note">
          {warning}
        </p>{/each}{#if active.original}<button
          class="secondary"
          onclick={() => void doExport('original')}
          ><Download size={16} />Download unchanged original</button
        >{/if}</Modal
    >{/if}
</div>
