import { Mark, type Editor } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { isHistoryTransaction } from '@tiptap/pm/history';
import { Mapping, RemoveMarkStep, ReplaceStep } from '@tiptap/pm/transform';
import { closeHistory } from '@tiptap/pm/history';
import type { EditorView } from '@tiptap/pm/view';

const checkpoints = new WeakMap<EditorView, () => void>();
/** A successful DOCX save starts a new typing session without adding an undo
 * event or rewriting existing text, marks, selection or saved content. */
export function checkpointWordEditSession(view: EditorView) {
  checkpoints.get(view)?.();
}

const validSession = (value: unknown): value is string =>
  typeof value === 'string' && /^[a-f\d]{32}$/.test(value);

/** Document metadata belongs to ordinary undo and typed content, never HTML. */
export function initializeWordEditSession(editor: Editor, value?: string) {
  editor.view.dispatch(editor.state.tr
    .setDocAttribute('wordInitialEditSession', validSession(value) ? value : null)
    .setMeta('addToHistory', false).setMeta('preventUpdate', true));
}

/** New paragraphs have no retained XML run map. Keep text entered in each
 * editing session separate so DOCX save/reopen does not reshape an older run.
 * This is local provenance, with no visible formatting or revision UI. */
export const WordEditRun = Mark.create({
  name: 'wordEditRun',
  addOptions: () => ({ retained: false }),
  addGlobalAttributes() {
    return this.options.retained ? [{ types: ['doc'], attributes: {
      wordInitialEditSession: { default: null, rendered: false, parseHTML: () => null },
    } }] : [];
  },
  addAttributes: () => ({
    session: {
      default: null,
      parseHTML: (el) => el.getAttribute('data-word-edit-run'),
      renderHTML: (attrs) =>
        validSession(attrs.session) ? { 'data-word-edit-run': attrs.session } : {},
    },
  }),
  parseHTML: () => [
    {
      tag: 'span[data-word-edit-run]',
      getAttrs: (el) => (validSession(el.getAttribute('data-word-edit-run')) ? {} : false),
    },
  ],
  renderHTML: ({ HTMLAttributes }) => ['span', HTMLAttributes, 0],
  addProseMirrorPlugins() {
    let session = crypto.randomUUID().replaceAll('-', '');
    const type = this.type;
    const retained = this.options.retained;
    let compositionOrigin: { session: string; document: Editor['state']['doc'] } | undefined;
    return [
      new Plugin({
        key: new PluginKey('wordEditRun'),
        view(view) {
          checkpoints.set(view, () => {
            compositionOrigin = undefined;
            session = crypto.randomUUID().replaceAll('-', '');
            view.dispatch(closeHistory(view.state.tr).setMeta('addToHistory', false));
          });
          return {
            destroy: () => {
              checkpoints.delete(view);
            },
          };
        },
        appendTransaction(transactions, _before, next) {
          if (transactions.some(isHistoryTransaction)) return null;
          const steps = transactions.flatMap((tr) =>
            tr.steps.map((step, index) => ({ step, doc: tr.docs[index],
              composition: tr.getMeta('composition') !== undefined })),
          );
          const tr = next.tr;
          if (retained && !validSession(next.doc.attrs.wordInitialEditSession) &&
              transactions.some((transaction) => transaction.docChanged &&
                transaction.getMeta('addToHistory') !== false)) {
            const previous = new Set<string>();
            _before.doc.descendants((node) => {
              const mark = node.marks.find((mark) => mark.type === type);
              if (validSession(mark?.attrs.session)) previous.add(mark.attrs.session);
            });
            // A single legacy typing session is recoverable. Multiple older
            // sessions without their origin cannot establish an ordering.
            if (previous.size <= 1) {
              const initial = [...previous][0] || session;
              tr.setDocAttribute('wordInitialEditSession', initial);
              if (!previous.size && steps.some(({ composition }) => composition))
                compositionOrigin = { session: initial, document: _before.doc };
            }
          }
          for (let index = 0; index < steps.length; index++) {
            const { step, doc, composition } = steps[index];
            if (composition && step instanceof RemoveMarkStep && step.mark.type === type) {
              // Chromium replaces the candidate's DOM wrappers during IME.
              // ProseMirror reports their loss as a composition mark removal.
              // Retain this invisible save provenance on surviving text; it is
              // not a user formatting edit and must not merge old/new DOCX runs.
              const mapping = new Mapping(steps.slice(index + 1).map(({ step }) => step.getMap()));
              const start = mapping.map(step.from, 1), end = mapping.map(step.to, -1);
              if (start < end) tr.addMark(start, end, step.mark);
              continue;
            }
            if (
              !(step instanceof ReplaceStep) ||
              step.slice.openStart ||
              step.slice.openEnd ||
              !step.slice.size
            )
              continue;
            let textOnly = true;
            step.slice.content.forEach((node) => {
              if (!node.isText && node.type.name !== 'wordHyphen') textOnly = false;
            });
            const from = doc.resolve(step.from),
              to = doc.resolve(step.to);
            if (
              !textOnly ||
              (!from.sameParent(to) && !retained) ||
              from.parent.type.name !== 'paragraph' ||
              (from.parent.attrs.sourceParagraph && !retained)
            )
              continue;
            const mapping = new Mapping(steps.slice(index + 1).map(({ step }) => step.getMap()));
            const start = mapping.map(step.from, 1);
            const end = mapping.map(step.from + step.slice.size, -1);
            if (start < end) tr.addMark(start, end, type.create({ session }));
          }
          if (compositionOrigin &&
              tr.doc.attrs.wordInitialEditSession === compositionOrigin.session &&
              tr.doc.content.eq(compositionOrigin.document.content)) {
            // Canceling the first IME candidate restores the untouched model,
            // including the absence of a first saved editing session.
            tr.setDocAttribute('wordInitialEditSession', null);
            compositionOrigin = undefined;
          }
          if (!tr.steps.length) return null;
          // Adding provenance must not replace an explicit pending formatting
          // choice, including the empty set used when leaving a hyperlink.
          if (next.storedMarks !== null) tr.setStoredMarks(next.storedMarks);
          return tr;
        },
      }),
    ];
  },
});

export function wordRunSession(
  marks: readonly { type: string; attrs?: Record<string, unknown> }[],
) {
  const value = marks.find((mark) => mark.type === 'wordEditRun')?.attrs?.session;
  return validSession(value) ? value : undefined;
}
