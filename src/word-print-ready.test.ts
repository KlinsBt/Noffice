import { afterEach, beforeEach, expect, it, vi } from 'vitest';
import { Editor } from '@tiptap/core';
import { wordExtensions } from './word-extensions';
import { prepareWordPrint } from './word-print-ready';

let editor: Editor;
let fonts: PropertyDescriptor | undefined;
beforeEach(() => {
  vi.useFakeTimers();
  fonts = Object.getOwnPropertyDescriptor(document, 'fonts');
  Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: Object.assign(new EventTarget(), { ready: Promise.resolve(), status: 'loaded' }),
  });
  editor = new Editor({ extensions: wordExtensions(), content: '<p>Before</p>' });
});
afterEach(() => {
  editor.destroy();
  if (fonts) Object.defineProperty(document, 'fonts', fonts);
  else Reflect.deleteProperty(document, 'fonts');
  vi.restoreAllMocks();
  vi.useRealTimers();
});

it('waits for the latest edit to settle and adds no document or history changes', async () => {
  let finished = false;
  const ready = prepareWordPrint(editor.view).then(() => {
    finished = true;
  });
  await vi.advanceTimersByTimeAsync(32);
  editor.commands.insertContent('Latest ');
  const edited = editor.getJSON();
  await vi.advanceTimersByTimeAsync(48);
  expect(finished).toBe(false);
  await vi.advanceTimersByTimeAsync(32);
  await ready;
  expect(editor.getJSON()).toEqual(edited);
  editor.commands.undo();
  expect(editor.getText()).toBe('Before');
  expect(editor.can().undo()).toBe(false);
});

it('waits for pending fonts before accepting stable layout', async () => {
  let finishFonts!: () => void;
  Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: Object.assign(new EventTarget(), {
      ready: new Promise<void>((resolve) => {
        finishFonts = resolve;
      }),
      status: 'loaded',
    }),
  });
  let finished = false;
  const ready = prepareWordPrint(editor.view).then(() => {
    finished = true;
  });
  await vi.advanceTimersByTimeAsync(200);
  expect(finished).toBe(false);
  finishFonts();
  await vi.advanceTimersByTimeAsync(80);
  await ready;
  expect(finished).toBe(true);
});

it('rejects a closed editor instead of printing another document', async () => {
  const rejected = expect(prepareWordPrint(editor.view)).rejects.toThrow('closed');
  await vi.advanceTimersByTimeAsync(16);
  editor.destroy();
  await vi.advanceTimersByTimeAsync(16);
  await rejected;
});

it('does not accept stable geometry while a newly selected font is loading', async () => {
  let finished = false;
  const ready = prepareWordPrint(editor.view).then(() => {
    finished = true;
  });
  await vi.advanceTimersByTimeAsync(16);
  Object.defineProperty(document.fonts, 'status', { configurable: true, value: 'loading' });
  await vi.advanceTimersByTimeAsync(200);
  expect(finished).toBe(false);
  Object.defineProperty(document.fonts, 'status', { configurable: true, value: 'loaded' });
  await vi.advanceTimersByTimeAsync(80);
  await ready;
  expect(finished).toBe(true);
});

it('rejects unfinished composition without altering it', async () => {
  vi.spyOn(editor.view, 'composing', 'get').mockReturnValue(true);
  await expect(prepareWordPrint(editor.view)).rejects.toThrow('Finish entering');
  expect(editor.getText()).toBe('Before');
});

it('times out unresolved font loading without printing stale content', async () => {
  Object.defineProperty(document, 'fonts', {
    configurable: true,
    value: Object.assign(new EventTarget(), { ready: new Promise(() => {}), status: 'loading' }),
  });
  const rejected = expect(prepareWordPrint(editor.view)).rejects.toThrow('still changing');
  await vi.advanceTimersByTimeAsync(5000);
  await rejected;
  expect(editor.getText()).toBe('Before');
});
