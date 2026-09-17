import { PluginKey } from '@tiptap/pm/state';
import type { DecorationSet } from '@tiptap/pm/view';

export type WordTabLayoutState = { signature: string; decorations: DecorationSet };
export const wordTabLayoutKey = new PluginKey<WordTabLayoutState>('wordTabLayout');
