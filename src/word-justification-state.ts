import { PluginKey } from '@tiptap/pm/state';
import type { DecorationSet } from '@tiptap/pm/view';

export const wordJustificationKey = new PluginKey<{
  signature: string;
  decorations: DecorationSet;
}>('wordJustification');
