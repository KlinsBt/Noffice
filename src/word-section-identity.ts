import { z } from 'zod';

export const retainedWordSectionId = z.string().regex(/^word\/document\.xml#section:\d+$/);
export const insertedWordSectionId = z.string().regex(/^authored-section:[a-f0-9]{32}$/);
export const wordSectionIdSchema = z.union([
  retainedWordSectionId,
  insertedWordSectionId,
  z.literal('authored-body'),
]);
export const wordSectionStartSchema = z.enum(['nextPage', 'continuous', 'evenPage', 'oddPage']);
export type WordSectionStart = z.infer<typeof wordSectionStartSchema>;
