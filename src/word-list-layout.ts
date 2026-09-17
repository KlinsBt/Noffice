import { z } from 'zod';

/** Derived from the retained numbering part. Twips, not browser list defaults.
 * Unsupported definitions remain explicit so a consumer cannot certify them. */
const definition = z.object({
  numId: z.string().regex(/^[1-9]\d{0,8}$/),
  level: z.literal(0),
  start: z.number().int().min(0).max(1000000),
  format: z.enum(['decimal', 'bullet']),
  text: z.string().min(1).max(40),
  left: z.number().int().min(0).max(14400),
  hanging: z.number().int().min(0).max(14400),
  font: z.string().min(1).max(80).regex(/^[\p{L}\p{N} ._-]+$/u),
  size: z.number().min(1).max(400).multipleOf(0.5),
}).strict().superRefine((value, context) => {
  if (value.hanging > value.left ||
      (value.format === 'decimal' ? value.text !== '%1.' : value.text !== '\u2022'))
    context.addIssue({ code: 'custom', message: 'Unsupported numbering geometry or marker.' });
});
export const wordListSourceSchema = z.discriminatedUnion('status', [
  z.object({ status: z.literal('resolved'), definition }).strict(),
  z.object({ status: z.literal('unsupported'), reason: z.string().min(1).max(200) }).strict(),
]);
export const wordNumberingSchema = z.object({
  version: z.literal(1),
  paragraphs: z.array(z.object({
    source: z.string().regex(/^\d+:\d+$/).max(30),
    numbering: wordListSourceSchema,
  }).strict()).max(50000),
}).strict().superRefine((value, context) => {
  if (new Set(value.paragraphs.map(p => p.source)).size !== value.paragraphs.length)
    context.addIssue({ code: 'custom', message: 'Duplicate numbered paragraph source.' });
});
export type WordListSource = z.infer<typeof wordListSourceSchema>;
export type WordNumbering = z.infer<typeof wordNumberingSchema>;
export type WordListDefinition = z.infer<typeof definition>;
