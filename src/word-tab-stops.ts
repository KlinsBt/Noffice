import { z } from 'zod';

export const wordTabStopSchema = z.object({
  position: z.number().int().min(-31680).max(31680),
  alignment: z.enum(['left', 'center', 'right', 'decimal', 'bar', 'num', 'start', 'end']),
  leader: z.enum(['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot']),
});
export const wordTabStopsSchema = z.array(wordTabStopSchema).max(256).superRefine((stops, ctx) => {
  const positions = new Set<number>();
  for (const stop of stops) {
    if (positions.has(stop.position)) ctx.addIssue({ code: 'custom', message: 'Duplicate tab position.' });
    positions.add(stop.position);
  }
});
export type WordTabStop = z.infer<typeof wordTabStopSchema>;

export function wordTabStops(value: unknown): WordTabStop[] | null {
  if (typeof value === 'string') {
    if (value.length > 30000) return null;
    try { value = JSON.parse(value); } catch { return null; }
  }
  const parsed = wordTabStopsSchema.safeParse(value);
  return parsed.success ? parsed.data.sort((a, b) => a.position - b.position) : null;
}

export function wordDefaultTab(value: unknown): number | null {
  if (!/^\d{1,5}$/.test(String(value))) return null;
  const interval = Number(value);
  return Number.isInteger(interval) && interval >= 0 && interval <= 32767 ? interval : null;
}

/** Document metadata, not the UI language, chooses the decimal-tab separator. */
export function wordDecimalSymbol(value: unknown): string | null {
  return typeof value === 'string' && [...value].length === 1 && !/[\s\u0000-\u001f\u007f]/.test(value)
    ? value : null;
}
