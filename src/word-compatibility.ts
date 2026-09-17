import { z } from 'zod';
import { WORD_NS, val } from './word-xml';

const modeSchema = z.union([z.literal(11), z.literal(12), z.literal(14), z.literal(15)]);
export const wordCompatibilitySchema = z.object({
  version: z.literal(1),
  settingsPath: z.string().min(1).max(500).nullable(),
  declaredModes: z.array(z.string().max(100)).max(32),
  mode: modeSchema.nullable(),
  modeOrigin: z.enum(['default', 'explicit', 'unresolved']),
  wordPerfectJustification: z.boolean().nullable(),
});
export type WordCompatibility = z.infer<typeof wordCompatibilitySchema>;

/** Immutable source provenance. Unknown/conflicting modes stay unresolved;
 * retaining the package is separate from certifying its layout policy. */
export function readWordCompatibility(settings?: { path: string; document: XMLDocument }): WordCompatibility {
  const blocks = settings ? [...settings.document.documentElement.children].filter((node) =>
    node.namespaceURI === WORD_NS && node.localName === 'compat') : [];
  const children = blocks.flatMap((block) => [...block.children]).filter((node) => node.namespaceURI === WORD_NS);
  const declaredModes = children.filter((node) => node.localName === 'compatSetting' &&
    val(node, 'name') === 'compatibilityMode' && val(node, 'uri') === 'http://schemas.microsoft.com/office/word')
    .map((node) => val(node));
  if (declaredModes.length > 32 || declaredModes.some((value) => value.length > 100))
    throw Error('The Word document has excessive compatibility settings.');
  const modes = declaredModes.map((value) => /^\+?\d+$/.test(value) ? modeSchema.safeParse(Number(value)) : undefined);
  const known = modes.map((result) => result?.success ? result.data : null);
  const mode = blocks.length > 1 ? null : !declaredModes.length ? 12 :
    known.every((value) => value !== null && value === known[0]) ? known[0] : null;
  const flags = children.filter((node) => node.localName === 'wpJustification').map((node) => {
    const raw = node.getAttributeNS(WORD_NS, 'val');
    return raw === null || ['1', 'true', 'on'].includes(raw) ? true :
      ['0', 'false', 'off'].includes(raw) ? false : null;
  });
  const flag = !flags.length ? false : flags.every((value) => value !== null && value === flags[0]) ? flags[0] : null;
  return {
    version: 1,
    settingsPath: settings?.path || null,
    declaredModes,
    mode,
    modeOrigin: mode === null ? 'unresolved' : declaredModes.length ? 'explicit' : 'default',
    // Pinned native Word ignores this legacy option in mode15, even if present.
    wordPerfectJustification: mode === null ? null : mode === 15 ? false : flag,
  };
}
