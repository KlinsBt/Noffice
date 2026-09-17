import type { Content } from './model';

/** Stable across JSON/native-backup normalization; document IDs remain part of the baseline. */
export async function contentFingerprint(content: Content): Promise<string> {
  function canonical(value: unknown): unknown {
    if (Array.isArray(value)) return value.map(canonical);
    if (value && typeof value === 'object')
      return Object.fromEntries(
        Object.entries(value)
          .filter(([, v]) => v !== undefined)
          .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
          .map(([k, v]) => [k, canonical(v)]),
      );
    return value;
  }
  // Section source metadata is derived from retained bytes, not an editable layout value.
  // Adding it to an old snapshot must not invalidate its unchanged-original identity.
  const fingerprintContent =
    content.kind === 'word'
      ? {
          ...content,
          docxStructure: undefined,
          lineSpacingVersion: undefined,
          fontMetricsVersion: undefined,
          kerningVersion: undefined,
          paragraphKerningVersion: undefined,
          styleDefaultsVersion: undefined,
          tabStopsVersion: undefined,
          runColorsVersion: undefined,
          fontFeaturesVersion: undefined,
          paragraphScriptsVersion: undefined,
          paragraphSpacingVersion: undefined,
          sectionDefaultsVersion: undefined,
          hyphenVersion: undefined,
          numbering: undefined,
          stories:
            content.stories?.parts.length || content.stories?.evenAndOddHeaders
              ? content.stories
              : undefined,
        }
      : content;
  const bytes = new TextEncoder().encode(JSON.stringify(canonical(fingerprintContent)));
  const hash = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(hash), (b) => b.toString(16).padStart(2, '0')).join('');
}

/** Reject recursive conversion inputs before they reach Mammoth/ProseMirror. */
export function assertXmlComplexity(xml: string, maxDepth = 256) {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml))
    throw new Error('Document XML cannot contain DTD or entity declarations.');
  const tokens =
    /<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<(?:"[^"]*"|'[^']*'|[^'">])*>/g;
  let depth = 0,
    count = 0;
  for (const [token] of xml.matchAll(tokens)) {
    if (token.startsWith('<!') || token.startsWith('<?')) continue;
    if (token.startsWith('</')) depth--;
    else {
      if (++count > 500000) throw new Error('Document XML exceeds the 500,000 element limit.');
      if (++depth > maxDepth)
        throw new Error(`Document XML nesting exceeds the supported ${maxDepth} levels.`);
      if (token.endsWith('/>')) depth--;
    }
  }
}
