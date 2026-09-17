import source from '@fontsource-variable/inter/files/inter-latin-wght-normal.woff2?inline';
import decompress from 'woff2-encoder/decompress';

let decoded: Promise<Uint8Array> | undefined;

/** The bundled OFL font stays inside the offline export chunk. Decode this
 * trusted build asset once; no installed or imported font is redistributed. */
export async function wordDefaultFontBytes() {
  decoded ??= (async () => {
    const encoded = source.split(',')[1];
    if (!source.startsWith('data:') || !source.includes(';base64,') || !encoded)
      throw Error('The bundled Word font is unavailable.');
    const bytes = Uint8Array.from(atob(encoded), (c) => c.charCodeAt(0));
    if (bytes.length > 1024 * 1024) throw Error('The bundled Word font is too large.');
    const font = await decompress(bytes);
    if (font.length < 32 || font.length > 4 * 1024 * 1024)
      throw Error('The bundled Word font could not be decoded.');
    return font;
  })().catch((error) => {
    decoded = undefined;
    throw error;
  });
  return (await decoded).slice();
}
