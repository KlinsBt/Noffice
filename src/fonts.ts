/** Font assets stay on this device. No URLs or font discovery requests are used. */
export const bundledFonts = ['Inter', 'Source Serif 4', 'JetBrains Mono'];
export const systemFonts = [
  'Arial',
  'Aptos',
  'Calibri',
  'Cambria',
  'Candara',
  'Comic Sans MS',
  'Consolas',
  'Constantia',
  'Corbel',
  'Courier New',
  'Garamond',
  'Georgia',
  'Impact',
  'Palatino Linotype',
  'Segoe UI',
  'Tahoma',
  'Times New Roman',
  'Trebuchet MS',
  'Verdana',
];
export const defaultFont = 'Inter';
export function fontStack(family = defaultFont) {
  return `"${family.replace(/["\\\r\n]/g, '')}", "Inter", sans-serif`;
}
interface StoredFont {
  family: string;
  data: ArrayBuffer;
}
const loaded = new Map<string, FontFace>();
let initialization: Promise<string[]> | undefined;
function database(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('noffice-fonts', 1);
    request.onupgradeneeded = () =>
      request.result.createObjectStore('fonts', { keyPath: 'family' });
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}
async function activate(font: StoredFont) {
  const face = await new FontFace(font.family, font.data).load();
  const previous = loaded.get(font.family);
  if (previous) document.fonts.delete(previous);
  document.fonts.add(face);
  loaded.set(font.family, face);
}
export function initializeFonts(): Promise<string[]> {
  if (!initialization)
    initialization = (async () => {
      const db = await database();
      try {
        const fonts = await new Promise<StoredFont[]>((resolve, reject) => {
          const request = db.transaction('fonts').objectStore('fonts').getAll();
          request.onsuccess = () => resolve(request.result);
          request.onerror = () => reject(request.error);
        });
        await Promise.all(fonts.map(activate));
        return [...loaded.keys()];
      } finally {
        db.close();
      }
    })().catch((error) => {
      initialization = undefined;
      throw error;
    });
  return initialization;
}
export async function importFont(file: File, family: string): Promise<string[]> {
  family = family.trim();
  if (!/^[\p{L}\p{N}][\p{L}\p{N} ._-]{0,79}$/u.test(family))
    throw new Error(
      'Use a font family name with letters, numbers, spaces, dots, hyphens or underscores (up to 80 characters).',
    );
  if ([...bundledFonts, ...systemFonts].some((name) => name.toLowerCase() === family.toLowerCase()))
    throw new Error('Choose a distinct family name for your imported font.');
  if (!/\.(ttf|otf|woff2?)$/i.test(file.name) || file.size > 10 * 1024 * 1024)
    throw new Error('Choose a TTF, OTF, WOFF or WOFF2 font smaller than 10 MB.');
  await initializeFonts();
  const font = { family, data: await file.arrayBuffer() };
  // Validate the font before saving it; commit storage before reporting success.
  const face = await new FontFace(family, font.data).load();
  const db = await database();
  try {
    await new Promise<void>((resolve, reject) => {
      const tx = db.transaction('fonts', 'readwrite');
      tx.objectStore('fonts').put(font);
      tx.oncomplete = () => resolve();
      tx.onabort = () => reject(tx.error || new Error('Font could not be saved.'));
      tx.onerror = () => reject(tx.error);
    });
  } finally {
    db.close();
  }
  const previous = loaded.get(family);
  if (previous) document.fonts.delete(previous);
  document.fonts.add(face);
  loaded.set(family, face);
  initialization = Promise.resolve([...loaded.keys()]);
  return [...loaded.keys()];
}

/** Read a copy for a user-requested local export; activation and storage are unchanged. */
export async function storedFontData(family: string): Promise<ArrayBuffer | undefined> {
  const db = await database();
  try {
    const font = await new Promise<StoredFont | undefined>((resolve, reject) => {
      const request = db.transaction('fonts').objectStore('fonts').get(family);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
    return font?.data;
  } finally {
    db.close();
  }
}
