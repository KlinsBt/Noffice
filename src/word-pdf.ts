import { storedFontData } from './fonts';
import { validateWordPdfSnapshot, type WordPdfSnapshot } from './word-pdf-model';
import PdfWorker from './word-pdf-worker?worker&inline';

const bundled: Record<string, () => Promise<{ default: string }>> = {
  Inter: () => import('@fontsource-variable/inter/files/inter-latin-wght-normal.woff2?url'),
  'Source Serif 4': () =>
    import('@fontsource-variable/source-serif-4/files/source-serif-4-latin-wght-normal.woff2?url'),
  'JetBrains Mono': () =>
    import('@fontsource-variable/jetbrains-mono/files/jetbrains-mono-latin-wght-normal.woff2?url'),
};
type LocalFont = { family: string; style: string; blob: () => Promise<Blob> };

async function fontsForPdf(families: string[]) {
  let local: LocalFont[] | undefined;
  const fonts: { family: string; data: ArrayBuffer }[] = [];
  for (const key of families) {
    const [family, face] = key.split('\0');
    let data: ArrayBuffer | undefined;
    if (!face && bundled[family]) {
      const url = (await bundled[family]()).default;
      const response = await fetch(url);
      if (!response.ok) throw Error(`The bundled font ${family} could not be loaded.`);
      data = await response.arrayBuffer();
    } else if (!face) data = await storedFontData(family);
    if (!data) {
      const query = (window as Window & { queryLocalFonts?: () => Promise<LocalFont[]> })
        .queryLocalFonts;
      if (!query)
        throw Error(
          `This browser cannot access ${family} for PDF export. Use Print / Save as PDF.`,
        );
      try {
        const permission = await navigator.permissions.query({
          name: 'local-fonts' as PermissionName,
        });
        if (permission.state === 'denied') throw Error('Installed font permission denied.');
        local ??= await query.call(window);
      } catch {
        throw Error(
          'Font access was not granted. Allow installed-font access and try PDF export again.',
        );
      }
      const wantedStyle = face === 'bold' ? /^bold$/i
        : face === 'italic' ? /^(italic|oblique)$/i
        : face === 'boldItalic' ? /^bold[ -]?(italic|oblique)$/i
        : /^(regular|normal|roman|book)$/i;
      const selected = local.find(
        (font) =>
          font.family.toLowerCase() === family.toLowerCase() &&
          wantedStyle.test(font.style),
      );
      if (!selected) throw Error(`The ${face || 'regular'} font ${family} is unavailable for PDF export.`);
      const blob = await selected.blob();
      if (blob.size > 10 * 1024 * 1024) throw Error('The font exceeds the PDF export size limit.');
      data = await blob.arrayBuffer();
    }
    fonts.push({ family: key, data });
  }
  return fonts;
}

/** One disposable worker per export isolates font parser failures and bounds
 * computation without blocking editing or sharing mutable document state. */
export async function exportWordPdf(snapshot: WordPdfSnapshot): Promise<Blob> {
  const families = validateWordPdfSnapshot(snapshot);
  const fonts = await fontsForPdf(families);
  return new Promise<Blob>((resolve, reject) => {
    // Vite's external worker output is not in SvelteKit's service-worker build
    // inventory. Inline its bytes in this lazy cached module for offline export.
    const worker = new PdfWorker();
    const finish = (error?: Error, bytes?: Uint8Array<ArrayBuffer>) => {
      clearTimeout(timer);
      worker.terminate();
      if (error) reject(error);
      else resolve(new Blob([bytes!], { type: 'application/pdf' }));
    };
    const timer = setTimeout(
      () => finish(Error('PDF export took too long. The document remains available for editing.')),
      15000,
    );
    worker.onerror = () =>
      finish(Error('The PDF font could not be processed. The document has not changed.'));
    worker.onmessageerror = () => finish(Error('The PDF result could not be read.'));
    worker.onmessage = (
      event: MessageEvent<{ error?: string; bytes?: Uint8Array<ArrayBuffer> }>,
    ) => {
      if (event.data.error) finish(Error(event.data.error));
      else if (!(event.data.bytes instanceof Uint8Array))
        finish(Error('The PDF result is invalid.'));
      else finish(undefined, event.data.bytes);
    };
    try {
      worker.postMessage({ snapshot, fonts });
    } catch (error) {
      finish(error instanceof Error ? error : Error('The PDF request could not be sent.'));
    }
  });
}
