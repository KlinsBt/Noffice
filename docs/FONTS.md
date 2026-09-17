# Fonts and typography

The shared font picker is available in Word, Excel, and PowerPoint. Inter, Source Serif 4, and JetBrains Mono are bundled as local variable WOFF2 files with normal and italic faces. The current bundled files cover Latin; characters outside their coverage use browser/system fallback. The service worker caches the assets with the application. No Google Fonts/CDN calls occur at runtime.

The system section lists familiar families such as Arial, Aptos, Calibri, Cambria, Georgia, and Times New Roman. These only render as requested when installed on the device. Microsoft font binaries are not included. A family name in the menu is not an installation guarantee. Missing families fall back to the browser or bundled Inter.

Use **Add font** to load a local TTF, OTF, WOFF, or WOFF2 file up to 10 MB. Enter a distinct family name. The browser validates the binary before IndexedDB storage commits. Imported fonts become available across all three modes and survive reloads, including offline. Each imported name currently represents one normal face; multiple styles/weights, metadata extraction, font removal, and cross-tab synchronization remain tasks. Keep the source font file: clearing browser data also clears imported fonts.

Word and Excel font sizes are points. Word serializes font families as quoted CSS strings, so names containing separate digits (such as Source Serif 4) remain valid. Slide font size uses the deck's 960×540 canvas units, mapped to points in its 13⅓×7½-inch export.

DOCX export preserves direct family/size, bold/italic/underline, supported script marks, paragraph spacing, and whole-document orientation. DOCX import now preserves direct run family/size; theme and style inheritance, section layout and exact text metrics are incomplete. XLSX import/export preserves cell family, size, bold, italic and underline. PPTX uses one font/style per text object; mixed-run formatting and theme inheritance remain incomplete.

Exports store requested font names. **Font files are not embedded in Office exports or native backups.** Recipients need matching installed fonts; native backups require separately adding custom font files on a new device. Font embedding/portability remains unchecked in TASKS.md.

Font assets retain their OFL-1.1 licenses; see [third-party notices](../THIRD_PARTY_NOTICES.md). Noffice application source remains MIT.
