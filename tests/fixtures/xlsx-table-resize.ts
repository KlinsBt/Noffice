import JSZip from 'jszip';
import { tableRenameFixture } from './xlsx-table-rename';

/** A calculated table with an existing literal and a blank destination below it. */
export async function tableResizeFixture() {
  const zip = await JSZip.loadAsync(await tableRenameFixture());
  const path = 'xl/worksheets/sheet1.xml';
  zip.file(
    path,
    (await zip.file(path)!.async('string')).replace(/<c r="C6"[^>]*>.*?<\/c>/, '<c r="C6"/>'),
  );
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
