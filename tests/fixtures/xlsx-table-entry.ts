import JSZip from 'jszip';
import { tableRenameFixture } from './xlsx-table-rename';
export async function tableEntryFixture() {
  const zip = await JSZip.loadAsync(await tableRenameFixture());
  const path = 'xl/worksheets/sheet1.xml';
  zip.file(
    path,
    (await zip.file(path)!.async('string')).replace(/<c r="[ABC][56]"[^>]*>.*?<\/c>/g, ''),
  );
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}
