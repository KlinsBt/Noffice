import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';

// Bind native comparisons to the exact static artifact exercised by the browser.
// Include paths as well as bytes so moved or removed assets invalidate evidence.
export async function wordStoryBuildHash(): Promise<string> {
  const files: string[] = [];
  async function visit(relative: string) {
    for (const entry of await fs.readdir(path.join('build', relative), { withFileTypes: true })) {
      const name = relative ? `${relative}/${entry.name}` : entry.name;
      if (entry.isDirectory()) await visit(name);
      else if (entry.isFile()) files.push(name);
      else throw new Error(`Unexpected build entry: ${name}`);
    }
  }
  await visit('');
  if (!files.length) throw new Error('Missing static build');
  const manifest = await Promise.all(
    files.sort().map(async (name) => {
      const hash = createHash('sha256')
        .update(await fs.readFile(path.join('build', name)))
        .digest('hex');
      return `${name}\0${hash}\n`;
    }),
  );
  return createHash('sha256').update(manifest.join('')).digest('hex');
}
