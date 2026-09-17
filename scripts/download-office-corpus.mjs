import fs from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';

const manifest = JSON.parse(await fs.readFile('corpus/manifest.json', 'utf8'));
const root = path.resolve('.local/office-corpus');
const hash = (algorithm, bytes) => createHash(algorithm).update(bytes).digest('hex');
const blobHash = (bytes) =>
  hash('sha1', Buffer.concat([Buffer.from(`blob ${bytes.length}\0`), bytes]));
async function download(url, limit) {
  if (new URL(url).hostname !== 'raw.githubusercontent.com')
    throw Error('Unexpected download host');
  const response = await fetch(url, { signal: AbortSignal.timeout(45000), redirect: 'error' });
  if (!response.ok) throw Error(`HTTP ${response.status}: ${url}`);
  const chunks = [];
  let size = 0;
  for await (const chunk of response.body) {
    size += chunk.length;
    if (size > limit) throw Error(`Download exceeds ${limit} bytes`);
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}
const receipts = [];
for (let start = 0; start < manifest.files.length; start += 6) {
  const results = await Promise.allSettled(
    manifest.files.slice(start, start + 6).map(async (item) => {
        if (path.basename(item.filename) !== item.filename || !['word', 'excel', 'powerpoint'].includes(item.kind)) throw Error('Invalid corpus filename or kind');
      const destination = path.join(root, 'input', item.kind, item.filename);
      let bytes;
      try {
        bytes = await fs.readFile(destination);
      } catch {}
      if (!bytes || blobHash(bytes) !== item.gitBlobSha1)
        bytes = await download(item.url, item.bytes);
      if (bytes.length !== item.bytes || blobHash(bytes) !== item.gitBlobSha1)
        throw Error(`Integrity mismatch: ${item.id}`);
      if (bytes.readUInt32LE(0) !== 0x04034b50) throw Error(`Not an OOXML ZIP: ${item.id}`);
      await fs.mkdir(path.dirname(destination), { recursive: true });
      await fs.writeFile(destination, bytes);
      return {
        id: item.id,
        bytes: bytes.length,
        sha256: hash('sha256', bytes),
        source: item.source,
      };
    }),
  );
  for (const result of results) {
    if (result.status === 'rejected') throw result.reason;
    receipts.push(result.value);
  }
  console.log(`Verified ${receipts.length}/${manifest.files.length} Office files`);
}
await fs.mkdir(path.join(root, 'licenses'), { recursive: true });
for (const name of ['LICENSE', 'NOTICE']) {
  const bytes = await download(
    `https://raw.githubusercontent.com/apache/poi/${manifest.revision}/legal/${name}`,
    1000000,
  );
  await fs.writeFile(path.join(root, 'licenses', `Apache-POI-${name}.txt`), bytes);
}
await fs.writeFile(
  path.join(root, 'receipts.json'),
  JSON.stringify(
    { revision: manifest.revision, downloadedAt: new Date().toISOString(), files: receipts },
    null,
    2,
  ) + '\n',
);
console.log(
  `Corpus ready at ${root}: ${receipts.length} verified files; upstream notices retained.`,
);
