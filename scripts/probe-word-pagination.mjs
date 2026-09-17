import { paginationFixture } from './word-pagination-fixture.mjs';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const orphan = process.argv.includes('--orphan');
const widow = process.argv.includes('--widow');
const terminal = process.argv.includes('--terminal');
const root = orphan
  ? '.local/word-orphan'
  : terminal
    ? '.local/word-terminal'
    : widow
      ? '.local/word-widow'
      : '.local/word-pagination';
await fs.mkdir(root, { recursive: true });
const bytes = await paginationFixture(
  orphan
    ? { firstLines: 12, keepSecond: false }
    : terminal
      ? { firstLines: 13, terminalBreak: true }
      : widow
        ? { firstLines: 14 }
        : {},
);
await fs.writeFile(`${root}/source.docx`, bytes);
await fs.writeFile(
  `${root}/source.json`,
  JSON.stringify({ sha256: createHash('sha256').update(bytes).digest('hex') }, null, 2),
);
