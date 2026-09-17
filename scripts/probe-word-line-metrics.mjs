import {
  lineMetricsFixture,
  rawLineMetricsFixture,
  fontCascadeFixture,
  rawFontCascadeFixture,
} from './word-line-metrics-fixture.mjs';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
const cascade = process.argv.includes('--cascade');
const root = cascade ? '.local/word-font-cascade' : '.local/word-line-metrics';
await fs.mkdir(root, { recursive: true });
const bytes = await (process.argv.includes('--raw')
  ? cascade
    ? rawFontCascadeFixture()
    : rawLineMetricsFixture()
  : cascade
    ? fontCascadeFixture()
    : lineMetricsFixture());
await fs.writeFile(`${root}/source.docx`, bytes);
await fs.writeFile(
  `${root}/source.json`,
  JSON.stringify({ sha256: createHash('sha256').update(bytes).digest('hex') }),
);
