import { spawnSync } from 'node:child_process';
import { readFile, writeFile } from 'node:fs/promises';
import assert from 'node:assert/strict';
const output = '.local/ribbon-generator-test';
const run = (check = false) =>
  spawnSync(
    process.execPath,
    ['scripts/build-ribbon-checklists.mjs', ...(check ? ['--check'] : [])],
    { env: { ...process.env, NOFFICE_RIBBON_OUTPUT: output }, encoding: 'utf8' },
  );
assert.equal(run().status, 0);
const path = `${output}/WORD_COMMANDS.md`,
  original = await readFile(path, 'utf8');
const marked = original.replace(/^- \[ \]/m, '- [x]');
assert.notEqual(marked, original);
await writeFile(path, marked);
assert.equal(run().status, 0);
assert.equal(await readFile(path, 'utf8'), marked);
assert.equal(run(true).status, 0);
await writeFile(path, marked.replace('FileNewDefault', 'IncorrectCommand'));
assert.notEqual(run(true).status, 0);
assert.equal(run().status, 0);
assert.equal(await readFile(path, 'utf8'), marked);
console.log(
  'Catalog regeneration retains checked rows and rejects command drift in isolated test output.',
);
