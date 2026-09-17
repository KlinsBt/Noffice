import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';

const script = fileURLToPath(new URL('./generate-notices.mjs', import.meta.url));
const hash = (text) => createHash('sha256').update(text).digest('hex');
async function fixture() {
  const root = await mkdtemp(path.join(tmpdir(), 'noffice-notices-'));
  const put = async (name, text) => {
    await mkdir(path.dirname(path.join(root, name)), { recursive: true });
    await writeFile(path.join(root, name), text);
  };
  const manifest = JSON.stringify({ name: 'sample', version: '1.0.0', license: 'MIT' });
  await put('node_modules/sample/package.json', manifest);
  await put(
    'package-lock.json',
    JSON.stringify({
      packages: { 'node_modules/sample': { version: '1.0.0', integrity: 'test-integrity' } },
    }),
  );
  await put('docs/licenses/reviewed.json', '[]');
  return {
    root,
    put,
    manifest,
    run: (...args) =>
      spawnSync(process.execPath, [script, ...args], { cwd: root, encoding: 'utf8' }),
  };
}

test('read-only release check rejects absent notices and stale generated output', async () => {
  const f = await fixture();
  assert.equal(f.run().status, 0);
  assert.equal(f.run('--check').status, 1);
  await f.put('node_modules/sample/LICENSE', 'Sample upstream notice');
  assert.equal(f.run().status, 0);
  assert.equal(f.run('--check').status, 0);
  await f.put('static/THIRD_PARTY_LICENSES.txt', 'stale');
  const result = f.run('--check');
  assert.equal(result.status, 1);
  assert.match(result.stderr, /notices are stale/);
  assert.equal(
    await readFile(path.join(f.root, 'static/THIRD_PARTY_LICENSES.txt'), 'utf8'),
    'stale',
  );
});

test('empty license files and license directories do not count as bundled notices', async () => {
  const f = await fixture();
  await f.put('node_modules/sample/LICENSE', ' ');
  await mkdir(path.join(f.root, 'node_modules/sample/NOTICE'));
  assert.equal(f.run().status, 0);
  assert.equal(f.run('--check').status, 1);
  const inventory = JSON.parse(
    await readFile(path.join(f.root, 'docs/dependency-licenses.json'), 'utf8'),
  );
  assert.deepEqual(inventory[0].noticeFiles, []);
});

test('upstream reviews bind notice bytes, package identity, manifest and lock integrity', async () => {
  const f = await fixture();
  const notice = 'Exact upstream license';
  await f.put('docs/licenses/sample.txt', notice);
  const review = {
    name: 'sample',
    version: '1.0.0',
    license: 'MIT',
    file: 'docs/licenses/sample.txt',
    source: 'https://example.invalid/pinned/LICENSE',
    sha256: hash(notice),
    manifestSha256: hash(f.manifest),
    integrity: 'test-integrity',
  };
  await f.put('docs/licenses/reviewed.json', JSON.stringify([review]));
  assert.equal(f.run().status, 0);
  assert.equal(f.run('--check').status, 0);
  for (const change of [
    { sha256: 'wrong' },
    { manifestSha256: 'wrong' },
    { integrity: 'wrong' },
    { version: '2.0.0' },
  ]) {
    await f.put('docs/licenses/reviewed.json', JSON.stringify([{ ...review, ...change }]));
    const result = f.run();
    assert.equal(result.status, 1);
    assert.match(result.stderr, /Stale/);
  }
});
