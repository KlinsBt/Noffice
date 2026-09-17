import { readFile, readdir } from 'node:fs/promises';
import assert from 'node:assert/strict';
import path from 'node:path';
const pkg = JSON.parse(await readFile('package.json', 'utf8'));
assert(pkg.devDependencies['@sveltejs/adapter-static'], 'Static adapter is required');
assert(
  !pkg.dependencies.react && !pkg.dependencies['react-dom'],
  'React must not be an application dependency',
);
const html = await readFile('build/index.html', 'utf8');
assert(html.includes('_app/immutable'), 'Missing SvelteKit static app shell');
assert(
  (await readFile('src/routes/+layout.ts', 'utf8')).includes('ssr = false'),
  'Client-only rendering is required',
);
assert((await readFile('build/service-worker.js', 'utf8')).length > 0, 'Missing offline worker');
async function walk(root) {
  const paths = [];
  for (const entry of await readdir(root, { withFileTypes: true })) {
    const target = path.join(root, entry.name);
    if (entry.isDirectory()) paths.push(...(await walk(target)));
    else paths.push(target);
  }
  return paths;
}
const routes = await walk('src/routes');
assert(
  !routes.some((file) => /\+server\.|\.server\./.test(file)),
  'Runtime server routes are not allowed',
);
const source = await walk('src');
assert(!source.some((file) => /\.[jt]sx$/.test(file)), 'Editors must use native Svelte components');
const assets = await walk('build');
// Concurrent SvelteKit processes can replace generated version metadata while
// a build is running. Such a shell loads successfully but cannot start its app.
const shellIds = new Set(html.match(/\b__sveltekit_[a-z0-9]+\b/g) ?? []);
assert.equal(shellIds.size, 1, 'Static shell must define one SvelteKit startup identifier');
for (const file of assets.filter((file) => file.endsWith('.js'))) {
  const code = await readFile(file, 'utf8');
  for (const id of code.match(/\bglobalThis\.__sveltekit_[a-z0-9]+\b/g) ?? []) {
    assert(shellIds.has(id.slice('globalThis.'.length)),
      `SvelteKit startup identifier mismatch in ${file}; run tests and build sequentially`);
  }
}
assert(
  !assets.some((file) => /(?:^|[\\/])server(?:[\\/]|\.)/.test(file)),
  'Do not deploy a runtime server',
);
process.stdout.write(
  `Static architecture verified: ${assets.length} deployable files, native Svelte editors, no server routes.\n`,
);
