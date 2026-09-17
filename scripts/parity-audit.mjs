import { readFile, writeFile, readdir } from 'node:fs/promises';
import { resolve, relative, isAbsolute } from 'node:path';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const sha = (s) => createHash('sha256').update(s).digest('hex');
export const dimensions = [
  'import',
  'render',
  'create',
  'edit',
  'persist',
  'export',
  'interaction',
  'errors',
  'performance',
  'native',
];
const requiredFamilies =
  'S-BASE S-MODEL S-STORAGE S-TEXT S-DRAW S-PACKAGE S-CLIPBOARD S-ACCESS S-PRINT S-AUTOMATION S-NETWORK S-PERF W-STYLE W-FLOW W-LISTS W-TABLE W-OBJECT W-PAGES W-FIELDS W-REVIEW W-MERGE W-FORMS W-VIEWS W-AUTOMATION X-CALC X-ARRAY X-GRID X-FORMAT X-TABLE X-FILTER X-VALIDATE X-CONDITIONAL X-CHART X-PIVOT X-DATA X-ANALYSIS X-VIEW X-AUTOMATION X-FILES P-TEXT P-MASTER P-DRAW P-TABLE P-SLIDES P-ANIMATION P-MEDIA P-SHOW P-REVIEW P-OUTPUT P-AUTOMATION P-FILES'.split(
    ' ',
  );
export function inventory(texts) {
  const commands = new Map();
  for (const [app, text] of Object.entries(texts)) {
    const rows = new Set();
    for (const match of text.matchAll(/^- \[([ x])\] `([^`]+)`.*<!-- source-row:(\d+) -->/gm)) {
      const row = Number(match[3]),
        id = `${app}:${match[2]}`;
      if (rows.has(row)) throw Error(`Duplicate source row ${app}:${row}`);
      rows.add(row);
      if (!commands.has(id)) commands.set(id, { id, placements: [], checked: [] });
      commands.get(id).placements.push(row);
      if (match[1] === 'x') commands.get(id).checked.push(row);
    }
    const count = /\*\*(\d+) catalog placements\*\*/.exec(text);
    if (!count || Number(count[1]) !== rows.size) throw Error(`Incomplete inventory: ${app}`);
  }
  return [...commands.values()].sort((a, b) => a.id.localeCompare(b.id, 'en'));
}
export function validateRegistry(data, commands, required = []) {
  if (data.schemaVersion !== 1 || !Array.isArray(data.families) || !Array.isArray(data.contracts))
    throw Error('Invalid registry');
  const ids = new Set(),
    controls = new Set(commands.map((c) => c.id)),
    owners = new Map();
  for (const id of required)
    if (!data.families.some((f) => f.id === id)) throw Error(`Baseline family removed: ${id}`);
  for (const entry of [...data.families, ...data.contracts]) {
    if (!/^[A-Z][A-Z0-9-]+$/.test(entry.id) || ids.has(entry.id))
      throw Error(`Invalid/duplicate ID: ${entry.id}`);
    ids.add(entry.id);
    if (
      !['pending', 'specified', 'implementing', 'partial', 'verified', 'blocked'].includes(
        entry.status,
      )
    )
      throw Error(`Invalid status: ${entry.id}`);
    if (!entry.title || !Array.isArray(entry.dependencies))
      throw Error(`Incomplete entry: ${entry.id}`);
  }
  const all = [...data.families, ...data.contracts];
  for (const entry of all)
    for (const dep of entry.dependencies)
      if (!ids.has(dep)) throw Error(`Unknown dependency: ${dep}`);
  const active = new Set(),
    done = new Set();
  function visit(id) {
    if (active.has(id)) throw Error(`Dependency cycle: ${id}`);
    if (done.has(id)) return;
    active.add(id);
    all.find((e) => e.id === id).dependencies.forEach(visit);
    active.delete(id);
    done.add(id);
  }
  ids.forEach(visit);
  for (const c of data.contracts) {
    if (!data.families.some((f) => f.id === c.family) || !c.document || !Array.isArray(c.commands))
      throw Error(`Invalid contract: ${c.id}`);
    for (const control of c.commands) {
      if (!controls.has(control) || owners.has(control))
        throw Error(`Unknown/duplicate command ownership: ${control}`);
      owners.set(control, c);
    }
  }
  return owners;
}
export function validateSpecification(text) {
  for (const heading of [
    'Identity and scope',
    'Behavioral specification',
    'Model and implementation',
    'Acceptance matrix',
    'Evidence and completion',
  ])
    if (!text.includes(`## ${heading}`)) throw Error(`Incomplete contract document: ${heading}`);
  if (/Fill in|replace with actual|\[ \]/i.test(text))
    throw Error('Unfinished contract cannot be verified');
}
export async function validateProof(proof, expected, bytes) {
  if (proof.schemaVersion !== 1 || proof.contract !== expected.id || proof.result !== 'passed')
    throw Error('Invalid proof identity/result');
  for (const key of ['baselineHash', 'inventoryHash', 'implementationHash', 'contractHash'])
    if (proof[key] !== expected[key]) throw Error(`Stale proof: ${key}`);
  if (!Array.isArray(proof.checks)) throw Error('Missing acceptance checks');
  for (const dim of dimensions) {
    const checks = proof.checks.filter((c) => c.dimension === dim);
    if (!checks.length) throw Error(`Missing dimension: ${dim}`);
    for (const check of checks) {
      // A non-applicable dimension still needs a reviewed, hashed rationale artifact.
      if (
        check.exitCode !== 0 ||
        !check.command ||
        !check.artifacts?.length ||
        (check.notApplicable && !check.reason)
      )
        throw Error(`Incomplete check: ${dim}`);
      if (dim === 'native' && check.notApplicable)
        throw Error('Native Office evidence cannot be waived');
      for (const artifact of check.artifacts) {
        const content = await bytes(artifact.path);
        if (!content.length || sha(content) !== artifact.sha256)
          throw Error(`Changed/missing artifact: ${artifact.path}`);
      }
    }
  }
}
async function main() {
  const root = resolve(fileURLToPath(new URL('..', import.meta.url)));
  // Revalidate the pinned XLSX sources, not just the current Markdown counts.
  try {
    execFileSync(process.execPath, ['scripts/build-ribbon-checklists.mjs', '--check'], {
      cwd: root,
      stdio: 'pipe',
    });
  } catch (error) {
    throw Error(
      `Pinned catalog validation failed. On a fresh checkout run node scripts/download-ribbon-catalog.mjs. Inspect source drift before regenerating. ${error.stderr?.toString() || error.message}`,
    );
  }
  const safe = (path) => {
    const full = resolve(root, path),
      rel = relative(root, full);
    if (isAbsolute(path) || rel.startsWith('..') || isAbsolute(rel))
      throw Error('Evidence must stay inside the workspace');
    return full;
  };
  const read = (path) => readFile(safe(path));
  const texts = Object.fromEntries(
    await Promise.all(
      ['word', 'excel', 'powerpoint'].map(async (app) => [
        app,
        (await read(`docs/ribbon/${app.toUpperCase()}_COMMANDS.md`)).toString(),
      ]),
    ),
  );
  const commands = inventory(texts),
    data = JSON.parse(await read('docs/parity/registry.json'));
  const owners = validateRegistry(data, commands, requiredFamilies);
  const baselineHash = sha(await read('docs/parity/baseline.json'));
  const inventoryHash = sha(
    JSON.stringify(commands.map(({ id, placements }) => ({ id, placements }))),
  );
  const sourceFiles = [];
  async function walk(path) {
    for (const entry of await readdir(safe(path), { withFileTypes: true })) {
      if (entry.isSymbolicLink()) throw Error(`Symlink in implementation: ${path}/${entry.name}`);
      if (entry.isDirectory()) await walk(`${path}/${entry.name}`);
      else sourceFiles.push(`${path}/${entry.name}`);
    }
  }
  for (const dir of ['src', 'tests', 'scripts', 'corpus']) await walk(dir);
  sourceFiles.push(
    'package.json',
    'package-lock.json',
    'svelte.config.js',
    'vite.config.ts',
    'playwright.config.ts',
  );
  const implementationHash = sha(
    JSON.stringify(
      await Promise.all(sourceFiles.sort().map(async (path) => [path, sha(await read(path))])),
    ),
  );
  if (process.argv.includes('--fingerprint')) {
    console.log(JSON.stringify({ baselineHash, inventoryHash, implementationHash }, null, 2));
    return;
  }
  for (const c of data.contracts) await read(c.document);
  const failures = [];
  for (const entry of [...data.families, ...data.contracts].filter(
    (e) => e.status === 'verified',
  )) {
    if (!entry.document || !entry.proof)
      throw Error(`Verified entry lacks contract/proof: ${entry.id}`);
    validateSpecification((await read(entry.document)).toString());
    if (data.families.includes(entry)) {
      const children = data.contracts.filter((c) => c.family === entry.id);
      if (!children.length || children.some((c) => c.status !== 'verified'))
        throw Error(`Family lacks verified children: ${entry.id}`);
    }
    if (
      entry.dependencies.some(
        (d) => [...data.families, ...data.contracts].find((e) => e.id === d).status !== 'verified',
      )
    )
      throw Error(`Unverified dependency: ${entry.id}`);
    await validateProof(
      JSON.parse(await read(entry.proof)),
      {
        id: entry.id,
        baselineHash,
        inventoryHash,
        implementationHash,
        contractHash: sha(await read(entry.document)),
      },
      read,
    );
    if (entry.equivalence !== 'exact')
      failures.push(`${entry.id}: browser equivalent retains differences`);
  }
  for (const c of commands) {
    const owner = owners.get(c.id);
    if (c.checked.length && owner?.status !== 'verified')
      throw Error(`Checked ribbon row without verified contract: ${c.id}`);
    if (owner?.status !== 'verified') failures.push(c.id);
  }
  for (const f of data.families) if (f.status !== 'verified') failures.push(f.id);
  const summary = {
    placements: commands.reduce((n, c) => n + c.placements.length, 0),
    commands: commands.length,
    contracts: data.contracts.length,
    unassignedCommands: commands.filter((c) => !owners.has(c.id)).length,
    openFamilies: data.families.filter((f) => f.status !== 'verified').length,
    releaseReady: !failures.length,
    baselineHash,
    inventoryHash,
    implementationHash,
  };
  const lines = [
    '# Command coverage ledger',
    '',
    'Generated by `npm run parity:generate`. Every source command remains in this denominator; repeated placements share one command identity. Unassigned means acceptance research is still required, including any already implemented subset. This ledger does not measure a parity percentage.',
    '',
    '| Command | Source rows | Acceptance contract | Status |',
    '| --- | --- | --- | --- |',
  ];
  for (const c of commands) {
    const owner = owners.get(c.id);
    lines.push(
      `| \`${c.id}\` | ${c.placements.join(', ')} | ${owner ? `[${owner.id}](${relative(resolve(root, 'docs/parity'), safe(owner.document)).replaceAll('\\', '/')})` : 'Unassigned'} | ${owner?.status || 'pending'} |`,
    );
  }
  const generated = lines.join('\n') + '\n';
  if (process.argv.includes('--write'))
    await writeFile(safe('docs/parity/COMMAND_COVERAGE.md'), generated);
  else if ((await read('docs/parity/COMMAND_COVERAGE.md')).toString() !== generated)
    throw Error('Command coverage ledger is stale; run parity:generate');
  console.log(JSON.stringify(summary, null, 2));
  if (process.argv.includes('--strict') && failures.length) {
    console.error(
      `Parity release blocked: ${failures.length} unresolved command/family gates. No catalog command or baseline family is excluded.`,
    );
    process.exitCode = 1;
  }
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url))
  main().catch((e) => {
    console.error(e.message);
    process.exitCode = 1;
  });
