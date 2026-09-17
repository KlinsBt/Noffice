import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import {
  inventory,
  validateRegistry,
  validateProof,
  validateSpecification,
  dimensions,
} from './parity-audit.mjs';
test('rejects removal of a baseline family from the denominator', () =>
  assert.throws(
    () => validateRegistry(registry(), commands, ['W-REQUIRED']),
    /Baseline family removed/,
  ));
const specification = [
  'Identity and scope',
  'Behavioral specification',
  'Model and implementation',
  'Acceptance matrix',
  'Evidence and completion',
]
  .map((h) => `## ${h}\nRecorded case and artifact.`)
  .join('\n');
test('accepts complete specification structure', () => validateSpecification(specification));
test('rejects a completion claim without behavioral specification', () =>
  assert.throws(
    () => validateSpecification(specification.replace('## Behavioral specification', '## Notes')),
    /Incomplete contract/,
  ));
test('rejects unchecked acceptance tasks in a verified contract', () =>
  assert.throws(
    () => validateSpecification(specification + '\n- [ ] Native review'),
    /Unfinished/,
  ));
test('rejects unfilled contract placeholders', () =>
  assert.throws(
    () => validateSpecification(specification + '\nFill in expected behavior'),
    /Unfinished/,
  ));
const text =
  '**2 catalog placements**\n- [ ] `Save` <!-- source-row:2 -->\n- [ ] `Save` <!-- source-row:3 -->';
const commands = inventory({ word: text });
const registry = () => ({
  schemaVersion: 1,
  families: [{ id: 'W-BASE', title: 'Core', status: 'pending', dependencies: [] }],
  contracts: [
    {
      id: 'W-SAVE',
      title: 'Save',
      family: 'W-BASE',
      document: 'contract.md',
      status: 'partial',
      dependencies: [],
      commands: ['word:Save'],
    },
  ],
});
test('retains every placement while grouping duplicate command IDs', () =>
  assert.deepEqual(commands, [{ id: 'word:Save', placements: [2, 3], checked: [] }]));
test('rejects an omitted source row', () =>
  assert.throws(() => inventory({ word: text.replace('**2', '**3') }), /Incomplete/));
test('rejects duplicate source rows', () =>
  assert.throws(
    () => inventory({ word: text.replace('source-row:3', 'source-row:2') }),
    /Duplicate/,
  ));
test('assigns one full contract per command', () =>
  assert.equal(validateRegistry(registry(), commands).get('word:Save').id, 'W-SAVE'));
test('rejects references to absent commands', () => {
  const r = registry();
  r.contracts[0].commands = ['word:Missing'];
  assert.throws(() => validateRegistry(r, commands), /Unknown/);
});
test('rejects duplicate ownership', () => {
  const r = registry();
  r.contracts.push({ ...r.contracts[0], id: 'W-OTHER' });
  assert.throws(() => validateRegistry(r, commands), /duplicate command/);
});
test('rejects unknown dependencies', () => {
  const r = registry();
  r.contracts[0].dependencies = ['MISSING'];
  assert.throws(() => validateRegistry(r, commands), /Unknown dependency/);
});
test('rejects dependency cycles', () => {
  const r = registry();
  r.families[0].dependencies = ['W-SAVE'];
  r.contracts[0].dependencies = ['W-BASE'];
  assert.throws(() => validateRegistry(r, commands), /cycle/);
});
test('does not accept invented completion states', () => {
  const r = registry();
  r.contracts[0].status = 'done-ish';
  assert.throws(() => validateRegistry(r, commands), /Invalid status/);
});
const bytes = Buffer.from('independent evidence');
const hash = createHash('sha256').update(bytes).digest('hex');
const expected = {
  id: 'W-SAVE',
  baselineHash: 'baseline',
  inventoryHash: 'inventory',
  implementationHash: 'code',
  contractHash: 'contract',
};
const proof = () => ({
  schemaVersion: 1,
  ...expected,
  contract: 'W-SAVE',
  result: 'passed',
  checks: dimensions.map((dimension) => ({
    dimension,
    command: 'test-command',
    exitCode: 0,
    artifacts: [{ path: 'evidence.json', sha256: hash }],
  })),
});
test('accepts complete matching evidence', async () =>
  await validateProof(proof(), expected, async () => bytes));
for (const key of ['baselineHash', 'inventoryHash', 'implementationHash', 'contractHash'])
  test(`rejects stale ${key}`, async () => {
    const p = proof();
    p[key] = 'old';
    await assert.rejects(
      validateProof(p, expected, async () => bytes),
      /Stale/,
    );
  });
test('rejects missing acceptance dimension', async () => {
  const p = proof();
  p.checks.pop();
  await assert.rejects(
    validateProof(p, expected, async () => bytes),
    /Missing dimension/,
  );
});
test('rejects failed test execution', async () => {
  const p = proof();
  p.checks[0].exitCode = 1;
  await assert.rejects(
    validateProof(p, expected, async () => bytes),
    /Incomplete check/,
  );
});
test('rejects altered artifacts', async () =>
  await assert.rejects(
    validateProof(proof(), expected, async () => Buffer.from('changed')),
    /Changed/,
  ));
test('rejects unsupported native waiver', async () => {
  const p = proof();
  Object.assign(p.checks.at(-1), { notApplicable: true, reason: 'not available' });
  await assert.rejects(
    validateProof(p, expected, async () => bytes),
    /cannot be waived/,
  );
});
test('rejects a receipt for another feature', async () => {
  const p = proof();
  p.contract = 'W-OTHER';
  await assert.rejects(
    validateProof(p, expected, async () => bytes),
    /identity/,
  );
});
