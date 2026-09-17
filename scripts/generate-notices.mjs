import { readFile, readdir, writeFile, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { createHash } from 'node:crypto';

const hash = (bytes) => createHash('sha256').update(bytes).digest('hex');
const reviewed = JSON.parse(await readFile('docs/licenses/reviewed.json', 'utf8'));
const usedReviews = new Set();
const check = process.argv.includes('--check');

const lock = JSON.parse(await readFile('package-lock.json', 'utf8'));
let output =
  'Noffice dependency notices\n==========================\n\nNoffice source: MIT. Dependencies retain their own licenses.\nThis file includes installed production dependency notices.\n\n';
const entries = [];
for (const [location, info] of Object.entries(lock.packages)) {
  if (!location || info.dev) continue;
  const manifestBytes = await readFile(path.join(location, 'package.json'));
  const manifest = JSON.parse(manifestBytes.toString('utf8'));
  if (entries.some((entry) => entry.name === manifest.name && entry.version === manifest.version))
    continue;
  const directory = await readdir(location);
  const candidates = directory.filter((name) =>
    /^(licen[sc]e|copying|notice)([^a-z]|$)/i.test(name),
  );
  const names = [];
  let text = '';
  for (const name of candidates) {
    try {
      const body = await readFile(path.join(location, name), 'utf8');
      if (body.trim()) {
        text += `\n--- ${name} ---\n${body}\n`;
        names.push(name);
      }
    } catch {
      /* License directories need separate release review. */
    }
  }
  const review = reviewed.find(
    (entry) => entry.name === manifest.name && entry.version === manifest.version,
  );
  if (review) {
    const notice = await readFile(review.file);
    if (
      review.manifestSha256 !== hash(manifestBytes) ||
      review.integrity !== info.integrity ||
      review.license !== (manifest.license || info.license) ||
      review.sha256 !== hash(notice) ||
      !notice.toString('utf8').trim()
    )
      throw new Error(`Stale notice review: ${manifest.name}@${manifest.version}`);
    usedReviews.add(review);
    text += `\n--- Reviewed upstream notice: ${review.source} ---\n${notice.toString('utf8')}\n`;
    names.push(review.file);
  }
  if (!text) {
    const readme = directory.find((name) => /^readme([^.]*)(\.|$)/i.test(name));
    if (readme) {
      const body = await readFile(path.join(location, readme), 'utf8');
      const match = body.search(/(^|\n)#+\s*licen[sc]e\b|(^|\n)licen[sc]e\s*\n/i);
      if (
        match !== -1 &&
        /Permission is hereby granted|Redistribution and use|Permission to use/.test(
          body.slice(match),
        )
      ) {
        text = `\n--- ${readme} license section ---\n${body.slice(match)}\n`;
        names.push(`${readme} (license section)`);
      }
    }
  }
  entries.push({
    name: manifest.name,
    version: manifest.version,
    license: manifest.license || info.license || 'UNKNOWN',
    noticeFiles: names,
    text,
  });
}
entries.sort((a, b) => a.name.localeCompare(b.name));
if (usedReviews.size !== reviewed.length)
  throw new Error(
    'Stale or duplicate upstream notice review; review the changed dependency graph.',
  );
for (const entry of entries)
  output += `\n${'='.repeat(72)}\n${entry.name}@${entry.version}\nLicense: ${entry.license}\n${entry.text || 'See the upstream package for its license text.\n'}`;
const inventory =
  JSON.stringify(
    entries.map(({ text, ...entry }) => entry),
    null,
    2,
  ) + '\n';
if (check) {
  if (
    (await readFile('static/THIRD_PARTY_LICENSES.txt', 'utf8')) !== output ||
    (await readFile('docs/dependency-licenses.json', 'utf8')) !== inventory
  )
    throw new Error('Generated dependency notices are stale; run npm run notices.');
} else {
  await mkdir('static', { recursive: true });
  await writeFile('static/THIRD_PARTY_LICENSES.txt', output);
  await mkdir('docs', { recursive: true });
  await writeFile('docs/dependency-licenses.json', inventory);
}
process.stdout.write(
  `${check ? 'Checked' : 'Generated'} notices for ${entries.length} production packages.\n`,
);
const missing = entries.filter((entry) => !entry.noticeFiles.length || entry.license === 'UNKNOWN');
if (missing.length)
  process.stdout.write(
    `Packages needing upstream notice review: ${missing.map((entry) => entry.name).join(', ')}\n`,
  );
if (check && missing.length) process.exitCode = 1;
