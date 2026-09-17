import { isWordStoryPath } from './word-stories';

/** Resolve an internal OPC URI without filesystem/network access or allowing
 * encoded separators, URI suffixes or traversal beyond the package root. */
export function wordStoryPartTarget(from: string, target: string): string {
  if (!isWordStoryPath(from) || !target || target.length > 2048 || /[%](?:2f|5c)/i.test(target))
    throw Error('Invalid Word story part target.');
  let decoded: string;
  try {
    decoded = decodeURIComponent(target);
  } catch {
    throw Error('Invalid Word story part target.');
  }
  if (/[\\:?#\u0000-\u001f\u007f]/.test(decoded))
    throw Error('Word stories must reference internal package parts.');
  const output = decoded.startsWith('/') ? [] : from.split('/').slice(0, -1);
  for (const segment of decoded.replace(/^\//, '').split('/')) {
    if (!segment) throw Error('Invalid Word story part target.');
    if (segment === '.') continue;
    if (segment === '..') {
      if (!output.length) throw Error('Word story target leaves the package.');
      output.pop();
    } else output.push(segment);
  }
  const path = output.join('/');
  if (!path || path.length > 500) throw Error('Invalid Word story part target.');
  return path;
}
