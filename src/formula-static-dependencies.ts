import type { Sheet } from './model';
import { bindFormulaReferences, coordinates, type StaticFormulaReference } from './formulas';

/** Collision-free identity shared by export dependency owners and changed cells. */
export const formulaCellIdentity = (sheetId: string, ref: string) =>
  JSON.stringify([sheetId, ref.replaceAll('$', '').toUpperCase()]);

/** Static export invalidation is deliberately distinct from observed calculation
 * edges: discovering an inactive reference must not execute it or create a cycle.
 * Index rectangles instead of expanding potentially billions of grid cells.
 */
export function staticFormulaDependents(
  sheets: Sheet[],
  changed: Iterable<string>,
  limit = 100000,
): Set<string> {
  const dirty = new Set(changed);
  const owners = new Set<string>();
  const bySheet = new Map<string, { owner: string; reference: StaticFormulaReference }[]>();
  let count = 0;
  let overflow = false;
  for (const sheet of sheets)
    for (const [ref, cell] of Object.entries(sheet.cells)) {
      if (!cell.value.startsWith('=') || cell.dataType === 'text' || cell.dataType === 'error')
        continue;
      const owner = formulaCellIdentity(sheet.id, ref);
      owners.add(owner);
      if (overflow) continue;
      const binding = bindFormulaReferences(sheets, sheet, cell.value);
      if (!binding.complete) dirty.add(owner);
      for (const reference of binding.references) {
        if (++count > limit) {
          overflow = true;
          break;
        }
        const entries = bySheet.get(reference.sheetId) || [];
        entries.push({ owner, reference });
        bySheet.set(reference.sheetId, entries);
      }
    }
  // A bounded scan also handles transitive edges, cycles and missing cells.
  // If either budget runs out, recalculate every formula rather than trust a
  // partial index. This bounds graph work, not the whole calculation engine.
  const all = () => new Set([...dirty, ...owners]);
  if (overflow) return all();
  const queue = [...dirty];
  let visits = 0;
  for (let i = 0; i < queue.length; i++) {
    const [sheetId, ref] = JSON.parse(queue[i]) as [string, string];
    const [row, col] = coordinates(ref);
    for (const {
      owner,
      reference: { from, to },
    } of bySheet.get(sheetId) || []) {
      if (++visits > limit) return all();
      if (
        !dirty.has(owner) &&
        row >= Math.min(from.row, to.row) &&
        row <= Math.max(from.row, to.row) &&
        col >= Math.min(from.col, to.col) &&
        col <= Math.max(from.col, to.col)
      ) {
        dirty.add(owner);
        queue.push(owner);
      }
    }
  }
  return dirty;
}
