import type { Cell, Sheet } from './model';
import { address, coordinates } from './formulas';
import { shiftValidation, validationBounds, type ValidationRule } from './sheet-validation';
const child = (node: Element, name: string) =>
  Array.from(node.children).find((e) => e.localName === name);
export function readValidationRanges(root: Element): NonNullable<Sheet['validationRanges']> {
  const parent = child(root, 'dataValidations');
  if (!parent) return [];
  if (parent.children.length > 10000) throw Error('This sheet has too many validation rules.');
  return Array.from(parent.children)
    .filter((e) => e.localName === 'dataValidation')
    .map((node) => {
      const rule: ValidationRule = {
        type: node.getAttribute('type') || 'none',
        formulae: ['formula1', 'formula2'].flatMap((name) => {
          const value = child(node, name);
          return value ? [value.textContent || ''] : [];
        }),
      };
      for (const name of [
        'operator',
        'errorStyle',
        'error',
        'prompt',
        'errorTitle',
        'promptTitle',
      ] as const)
        if (node.hasAttribute(name)) rule[name] = node.getAttribute(name)!;
      for (const name of ['allowBlank', 'showErrorMessage', 'showInputMessage'] as const)
        if (node.hasAttribute(name)) rule[name] = ['1', 'true'].includes(node.getAttribute(name)!);
      const ref = node.getAttribute('sqref') || '';
      if (!ref || ref.length > 32767) throw Error('Invalid validation range.');
      ref.trim().split(/\s+/).forEach(validationBounds);
      return { ref, rule };
    });
}
function writeRule(parent: Element, ref: string, rule: ValidationRule) {
  const node = parent.ownerDocument.createElementNS(parent.namespaceURI, 'dataValidation');
  node.setAttribute('sqref', ref);
  node.setAttribute('type', rule.type);
  for (const name of [
    'operator',
    'errorStyle',
    'error',
    'prompt',
    'errorTitle',
    'promptTitle',
    'allowBlank',
    'showErrorMessage',
    'showInputMessage',
  ] as const) {
    if (rule[name] !== undefined)
      node.setAttribute(
        name,
        typeof rule[name] === 'boolean' ? (rule[name] ? '1' : '0') : String(rule[name]),
      );
  }
  rule.formulae.slice(0, 2).forEach((f, i) => {
    const el = parent.ownerDocument.createElementNS(parent.namespaceURI, 'formula' + (i + 1));
    el.textContent = String(f).replace(/^=/, '');
    node.appendChild(el);
  });
  parent.appendChild(node);
}
export function writeValidationChanges(
  root: Element,
  before: Record<string, Cell>,
  after: Record<string, Cell>,
  initial: Sheet['validationRanges'] = [],
) {
  const refs = [...new Set([...Object.keys(before), ...Object.keys(after)])].filter(
    (ref) => JSON.stringify(before[ref]?.validation) !== JSON.stringify(after[ref]?.validation),
  );
  if (!refs.length && !initial?.length) return false;
  if (refs.length > 10000)
    throw Error('At most 10,000 validation cells can be exported per sheet edit.');
  // x14 rules have separate expression namespaces and precedence; never silently shadow them.
  if (
    Array.from(root.getElementsByTagNameNS('*', 'dataValidation')).some(
      (e) => e.namespaceURI !== root.namespaceURI,
    )
  )
    throw Error('Editing extended data validation rules is not supported yet.');
  let parent = child(root, 'dataValidations');
  if (!parent) {
    parent = root.ownerDocument.createElementNS(root.namespaceURI, 'dataValidations');
    const later = [
      'hyperlinks',
      'printOptions',
      'pageMargins',
      'pageSetup',
      'headerFooter',
      'rowBreaks',
      'colBreaks',
      'customProperties',
      'cellWatches',
      'ignoredErrors',
      'smartTags',
      'drawing',
      'legacyDrawing',
      'legacyDrawingHF',
      'picture',
      'oleObjects',
      'controls',
      'webPublishItems',
      'tableParts',
      'extLst',
    ];
    root.insertBefore(
      parent,
      Array.from(root.children).find((e) => later.includes(e.localName)) || null,
    );
  }
  for (const entry of initial || []) writeRule(parent, entry.ref, entry.rule);
  let work = 0;
  for (const ref of refs) {
    const [r, c] = coordinates(ref);
    for (const node of Array.from(parent.children)) {
      const ranges = (node.getAttribute('sqref') || '').trim().split(/\s+/);
      let hit = false;
      const remaining: string[] = [];
      for (const range of ranges) {
        if (++work > 5000000)
          throw Error('Validation range editing exceeded the processing limit.');
        const [top, left, bottom, right] = validationBounds(range);
        if (r < top || r > bottom || c < left || c > right) {
          remaining.push(range);
          continue;
        }
        hit = true;
        for (const [a, b, x, y] of [
          [top, left, r - 1, right],
          [r + 1, left, bottom, right],
          [r, left, r, c - 1],
          [r, c + 1, r, right],
        ])
          if (a <= x && b <= y) remaining.push(`${address(a, b)}:${address(x, y)}`);
      }
      if (!hit) continue;
      const origin = ranges[0].split(':')[0];
      for (const range of remaining) {
        const clone = node.cloneNode(true) as Element;
        clone.setAttribute('sqref', range);
        const rule: ValidationRule = {
          type: node.getAttribute('type') || 'none',
          formulae: ['formula1', 'formula2'].map((name) => child(node, name)?.textContent || ''),
        };
        const shifted = shiftValidation(rule, origin, range.split(':')[0]);
        ['formula1', 'formula2'].forEach((name, i) => {
          const el = child(clone, name);
          if (el) el.textContent = String(shifted.formulae[i]);
        });
        parent.insertBefore(clone, node);
      }
      node.remove();
    }
    const rule = after[ref]?.validation;
    if (rule && rule.type !== 'none') writeRule(parent, ref, rule);
  }
  parent.setAttribute('count', String(parent.children.length));
  if (!parent.children.length) parent.remove();
  return true;
}
