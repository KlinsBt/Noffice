import type { SlideElement } from './model';

type Run = NonNullable<SlideElement['sourceText']>['runs'][number];
const keys = ['fontSize', 'fontFamily', 'bold', 'italic', 'underline', 'color', 'align'] as const;

/** Project the retained run styles through the same paragraph-local splice as the PPTX writer. */
export function slideTextParagraphs(element: SlideElement): Run[][] {
  const source = element.sourceText;
  if (!source) return element.text.split('\n').map((text) => [{ ...element, text }]);
  const paragraphs: Run[][] = [[]];
  for (const run of source.runs) {
    const parts = run.text.split('\n');
    for (const [index, text] of parts.entries()) {
      if (index) paragraphs.push([]);
      if (text || parts.length === 1) paragraphs[paragraphs.length - 1].push({ ...run, text });
    }
  }
  return element.text.split('\n').map((next, index) => {
    const original = paragraphs[index]?.length ? paragraphs[index] : [{ ...source.base, text: '' }];
    const previous = original.map((r) => r.text).join('');
    let prefix = 0,
      suffix = 0;
    while (prefix < previous.length && prefix < next.length && previous[prefix] === next[prefix])
      prefix++;
    while (
      suffix < previous.length - prefix &&
      suffix < next.length - prefix &&
      previous[previous.length - 1 - suffix] === next[next.length - 1 - suffix]
    )
      suffix++;
    const until = previous.length - suffix;
    const replacement = next.slice(prefix, next.length - suffix);
    let offset = 0,
      inserted = false;
    return original.map((raw) => {
      const end = offset + raw.text.length;
      const insert =
        !inserted &&
        (prefix < end || (prefix === end && (prefix === until || raw === original.at(-1))));
      const run = {
        ...raw,
        text:
          raw.text.slice(0, Math.max(0, prefix - offset)) +
          (insert ? replacement : '') +
          raw.text.slice(Math.max(0, until - offset)),
      };
      offset = end;
      if (insert) inserted = true;
      for (const key of keys) {
        const baseline =
          key === 'fontSize' ? Math.min(160, Math.max(8, source.base.fontSize)) : source.base[key];
        if (element[key] !== baseline) Object.assign(run, { [key]: element[key] });
      }
      return run;
    });
  });
}
