import { drawingChild as child, type presentationPaints } from './drawing-colors';
import type { SlideElement } from './model';

/** Resolve Latin run properties individually, retaining explicit false values. */
export function presentationText(
  shape: Element,
  layout: Element | undefined,
  masterShape: Element | undefined,
  master: Document | null,
  theme: Document | null,
  presentation: Document,
  paints: ReturnType<typeof presentationPaints>,
  width: number,
): NonNullable<SlideElement['sourceText']> {
  const body = child(shape, 'txBody');
  const paragraphs = body ? [...body.children].filter((p) => p.localName === 'p') : [];
  const ph = shape.getElementsByTagNameNS('*', 'ph')[0];
  const type =
    ph?.getAttribute('type') ||
    layout?.getElementsByTagNameNS('*', 'ph')[0]?.getAttribute('type') ||
    'obj';
  const category = ['title', 'ctrTitle'].includes(type)
    ? 'titleStyle'
    : type === 'body' || (type === 'obj' && ph)
      ? 'bodyStyle'
      : 'otherStyle';
  const themeFont = (name: string | undefined) => {
    if (!name || !/^\+(mj|mn)-lt$/.test(name)) return name;
    const group = theme?.getElementsByTagNameNS(
      '*',
      name.startsWith('+mj') ? 'majorFont' : 'minorFont',
    )[0];
    return child(group, 'latin')?.getAttribute('typeface') || undefined;
  };
  const runs: NonNullable<SlideElement['sourceText']>['runs'] = [];
  for (const [index, paragraph] of paragraphs.entries()) {
    const pPr = child(paragraph, 'pPr');
    const level = Number(pPr?.getAttribute('lvl') || 0);
    if (!Number.isInteger(level) || level < 0 || level > 8)
      throw Error('Invalid PowerPoint paragraph level.');
    const levelName = `lvl${level + 1}pPr`;
    const levels = [shape, layout, masterShape].map((s) =>
      child(child(child(s, 'txBody'), 'lstStyle'), levelName),
    );
    const masterLevel = child(
      child(master?.getElementsByTagNameNS('*', 'txStyles')[0], category),
      levelName,
    );
    const defaultLevel = child(
      presentation.getElementsByTagNameNS('*', 'defaultTextStyle')[0],
      levelName,
    );
    const paragraphSources = [pPr, ...levels, masterLevel, defaultLevel];
    const alignment = paragraphSources.find((s) => s?.hasAttribute('algn'))?.getAttribute('algn');
    const align = alignment === 'ctr' ? 'center' : alignment === 'r' ? 'right' : 'left';
    const nodes = [...paragraph.children].filter((s) => ['r', 'fld', 'br'].includes(s.localName));
    if (!nodes.length) nodes.push(paragraph);
    for (const node of nodes) {
      const direct = child(node, node === paragraph ? 'endParaRPr' : 'rPr');
      const sources = [direct, ...paragraphSources.map((s) => child(s, 'defRPr'))];
      const attr = (name: string) => sources.find((s) => s?.hasAttribute(name))?.getAttribute(name);
      const face = sources.map((s) => child(s, 'latin')?.getAttribute('typeface')).find(Boolean);
      const fontRef = child(child(shape, 'style'), 'fontRef');
      const fontFamily = themeFont(
        face || (fontRef?.getAttribute('idx') === 'major' ? '+mj-lt' : '+mn-lt'),
      );
      const colorNode =
        sources.map((s) => child(s, 'solidFill')?.firstElementChild).find(Boolean) ||
        fontRef?.firstElementChild;
      const sz = Number(attr('sz') || 2400);
      if (!Number.isFinite(sz) || sz <= 0) throw Error('Invalid PowerPoint run font size.');
      runs.push({
        text:
          node === paragraph
            ? ''
            : node.localName === 'br'
              ? '\n'
              : child(node, 't')?.textContent || '',
        fontSize: (sz * 127 * 960) / width,
        fontFamily,
        bold: ['1', 'true'].includes(attr('b') || ''),
        italic: ['1', 'true'].includes(attr('i') || ''),
        underline: !!attr('u') && attr('u') !== 'none',
        color: paints.color(colorNode)?.color || '#263d34',
        align,
      });
    }
    if (index < paragraphs.length - 1) runs.push({ ...runs[runs.length - 1], text: '\n' });
  }
  if (!runs.length)
    runs.push({
      text: '',
      fontSize: (24 * 12700 * 960) / width,
      bold: false,
      italic: false,
      underline: false,
      color: '#263d34',
      align: 'left',
    });
  return { text: runs.map((r) => r.text).join(''), runs, base: { ...runs[0] } };
}
