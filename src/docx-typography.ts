/** Preserve direct run font properties that Mammoth's default HTML mapping omits. */
import { wordKerningXmlValue, wordKerningStyle } from './word-kerning';
import { wordFontFeaturesXml, wordFontFeaturesStyle } from './word-font-features';
interface MammothNode {
  type?: string;
  children?: MammothNode[];
  font?: string;
  fontSize?: number;
  styleId?: string;
  styleName?: string;
}
export function docxTypography(options: { legacyRunColors?: boolean; legacyFontFeatures?: boolean } = {}) {
  const styles: { family?: string; size?: number; kerning?: number; color?: string; features?: number }[] = [];
  const readingRuns = new Map<string, { kerning?: number; color?: string; features?: number }>();
  const readingKeys = new Map<string, string>();
  const namespace = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  function retainRunStyle(run: Element, token: string) {
    let properties = [...run.children].find(
      (e) => e.namespaceURI === namespace && e.localName === 'rPr',
    );
    const kern =
      properties &&
      [...properties.children].find((e) => e.namespaceURI === namespace && e.localName === 'kern');
    const value = kern ? wordKerningXmlValue(kern.getAttributeNS(namespace, 'val')) : undefined;
    if (value === null)
      throw Error('The Word font kerning threshold is invalid or exceeds the font-size limit.');
    const colorElement = properties && [...properties.children].find((e) => e.namespaceURI === namespace && e.localName === 'color');
    const rgb = colorElement?.getAttributeNS(namespace, 'val');
    // Theme/tint/shade and automatic-color semantics need their own resolver.
    // Only a literal RGB run color is mapped into this reading-only style.
    const color = !options.legacyRunColors && rgb && /^[0-9a-f]{6}$/i.test(rgb)
      && !['themeColor', 'themeTint', 'themeShade'].some((name) => colorElement!.hasAttributeNS(namespace, name))
      ? '#' + rgb.toLowerCase() : undefined;
    const features = options.legacyFontFeatures ? undefined : wordFontFeaturesXml(properties);
    if (value === undefined && color === undefined && features === undefined) return;
    if (!properties) {
      properties = run.ownerDocument.createElementNS(namespace, 'w:rPr');
      run.prepend(properties);
    }
    const extra = { kerning: value, color, features }, key = JSON.stringify(extra);
    let id = readingKeys.get(key);
    if (!id) {
      if (readingRuns.size >= 4096) throw Error('The Word document exceeds the supported run-style limit.');
      id = `${token}RUN${readingRuns.size}`;readingKeys.set(key, id);readingRuns.set(id, extra);
    }
    const style =
      [...properties!.children].find(
        (e) => e.namespaceURI === namespace && e.localName === 'rStyle',
      ) || run.ownerDocument.createElementNS(namespace, 'w:rStyle');
    style.setAttributeNS(namespace, 'w:val', id);
    if (!style.parentElement) properties!.prepend(style);
  }
  function readingStyles(styles: Document) {
    for (const id of readingRuns.keys()) {
      const style = styles.createElementNS(namespace, 'w:style');
      style.setAttributeNS(namespace, 'w:type', 'character');
      style.setAttributeNS(namespace, 'w:styleId', id);
      const name = styles.createElementNS(namespace, 'w:name');
      name.setAttributeNS(namespace, 'w:val', id);
      style.append(name);
      styles.documentElement.append(style);
    }
    return readingRuns.size > 0;
  }
  const keys = new Map<string, number>();
  const styleMap = ['u => u'];
  function transformDocument(node: MammothNode): MammothNode {
    if (
      node.type === 'run' &&
      (node.font || node.fontSize || readingRuns.has(node.styleId || ''))
    ) {
      const style = {
        family: node.font?.slice(0, 200),
        size: node.fontSize,
        ...readingRuns.get(node.styleId || ''),
      };
      const key = JSON.stringify(style);
      let index = keys.get(key);
      if (index === undefined && styles.length < 4096) {
        index = styles.length;
        keys.set(key, index);
        styles.push(style);
        styleMap.push(`r[style-name='NofficeFont${index}'] => span.noffice-font-${index}:fresh`);
      }
      if (index !== undefined) {
        node = { ...node, styleId: `NofficeFont${index}`, styleName: `NofficeFont${index}` };
      }
    }
    return node.children ? { ...node, children: node.children.map(transformDocument) } : node;
  }
  function decorate(html: string): string {
    const dom = new DOMParser().parseFromString(html, 'text/html');
    dom.querySelectorAll('span[class^="noffice-font-"]').forEach((node) => {
      const style = styles[Number(node.className.slice('noffice-font-'.length))];
      if (!style) return;
      const element = node as HTMLElement;
      if (style.family) element.style.fontFamily = JSON.stringify(style.family);
      if (style.size && Number.isFinite(style.size))
        element.style.fontSize = `${Math.max(1, Math.min(1638, style.size))}pt`;
      if (style.color) element.style.color = style.color;
      if (style.features !== undefined) {
        element.dataset.wordFontFeatures = String(style.features);
        element.style.fontFeatureSettings = wordFontFeaturesStyle(style.features);
      }
      if (style.kerning !== undefined) {
        element.dataset.wordKerning = String(style.kerning);
        element.style.fontKerning = wordKerningStyle(style.kerning, element.style.fontSize);
      }
      element.removeAttribute('class');
    });
    return dom.body.innerHTML;
  }
  return { styleMap, transformDocument, decorate, retainRunStyle, readingStyles };
}
