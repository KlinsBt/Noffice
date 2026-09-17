/** Every direct drawing object is retained, including objects the editor cannot render. */
export function pptxStackNodes(tree: Element): Map<string, Element> {
  const result = new Map<string, Element>();
  Array.from(tree.children).forEach((node, i) => {
    if (['nvGrpSpPr', 'grpSpPr', 'extLst'].includes(node.localName)) return;
    const id = node.getElementsByTagNameNS('*', 'cNvPr')[0]?.getAttribute('id');
    const key = id ? `source:${id}` : `opaque:${i}`;
    if (result.has(key)) throw Error('The slide has duplicate drawing identities.');
    result.set(key, node);
  });
  return result;
}
