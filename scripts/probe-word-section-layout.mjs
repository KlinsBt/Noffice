import { createServer } from 'vite';
import { JSDOM } from 'jsdom';
import { Document, Packer, Paragraph, Header, Footer, PageOrientation } from 'docx';
import JSZip from 'jszip';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';

const root = '.local/word-layout';
await fs.mkdir(root, { recursive: true });
const header = (text) => new Header({ children: [new Paragraph(text)] });
const input = await Packer.toBuffer(
  new Document({
    evenAndOddHeaderAndFooters: true,
    sections: [
      {
        properties: { page: { size: { width: 12240, height: 15840 } } },
        headers: { default: header('Inherited odd'), even: header('Inherited even') },
        footers: { default: new Footer({ children: [new Paragraph('Inherited footer')] }) },
        children: [new Paragraph('First section')],
      },
      {
        properties: {
          page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE } },
        },
        children: [new Paragraph('Second section')],
      },
      {
        properties: { page: { size: { width: 12240, height: 15840 } } },
        headers: { even: header('Replacement even') },
        children: [new Paragraph('Third section')],
      },
    ],
  }),
);
await fs.writeFile(`${root}/inheritance.docx`, input);
const dom = new JSDOM();
globalThis.DOMParser = dom.window.DOMParser;
globalThis.XMLSerializer = dom.window.XMLSerializer;
const server = await createServer({
  configFile: false,
  cacheDir: '.local/word-layout/vite-cache',
  server: { middlewareMode: true },
  appType: 'custom',
});
try {
  const { readDocxStructure } = await server.ssrLoadModule('/src/docx-sections.ts');
  const { resolveWordSections } = await server.ssrLoadModule('/src/word-section-layout.ts');
  const zip = await JSZip.loadAsync(input);
  const parse = (xml) => new DOMParser().parseFromString(xml, 'application/xml');
  const structure = readDocxStructure(parse(await zip.file('word/document.xml').async('string')));
  const sections = resolveWordSections({
    kind: 'word',
    html: '',
    paper: 'letter',
    margin: 'normal',
    orientation: 'portrait',
    docxStructure: structure,
  });
  const rels = parse(await zip.file('word/_rels/document.xml.rels').async('string'));
  const relationships = new Map(
    [...rels.documentElement.children].map((r) => [r.getAttribute('Id'), r.getAttribute('Target')]),
  );
  const text = async (ref) => {
    if (!ref) return null;
    const target = relationships.get(ref.relationshipId);
    if (!target || target.includes('..') || target.startsWith('/'))
      throw Error('Unexpected authored header/footer target');
    const part = parse(await zip.file('word/' + target).async('string'));
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    return {
      text: [...part.getElementsByTagNameNS(W, 'p')]
        .map((p) => [...p.getElementsByTagNameNS(W, 't')].map((t) => t.textContent).join('') + '\r')
        .join(''),
      inherited: ref.inherited,
    };
  };
  const expected = [];
  for (const section of sections) {
    const headers = {},
      footers = {};
    for (const slot of ['default', 'first', 'even']) {
      headers[slot] = await text(section.headers[slot]);
      footers[slot] = await text(section.footers[slot]);
    }
    expected.push({
      width: section.width,
      height: section.height,
      margins: section.margins,
      headers,
      footers,
    });
  }
  await fs.writeFile(
    `${root}/model-expected.json`,
    JSON.stringify(
      { sourceHash: createHash('sha256').update(input).digest('hex'), sections: expected },
      null,
      2,
    ),
  );
  console.log(
    `Created ${expected.length} native-comparison section cases from the runtime resolver.`,
  );
} finally {
  await server.close();
  dom.window.close();
}
