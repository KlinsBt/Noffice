import { z } from 'zod';
import { wordNumberingSchema } from './word-list-layout';
import { docxStructureSchema } from './docx-sections';
import { wordSectionStateSchema } from './word-section-breaks';
import { emptyWordParagraph } from './word-defaults';
import { wordStoriesSchema } from './word-stories';

export type AppKind = 'word' | 'excel' | 'powerpoint';
export const appInfo = {
  word: {
    name: 'Word',
    type: 'Document',
    plural: 'Documents',
    extension: 'docx',
    color: '#526eac',
    letter: 'W',
  },
  excel: {
    name: 'Excel',
    type: 'Spreadsheet',
    plural: 'Spreadsheets',
    extension: 'xlsx',
    color: '#418369',
    letter: 'X',
  },
  powerpoint: {
    name: 'PowerPoint',
    type: 'Presentation',
    plural: 'Presentations',
    extension: 'pptx',
    color: '#c37a56',
    letter: 'P',
  },
};

const validationSchema = z.object({
  type: z.string(),
  operator: z.string().optional(),
  formulae: z.array(z.union([z.string(), z.number()])),
  allowBlank: z.boolean().optional(),
  showErrorMessage: z.boolean().optional(),
  errorStyle: z.string().optional(),
  error: z.string().optional(),
  prompt: z.string().optional(),
  errorTitle: z.string().optional(),
  promptTitle: z.string().optional(),
  showInputMessage: z.boolean().optional(),
});
const cellSchema = z.object({
  value: z.string().max(32767),
  dataType: z.enum(['text', 'number', 'boolean', 'error', 'date']).optional(),
  cachedValue: z.union([z.string(), z.number(), z.boolean()]).optional(),
  numFmt: z.string().max(2048).optional(),
  date1904: z.boolean().optional(),
  wrap: z.boolean().optional(),
  vertical: z.enum(['top', 'middle', 'bottom']).optional(),
  locked: z.boolean().optional(),
  borders: z
    .object({
      top: z.string().optional(),
      bottom: z.string().optional(),
      left: z.string().optional(),
      right: z.string().optional(),
    })
    .optional(),
  hyperlink: z.string().optional(),
  hyperlinkTooltip: z.string().max(32767).optional(),
  validation: validationSchema.optional(),
  note: z.string().optional(),
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  fontFamily: z.string().max(200).optional(),
  fontSize: z.number().min(1).max(409).optional(),
  color: z.string().optional(),
  fill: z.string().optional(),
  format: z.enum(['general', 'number', 'currency', 'percent']).optional(),
  align: z.enum(['left', 'center', 'right']).optional(),
});
const sheetSchema = z.object({
  id: z.string(),
  name: z.string().min(1).max(31),
  cells: z.record(z.string().regex(/^[A-Z]{1,3}[1-9]\d{0,6}$/), cellSchema),
  arrayFormulas: z
    .array(z.object({ anchor: z.string().max(12), ref: z.string().max(25) }))
    .max(10000)
    .optional(),
  rows: z.number().int().min(1).max(10000),
  cols: z.number().int().min(1).max(256),
  sourcePath: z.string().optional(),
  date1904: z.boolean().optional(),
  state: z.enum(['visible', 'hidden', 'veryHidden']).optional(),
  columnWidths: z.record(z.string(), z.number().positive().max(2000)).optional(),
  defaultRowHeight: z.number().positive().max(2000).optional(),
  defaultColumnWidth: z.number().positive().max(2000).optional(),
  rowHeights: z.record(z.string(), z.number().positive().max(2000)).optional(),
  hiddenRows: z.array(z.number().int()).optional(),
  filterMode: z.boolean().optional(),
  autoFilters: z
    .array(
      z.object({
        ref: z.string().max(40),
        sourcePath: z.string().optional(),
        columns: z
          .array(
            z.object({
              col: z.number().int().min(0).max(16383),
              values: z.array(z.string()).max(10000).optional(),
              blank: z.boolean().optional(),
              custom: z
                .array(z.object({ operator: z.string(), value: z.string() }))
                .max(2)
                .optional(),
              and: z.boolean().optional(),
              unsupported: z.boolean().optional(),
            }),
          )
          .max(256),
      }),
    )
    .max(1000)
    .optional(),
  hiddenColumns: z.array(z.number().int()).optional(),
  merges: z.array(z.string()).optional(),
  validationRanges: z
    .array(z.object({ ref: z.string().max(32767), rule: validationSchema }))
    .max(10000)
    .optional(),
  gridLines: z.boolean().optional(),
  protected: z.boolean().optional(),
  frozenRows: z.number().int().min(0).max(10000).optional(),
  frozenColumns: z.number().int().min(0).max(256).optional(),
  images: z
    .array(
      z.object({
        src: z.string(),
        x: z.number().finite(),
        y: z.number().finite(),
        w: z.number().positive(),
        h: z.number().positive(),
      }),
    )
    .max(500)
    .optional(),
  definedNames: z.record(z.string(), z.string()).optional(),
  nameDefinitions: z
    .array(
      z.object({
        name: z.string().max(255),
        formula: z.string().max(8192),
        scope: z.enum(['workbook', 'worksheet']),
        referenceOrigin: z.literal('ooxml-a1').optional(),
      }),
    )
    .max(10000)
    .optional(),
  tableTheme: z
    .array(z.string().regex(/^#[a-f\d]{6}$/i))
    .length(6)
    .optional(),
  tableEditingBlocked: z.string().optional(),
  tables: z
    .array(
      z.object({
        name: z.string().max(255),
        ref: z.string().max(40),
        columns: z.array(z.string().max(32767)).max(256),
        calculatedColumns: z.array(z.string().max(8192).nullable()).max(256).optional(),
        totalLabels: z.array(z.string().max(32767).nullable()).max(256).optional(),
        totalFormulas: z.array(z.string().max(8192).nullable()).max(256).optional(),
        totalsActivationBlocked: z.string().optional(),
        calculationBlocked: z.string().optional(),
        headerRows: z.number().int().min(0).max(1),
        totalRows: z.number().int().min(0).max(1),
        sourcePath: z.string().optional(),
        resizeBlocked: z.string().optional(),
        style: z
          .object({
            name: z.string().max(255),
            rowStripes: z.boolean(),
            columnStripes: z.boolean(),
            firstColumn: z.boolean(),
            lastColumn: z.boolean(),
          })
          .optional(),
      }),
    )
    .max(1000)
    .optional(),
});
const slideRunSchema = z.object({
  text: z.string().max(100000),
  fontSize: z.number().finite().positive().max(100000),
  fontFamily: z.string().max(200).optional(),
  bold: z.boolean(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  color: z.string(),
  align: z.enum(['left', 'center', 'right']),
});
const elementSchema = z.object({
  id: z.string(),
  type: z.enum(['text', 'rect', 'ellipse', 'image']),
  x: z.number().finite(),
  y: z.number().finite(),
  w: z.number().positive().max(2000),
  h: z.number().positive().max(2000),
  rotation: z.number().finite().min(0).max(360).optional(),
  flipH: z.boolean().optional(),
  flipV: z.boolean().optional(),
  text: z.string().max(100000),
  color: z.string(),
  fill: z.string(),
  fillOpacity: z.number().min(0).max(1).optional(),
  outline: z
    .object({
      color: z.string().regex(/^(#[a-f\d]{6}|transparent)$/i),
      width: z.number().finite().min(0).max(1000),
      opacity: z.number().min(0).max(1),
      dash: z.enum([
        'solid',
        'dot',
        'dash',
        'lgDash',
        'dashDot',
        'lgDashDot',
        'lgDashDotDot',
        'sysDash',
        'sysDot',
        'sysDashDot',
        'sysDashDotDot',
      ]),
    })
    .optional(),
  fontSize: z.number().min(8).max(160),
  bold: z.boolean(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  fontFamily: z.string().max(200).optional(),
  align: z.enum(['left', 'center', 'right']),
  src: z.string().optional(),
  sourceShapeId: z.string().optional(),
  sourceStackKey: z.string().max(500).optional(),
  sourceGroupIds: z.array(z.string().max(100)).max(100).optional(),
  sourceGroupTransform: z
    .object({
      dimensions: z
        .tuple([z.number().finite().positive(), z.number().finite().positive()])
        .optional(),
      matrix: z.tuple([
        z.number().finite(),
        z.number().finite(),
        z.number().finite(),
        z.number().finite(),
        z.number().finite(),
        z.number().finite(),
      ]),
      local: z.object({
        x: z.number().finite(),
        y: z.number().finite(),
        w: z.number().finite().positive(),
        h: z.number().finite().positive(),
        rotation: z.number().finite(),
      }),
      projected: z.boolean(),
    })
    .optional(),
  sourcePlaceholder: z
    .object({
      index: z.string().max(100),
      type: z.string().max(100),
      layoutShapeId: z.string().max(100).optional(),
      masterShapeId: z.string().max(100).optional(),
      geometry: z.enum(['slide', 'layout', 'master', 'default']),
    })
    .optional(),
  sourceText: z
    .object({
      text: z.string().max(100000),
      runs: z.array(slideRunSchema).max(100000),
      base: slideRunSchema,
    })
    .optional(),
});
const slideSchema = z.object({
  id: z.string(),
  sourcePath: z.string().max(500).optional(),
  sourceLayoutPath: z.string().max(500).optional(),
  sourceMasterPath: z.string().max(500).optional(),
  sourceThemePath: z.string().max(500).optional(),
  stackOrder: z.array(z.string().max(500)).max(10000).optional(),
  background: z.string(),
  backgroundOpacity: z.number().min(0).max(1).optional(),
  notes: z.string(),
  elements: z.array(elementSchema).max(500),
});
export const contentSchema = z.discriminatedUnion('kind', [
  z.object({
    kind: z.literal('word'),
    docxStructure: docxStructureSchema.optional(),
    initialEditSession: z.string().regex(/^[a-f\d]{32}$/).optional(),
    stories: wordStoriesSchema.optional(),
    lineSpacingVersion: z.literal(1).optional(),
    fontMetricsVersion: z.literal(1).optional(),
    kerningVersion: z.literal(1).optional(),
    paragraphKerningVersion: z.literal(1).optional(),
    styleDefaultsVersion: z.literal(1).optional(),
    tabStopsVersion: z.literal(1).optional(),
    runColorsVersion: z.literal(1).optional(),
    fontFeaturesVersion: z.literal(1).optional(),
    paragraphScriptsVersion: z.literal(1).optional(),
    paragraphSpacingVersion: z.literal(1).optional(),
    sectionDefaultsVersion: z.literal(1).optional(),
    hyphenVersion: z.literal(1).optional(),
    numbering: wordNumberingSchema.optional(),
    sectionState: wordSectionStateSchema.optional(),
    html: z.string().max(10000000),
    paper: z.enum(['a4', 'letter']),
    margin: z.enum(['normal', 'narrow', 'wide']),
    orientation: z.enum(['portrait', 'landscape']).optional(),
    pageOverrides: z
      .object({
        paper: z.literal(true).optional(),
        margin: z.literal(true).optional(),
        orientation: z.literal(true).optional(),
      })
      .optional(),
  }),
  z.object({
    kind: z.literal('excel'),
    sheets: z.array(sheetSchema).min(1).max(100),
    // A structural checkpoint is a retained XLSX package, separate from the user's original.
    // Keeping it in content makes undo and IndexedDB revisions atomic with the edited grid.
    xlsxStructureBase: z.string().max(70000000).optional(),
  }),
  z.object({
    kind: z.literal('powerpoint'),
    textModelVersion: z.literal(1).optional(),
    groupModelVersion: z.literal(1).optional(),
    aspectRatio: z.number().finite().min(0.1).max(10).optional(),
    slides: z.array(slideSchema).min(1).max(500),
  }),
]);
export type Content = z.infer<typeof contentSchema>;
export type WordContent = Extract<Content, { kind: 'word' }>;
export type WorkbookContent = Extract<Content, { kind: 'excel' }>;
export type DeckContent = Extract<Content, { kind: 'powerpoint' }>;
export type Cell = z.infer<typeof cellSchema>;
export type Sheet = z.infer<typeof sheetSchema>;
export type Slide = z.infer<typeof slideSchema>;
export type SlideElement = z.infer<typeof elementSchema>;
export interface OfficeFile {
  id: string;
  name: string;
  kind: AppKind;
  content: Content;
  createdAt: number;
  updatedAt: number;
  revision: number;
  favorite: boolean;
  trashed: boolean;
  warnings: string[];
  original?: { name: string; data: ArrayBuffer; contentFingerprint?: string };
}
export const uid = () => crypto.randomUUID();
export function textElement(text: string, overrides: Partial<SlideElement> = {}): SlideElement {
  return {
    id: uid(),
    type: 'text',
    x: 80,
    y: 100,
    w: 800,
    h: 120,
    text,
    color: '#263d34',
    fill: 'transparent',
    fontSize: 48,
    bold: false,
    align: 'left',
    ...overrides,
  };
}
export function newSlide(layout: 'title' | 'content' | 'blank' = 'title'): Slide {
  return {
    id: uid(),
    background: '#f5f3ea',
    notes: '',
    elements:
      layout === 'blank'
        ? []
        : layout === 'title'
          ? [
              textElement('A new perspective.', { y: 165, fontSize: 62, bold: true, h: 150 }),
              textElement('Add your subtitle here', {
                y: 340,
                h: 60,
                fontSize: 24,
                color: '#6c7e73',
              }),
            ]
          : [
              textElement('Your next big idea', { y: 65, fontSize: 44, bold: true, h: 70 }),
              textElement(
                'Start with what matters.\n\nAdd the details that bring your story to life.',
                { y: 180, fontSize: 28, h: 280 },
              ),
            ],
  };
}
export function newFile(kind: AppKind, name?: string, content?: Content): OfficeFile {
  return {
    id: uid(),
    name: name || `Untitled ${appInfo[kind].type.toLowerCase()}`,
    kind,
    content:
      content ||
      (kind === 'word'
        ? { kind, html: emptyWordParagraph, paper: 'a4', margin: 'normal', tabStopsVersion: 1, runColorsVersion: 1, fontFeaturesVersion: 1, paragraphScriptsVersion: 1, paragraphSpacingVersion: 1 }
        : kind === 'excel'
          ? { kind, sheets: [{ id: uid(), name: 'Sheet 1', cells: {}, rows: 100, cols: 26 }] }
          : { kind, slides: [newSlide()] }),
    createdAt: Date.now(),
    updatedAt: Date.now(),
    revision: 0,
    favorite: false,
    trashed: false,
    warnings: [],
  };
}
export function samples(): OfficeFile[] {
  const proposal = newFile('word', 'A fresh start', {
    kind: 'word',
    paper: 'a4',
    margin: 'normal',
    html: '<p>THE NOFFICE NOTEBOOK</p><h1>A little space for<br>big ideas.</h1><p>Every good thing starts with a first draft. This is yours.</p><h2>Make yourself at home</h2><p>Write a proposal, plan your next project, or finally put that idea into words. Your workspace brings documents, spreadsheets, and presentations together, right here in your browser.</p><blockquote><p>Less getting in the way. More getting things done.</p></blockquote><h2>A few things to try</h2><ul><li><p>Make this document your own with the formatting toolbar.</p></li><li><p>Open Excel to turn a few numbers into a plan.</p></li><li><p>Build a story, one slide at a time, in PowerPoint.</p></li></ul><p>Your files save automatically on this device. Export a copy whenever you need one.</p>',
  });
  const cells: Record<string, Cell> = {};
  const data = [
    ['Project budget', '', '', ''],
    ['Item', 'Quantity', 'Unit cost', 'Total'],
    ['Research', '12', '85', '=B3*C3'],
    ['Design', '24', '95', '=B4*C4'],
    ['Development', '40', '110', '=B5*C5'],
    ['Testing', '16', '80', '=B6*C6'],
    ['', '', '', ''],
    ['Total', '', '', '=SUM(D3:D6)'],
  ];
  data.forEach((row, r) =>
    row.forEach((v, c) => {
      if (v)
        cells[`${String.fromCharCode(65 + c)}${r + 1}`] = {
          value: v,
          bold: r < 2 || r === 7,
          fill: r === 1 ? '#e4eee7' : undefined,
          format: c > 1 && r > 1 ? 'currency' : undefined,
        };
    }),
  );
  const budget = newFile('excel', 'Project budget', {
    kind: 'excel',
    sheets: [{ id: uid(), name: 'Budget', cells, rows: 100, cols: 26 }],
  });
  const deck = newFile('powerpoint', 'The next chapter', {
    kind: 'powerpoint',
    slides: [
      newSlide(),
      {
        ...newSlide('content'),
        elements: [
          textElement('Good things take shape.', { y: 90, bold: true }),
          textElement(
            '01   Find the possibility\n\n02   Make a little progress\n\n03   Share it with the world',
            { y: 220, fontSize: 28, h: 240 },
          ),
        ],
      },
      {
        ...newSlide('blank'),
        background: '#263d34',
        elements: [
          textElement('Let’s make it happen.', {
            y: 200,
            color: '#f5f3ea',
            fontSize: 64,
            bold: true,
          }),
        ],
      },
    ],
  });
  return [proposal, budget, deck];
}
