import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { wordStoryBuildHash } from './word-story-artifacts';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-story-fresh.json', 'utf8'),
) as typeof import('./fixtures/native-word-story-fresh.json');
const hash = (bytes: Buffer) => createHash('sha256').update(bytes).digest('hex');

for (const sample of reference.rows)
  test(`fresh ${sample.name} creates, undoes, reloads and exports independent stories`, async ({
    page,
  }) => {
    test.setTimeout(120000);
    const root = '.local/word-story-fresh/browser/' + sample.name;
    await fs.mkdir(root, { recursive: true });
    const errors: string[] = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.setViewportSize({ width: 1500, height: 1200 });
    await page.goto('/');
    await page.getByRole('button', { name: 'Start Word', exact: true }).click();
    const name = 'Fresh ' + sample.name;
    await page.getByRole('textbox', { name: 'File name', exact: true }).fill(name);
    const body = page.getByRole('textbox', { name: 'Document text', exact: true });
    await body.focus();
    await page.keyboard.insertText('line01');
    for (let i = 2; i <= 70; i++) {
      await page.keyboard.press('Shift+Enter');
      await page.keyboard.insertText('line' + String(i).padStart(2, '0'));
    }
    await expect(page.locator('.section-page')).toHaveCount(sample.state.pages);
    const controls = async () => {
      await page.getByRole('button', { name: 'Insert', exact: true }).click();
      await page.getByRole('button', { name: 'Headers and footers', exact: true }).click();
    };
    await controls();
    await page.getByRole('button', { name: 'Page options', exact: true }).click();
    await page
      .getByRole('checkbox', { name: 'Different first page', exact: true })
      .setChecked(sample.slot === 2);
    await page
      .getByRole('checkbox', { name: 'Different odd and even pages', exact: true })
      .setChecked(sample.slot === 3);
    await expect(
      page.getByRole('textbox', { name: 'Header from top (pt)', exact: true }),
    ).toHaveValue('35.4');
    await page.getByRole('button', { name: 'Apply page options', exact: true }).click();
    const kind = sample.kind === 'Headers' ? 'header' : 'footer';
    const label = ({ 1: 'Default', 2: 'First-page', 3: 'Even-page' } as const)[
      sample.slot as 1 | 2 | 3
    ];
    const open = async () => {
      await controls();
      await page
        .getByRole('button', { name: new RegExp(`^Section 1 \\u2014 ${label} ${kind}`) })
        .click();
      return page.getByRole('textbox', { name: 'Header or footer text', exact: true });
    };
    let story = await open();
    await expect(story).toHaveText('');
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    story = await open();
    await story.focus();
    await page.keyboard.type(sample.text);
    await page.keyboard.press('Control+z');
    await expect(story).not.toHaveText(sample.text);
    await page.keyboard.press('Control+y');
    await expect(story).toHaveText(sample.text);
    await page.getByRole('button', { name: 'Apply header or footer', exact: true }).click();
    const painted = page.locator('.word-page-story').filter({ hasText: sample.text });
    try {
      await expect(painted).toHaveCount(sample.slot === 1 ? sample.state.pages : 1);
    } catch (error) {
      await fs.writeFile(
        root + '/layout-failure.json',
        JSON.stringify(
          await page.evaluate(() => ({
            text: document.body.innerText,
            hosts: [...document.querySelectorAll('.word-story-measure-host')].map((h) => ({
              html: h.innerHTML,
              paragraphs: [...h.querySelectorAll('p')].map((p) => {
                const c = getComputedStyle(p);
                return {
                  height: c.height,
                  line: c.lineHeight,
                  before: c.marginTop,
                  after: c.marginBottom,
                  width: c.width,
                  bounds: p.getBoundingClientRect().toJSON(),
                  node: (
                    p as unknown as { pmViewDesc?: { node: { toJSON(): unknown } } }
                  ).pmViewDesc?.node.toJSON(),
                };
              }),
            })),
            body: document.querySelector('[aria-label="Document text"]')?.outerHTML,
          })),
          null,
          2,
        ),
      );
      throw error;
    }
    await page.getByRole('button', { name: 'Home', exact: true }).click();
    await page.getByRole('button', { name: 'Undo', exact: true }).click();
    await expect(painted).toHaveCount(0);
    await page.getByRole('button', { name: 'Redo', exact: true }).click();
    await expect(painted).toHaveCount(sample.slot === 1 ? sample.state.pages : 1);
    await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
    await page.reload();
    await page.getByRole('button', { name: 'Recent files', exact: true }).click();
    await page.getByRole('button', { name, exact: true }).click();
    await expect(await open()).toHaveText(sample.text);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    const outputs: Record<string, string> = {};
    const download = async (filename: string, label: RegExp | string) => {
      await page.getByRole('button', { name: 'Export', exact: true }).click();
      const waiting = page.waitForEvent('download');
      await page.getByRole('button', { name: label, exact: typeof label === 'string' }).click();
      await (await waiting).saveAs(root + '/' + filename);
      outputs[filename] = hash(await fs.readFile(root + '/' + filename));
    };
    await download('fresh.noffice', /^Noffice backup/);
    const backup = JSON.parse(await fs.readFile(root + '/fresh.noffice', 'utf8'));
    expect(backup.original).toBeUndefined();
    expect(backup.content.docxStructure).toBeUndefined();
    expect(backup.content.stories.parts).toHaveLength(6);
    expect(
      backup.content.stories.references.every(
        (r: { sectionId: string }) => r.sectionId === 'authored-body',
      ),
    ).toBe(true);
    await download('fresh.docx', 'DOCX file Editable in Microsoft Word');
    await download('fresh.pdf', 'PDF file');
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: 'Reimported.docx',
      buffer: await fs.readFile(root + '/fresh.docx'),
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    await expect(page.locator('.section-page')).toHaveCount(sample.state.pages);
    await expect(await open()).toHaveText(sample.text);
    await page.getByRole('button', { name: 'Cancel', exact: true }).click();
    await download('reimported.docx', 'DOCX file Editable in Microsoft Word');
    expect(outputs['reimported.docx']).toBe(outputs['fresh.docx']);
    await download('reimported.pdf', 'PDF file');
    expect(errors).toEqual([]);
    await fs.writeFile(
      root + '/browser-report.json',
      JSON.stringify(
        {
          sourceHash: reference.sourceHash,
          buildHash: await wordStoryBuildHash(),
          outputs,
          errors,
        },
        null,
        2,
      ),
    );
  });
