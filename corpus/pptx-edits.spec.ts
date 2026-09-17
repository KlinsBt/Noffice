import { test, expect } from '@playwright/test';
import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import manifest from './manifest.json' with { type: 'json' };
for (const filename of ['WithMaster.pptx', 'SampleShow.pptx', 'bar-chart.pptx'])
  test(`retained existing PPTX text: ${filename}`, async ({ page }) => {
    const fixture = manifest.files.find((f) => f.filename === filename)!;
    const input = await fs.readFile(`.local/office-corpus/input/powerpoint/${filename}`);
    expect(
      createHash('sha1')
        .update(Buffer.concat([Buffer.from(`blob ${input.length}\0`), input]))
        .digest('hex'),
    ).toBe(fixture.gitBlobSha1);
    await page.goto('/');
    await page.getByRole('button', { name: 'Start PowerPoint', exact: true }).waitFor();
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: filename,
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      buffer: input,
    });
    const object = page.locator('.slide-element.text').first();
    await expect(object).toBeVisible();
    const original = await object.getAttribute('aria-label');
    await object.dblclick();
    const text = original + ' — edited in Noffice';
    await page.getByRole('textbox', { name: 'Edit slide text', exact: true }).fill(text);
    await page.getByLabel('Speaker notes', { exact: true }).click();
    await page.getByRole('button', { name: 'Export', exact: true }).click();
    const pending = page.waitForEvent('download');
    await page
      .getByRole('button', { name: 'PPTX file Editable in Microsoft PowerPoint', exact: true })
      .click();
    const output = await fs.readFile((await (await pending).path())!);
    const directory = `.local/office-corpus/pptx-existing-edits/${fixture.id}`;
    await fs.mkdir(directory, { recursive: true });
    await fs.writeFile(`${directory}/edited.pptx`, output);
    const before = await JSZip.loadAsync(input),
      after = await JSZip.loadAsync(output);
    expect(Object.keys(after.files).sort()).toEqual(Object.keys(before.files).sort());
    for (const path of Object.keys(before.files))
      if (!before.files[path].dir && path !== 'ppt/slides/slide1.xml')
        expect(await after.file(path)!.async('uint8array'), path).toEqual(
          await before.file(path)!.async('uint8array'),
        );
    await page.locator('input[type=file][multiple]').setInputFiles({
      name: 'Reimport.pptx',
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      buffer: output,
    });
    await expect(page.getByRole('group', { name: text, exact: true })).toBeVisible();
    await fs.writeFile(
      `${directory}/result.json`,
      JSON.stringify(
        {
          source: fixture.source,
          inputSha256: createHash('sha256').update(input).digest('hex'),
          outputSha256: createHash('sha256').update(output).digest('hex'),
          changedPart: 'ppt/slides/slide1.xml',
          allUnrelatedPayloadsIdentical: true,
        },
        null,
        2,
      ),
    );
  });
