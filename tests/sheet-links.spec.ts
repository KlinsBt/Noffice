import { expect, test } from '@playwright/test';
import fs from 'node:fs/promises';
import JSZip from 'jszip';
import { sheetLinksFixture } from './fixtures/xlsx-links';
test('Excel links support editing, internal navigation, protected cells, undo, reload and retained export', async ({
  page,
}) => {
  const input = Buffer.from(await sheetLinksFixture());
  await fs.mkdir('.local/xlsx-validation', { recursive: true });
  await fs.writeFile('.local/xlsx-validation/links-source.xlsx', input);
  await page.goto('/');
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Link workbook.xlsx',
    mimeType: 'application/octet-stream',
    buffer: input,
  });
  const cell = (ref: string) => page.getByRole('gridcell', { name: ref, exact: true });
  const openEditor = () => page.getByRole('button', { name: 'Hyperlink', exact: true }).click();
  const apply = () => page.getByRole('button', { name: 'Apply link', exact: true }).click();
  await cell('A1').click();
  await page.keyboard.press('Control+k');
  await expect(page.getByRole('dialog', { name: 'Cell hyperlink' })).toBeVisible();
  await page.getByLabel('Link address', { exact: true }).fill('javascript:alert(1)');
  await apply();
  await expect(page.getByRole('alert')).toContainText('Use https://');
  await page.getByLabel('Link address', { exact: true }).fill('https://example.com/edited?x=1&y=2');
  await page.getByLabel('Link text', { exact: true }).fill('Edited label');
  await page.getByLabel('Link ScreenTip', { exact: true }).fill('Edited tip');
  await apply();
  await expect(cell('A1')).toHaveText('Edited label');
  await expect(cell('A1').locator('[title]')).toHaveAttribute('title', 'Edited tip');
  await page.getByRole('button', { name: 'Undo', exact: true }).click();
  await expect(cell('A1')).toHaveText('Original label');
  await page.getByRole('button', { name: 'Redo', exact: true }).click();
  await page
    .context()
    .route('https://example.com/edited?*', (route) =>
      route.fulfill({ contentType: 'text/html', body: '<title>Local link target</title>' }),
    );
  const opened = page.waitForEvent('popup');
  await page.getByRole('button', { name: 'Open link', exact: true }).click();
  const popup = await opened;
  await expect(popup).toHaveURL('https://example.com/edited?x=1&y=2');
  expect(await popup.evaluate(() => window.opener)).toBeNull();
  await popup.close();
  await cell('A2').click();
  await openEditor();
  await page.getByLabel('Link address', { exact: true }).fill('mailto:hello@example.com');
  await apply();
  await expect(cell('A2')).toHaveText('42');
  await cell('A3').click();
  await openEditor();
  await expect(page.getByLabel('Link text', { exact: true })).toBeDisabled();
  await page.getByLabel('Link address', { exact: true }).fill('#A2');
  await apply();
  await expect(cell('A3')).toHaveText('84');
  await cell('C1').click();
  await page.getByRole('button', { name: 'Remove links', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Open link', exact: true })).toBeDisabled();
  await cell('C2').click();
  await expect(page.getByRole('button', { name: 'Open link', exact: true })).toBeEnabled();
  await cell('D1').click();
  await page.getByRole('button', { name: 'Open link', exact: true }).click();
  await expect(
    page.getByText('Use https://, http://, mailto: or #Sheet!A1 for a workbook location.', {
      exact: true,
    }),
  ).toBeVisible();
  await cell('B1').click();
  await page.getByRole('button', { name: 'Open link', exact: true }).click();
  await expect(cell('B3')).toHaveText('Destination');
  await expect(cell('B3')).toHaveAttribute('aria-selected', 'true');
  await page.getByRole('button', { name: 'Protected', exact: true }).click();
  await cell('A1').click();
  await openEditor();
  await page.getByLabel('Link address', { exact: true }).fill('#Links!A1');
  await apply();
  await expect(page.getByRole('alert')).toContainText('protected');
  await page.getByRole('button', { name: 'Cancel', exact: true }).click();
  await page.getByRole('button', { name: 'Links', exact: true }).click();
  await expect(page.getByText('Saved on this device', { exact: true })).toBeVisible();
  await page.reload();
  await page.getByRole('button', { name: 'Recent files', exact: true }).click();
  await page.getByRole('button', { name: 'Link workbook', exact: true }).click();
  await expect(cell('A1')).toHaveText('Edited label');
  await expect(cell('A3')).toHaveText('84');
  await page.getByRole('button', { name: 'Export', exact: true }).click();
  const pending = page.waitForEvent('download');
  await page
    .getByRole('button', { name: 'XLSX file Editable in Microsoft Excel', exact: true })
    .click();
  const output = await fs.readFile((await (await pending).path())!);
  await fs.writeFile('.local/xlsx-validation/links-edited.xlsx', output);
  const before = await JSZip.loadAsync(input),
    after = await JSZip.loadAsync(output);
  for (const path of Object.keys(before.files))
    if (
      !before.files[path].dir &&
      ![
        'xl/styles.xml',
        'xl/worksheets/sheet1.xml',
        'xl/worksheets/_rels/sheet1.xml.rels',
        'xl/workbook.xml',
      ].includes(path)
    )
      expect(await after.file(path)!.async('uint8array'), path).toEqual(
        await before.file(path)!.async('uint8array'),
      );
  await page.locator('input[type=file][multiple]').setInputFiles({
    name: 'Link result.xlsx',
    mimeType: 'application/octet-stream',
    buffer: output,
  });
  await expect(cell('A1')).toHaveText('Edited label');
  await expect(cell('A3')).toHaveText('84');
  await cell('A2').click();
  await openEditor();
  await expect(page.getByLabel('Link address', { exact: true })).toHaveValue(
    'mailto:hello@example.com',
  );
  await page.screenshot({ path: 'test-results/excel-links.png' });
});
