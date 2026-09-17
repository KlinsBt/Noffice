import { test, expect, type Page } from '@playwright/test';

async function inspectRibbon(page: Page, name: string, width: number) {
  const ribbon = page.getByRole('toolbar').first();
  await expect(ribbon).toBeVisible();
  const problems = await ribbon.evaluate((element) => {
    const errors: string[] = [];
    const bounds = element.getBoundingClientRect();
    if (element.scrollWidth > element.clientWidth + 1) errors.push('Horizontal ribbon overflow');
    if (document.documentElement.scrollWidth > window.innerWidth) errors.push('Page overflow');
    for (const group of element.querySelectorAll<HTMLElement>('[role="group"]')) {
      if (group.scrollWidth > group.clientWidth + 1)
        errors.push(`${group.ariaLabel}: clipped group`);
      const home =
        element
          .closest('.editor-body')
          ?.querySelector('.ribbon-tabs .selected')
          ?.textContent?.trim() === 'Home';
      if (bounds.width >= 2500 || (home && bounds.width >= 1900)) {
        const centers = [...group.querySelectorAll<HTMLElement>('button, select, input, .color-tool')]
          .filter((control) => !control.matches('input[type="color"]'))
          .map((control) => {
            const box = control.getBoundingClientRect();
            // Check row alignment: a centered native checkbox has a different
            // top edge from the taller numeric fields alongside it.
            return box.top + box.height / 2;
          });
        if (Math.max(...centers) - Math.min(...centers) > 2)
          errors.push(`${group.ariaLabel}: unnecessary stacking in a wide ribbon`);
      }
    }
    for (const control of element.querySelectorAll<HTMLElement>('button, select, input')) {
      const box = control.getBoundingClientRect();
      if (box.left < bounds.left || box.right > bounds.right) {
        errors.push(`${control.ariaLabel || control.textContent}: outside ribbon`);
      }
      if (control.matches('select, input[type="number"]') && box.width < 60) {
        errors.push(`${control.ariaLabel}: squeezed input`);
      }
      if (
        control.matches('button, select') &&
        parseFloat(getComputedStyle(control).fontSize) < 12
      ) {
        errors.push(`${control.ariaLabel || control.textContent}: tiny text`);
      }
    }
    return errors;
  });
  expect(problems, `${name} at ${width}px`).toEqual([]);
  // Scrolling must bring the final controls fully into view on small screens.
  const last = ribbon.locator('button, select, input').last();
  await last.scrollIntoViewIfNeeded();
  await expect(last).toBeInViewport({ ratio: 1 });
  await ribbon.evaluate((element) => (element.scrollTop = 0));
  await page.mouse.move(0, 0);
  await page.screenshot({ path: `test-results/ribbons/${name}-${width}.png` });
}

for (const mode of ['Word', 'Excel', 'PowerPoint']) {
  test(`${mode} ribbon keeps controls readable and reachable across window sizes`, async ({
    page,
  }) => {
    await page.goto('/');
    await page.getByRole('button', { name: `Start ${mode}`, exact: true }).click();
    for (const width of [2560, 1920, 1440, 1280, 1024, 768, 390, 320]) {
      await page.setViewportSize({ width, height: 900 });
      for (const tab of mode === 'Word'
        ? ['Home', 'Insert', 'Layout', 'Review', 'View']
        : ['Home']) {
        await page.locator('.ribbon-tabs').getByRole('button', { name: tab, exact: true }).click();
        await inspectRibbon(page, `${mode}-${tab}`, width);
      }
    }
    if (mode === 'Word') {
      await page
        .locator('.ribbon-tabs')
        .getByRole('button', { name: 'Insert', exact: true })
        .click();
      await page.getByRole('button', { name: 'Table', exact: true }).click();
      await expect(page.getByRole('group', { name: 'Table editing' })).toBeVisible();
      await inspectRibbon(page, 'Word-Table', 320);
      await page.setViewportSize({ width: 1440, height: 900 });
      await inspectRibbon(page, 'Word-Table', 1440);
    }
  });
}
