import { chromium } from '@playwright/test';
import fs from 'node:fs/promises';

const browser = await chromium.launch();
try {
  const page = await browser.newPage();
  const measurements = await page.evaluate(() => {
    const context = document.createElement('canvas').getContext('2d');
    return [10, 20, 30, 12288].map((pt) => {
      const p = document.createElement('div');
      p.style.cssText = `font-family:Arial;font-size:${pt}pt;line-height:normal;margin:0;padding:0;display:inline-block;white-space:pre`;
      p.append('Hg');
      const marker = document.createElement('span');
      marker.style.cssText =
        'display:inline-block;width:0;height:0;line-height:0;vertical-align:baseline';
      p.append(marker);
      document.body.append(p);
      const size = parseFloat(getComputedStyle(p).fontSize);
      context.font = `${size}px Arial`;
      const text = context.measureText('Hg');
      const result = {
        pt,
        size,
        height: p.getBoundingClientRect().height,
        baseline: marker.getBoundingClientRect().top - p.getBoundingClientRect().top,
        canvasAscent: text.fontBoundingBoxAscent,
        canvasDescent: text.fontBoundingBoxDescent,
      };
      p.remove();
      return result;
    });
  });
  await fs.writeFile(
    '.local/word-baselines/browser-font-probe.json',
    JSON.stringify({ browser: browser.version(), measurements }, null, 2),
  );
} finally {
  await browser.close();
}
