import { describe, expect, it } from 'vitest';
import { docxTypography } from './docx-typography';

describe('Word imported run-size bounds', () => {
  it.each([1, 1.5, 6, 7.5, 10, 160, 160.5, 200, 1638])('preserves %s points', (size) => {
    const typography = docxTypography();
    typography.transformDocument({ type: 'run', fontSize: size });
    const doc = new DOMParser().parseFromString(
      typography.decorate('<span class="noffice-font-0">a</span>'),
      'text/html',
    );
    expect(doc.querySelector('span')!.style.fontSize).toBe(`${size}pt`);
  });
});
