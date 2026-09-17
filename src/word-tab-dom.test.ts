import { expect, it } from 'vitest';
import { wordTabPaintBox } from './word-tab-dom';

it('resolves a fragment-decorated tab without borrowing a different paragraph measurement', () => {
  const root = document.createElement('div');
  root.innerHTML = '<p><span data-word-tab-measured="true"><span data-word-tab></span></span></p><p data-word-tab-measured="true"><span data-word-tab></span></p>';
  const [wrapped, unmeasured] = root.querySelectorAll<HTMLElement>('[data-word-tab]');
  expect(wordTabPaintBox(wrapped)).toBe(wrapped.parentElement);
  expect(wordTabPaintBox(unmeasured)).toBe(unmeasured);
  wrapped.setAttribute('data-word-tab-measured', 'false');
  expect(wordTabPaintBox(wrapped)).toBe(wrapped);
});
