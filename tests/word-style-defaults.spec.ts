import { storyKeyboardTests } from './word-story-keyboard-workflow';
import { readFileSync } from 'node:fs';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-style-defaults.json', 'utf8'),
) as typeof import('./fixtures/native-word-style-defaults.json');

storyKeyboardTests(
  reference.cases.map((sample) => sample.mode),
  'word-style-defaults',
  '.local/word-style-defaults/browser',
  { kinds: ['header'], text: 'Created header' },
);
