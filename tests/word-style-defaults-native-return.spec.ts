import { readFileSync } from 'node:fs';
import { storyNativeReturnTests } from './word-story-native-return-workflow';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-style-defaults.json', 'utf8'),
) as typeof import('./fixtures/native-word-style-defaults.json');
storyNativeReturnTests(
  reference.cases.map((s) => ({
    name: s.mode + '-header',
    text: 'Created header',
    kind: 'header' as const,
    slot: 2,
    pages: s.pages,
  })),
  '.local/word-style-defaults/browser',
  '.local/word-style-defaults/native-export-report.json',
);
