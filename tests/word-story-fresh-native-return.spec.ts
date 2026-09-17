import { readFileSync } from 'node:fs';
import { storyNativeReturnTests } from './word-story-native-return-workflow';
const reference = JSON.parse(
  readFileSync('tests/fixtures/native-word-story-fresh.json', 'utf8'),
) as typeof import('./fixtures/native-word-story-fresh.json');
storyNativeReturnTests(
  reference.rows.map((s) => ({
    name: s.name,
    text: s.text,
    kind: s.kind === 'Headers' ? 'header' : 'footer',
    slot: s.slot,
    pages: s.state.pages,
  })),
  '.local/word-story-fresh/browser',
  '.local/word-story-fresh/native-export-report.json',
);
