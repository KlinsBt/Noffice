import { defineConfig, devices } from '@playwright/test';
export default defineConfig({
  testDir: './corpus',
  testMatch: '**/*.spec.ts',
  timeout: 90000,
  workers: 2,
  fullyParallel: true,
  outputDir: '.local/office-corpus/traces',
  reporter: [['list'], ['json', { outputFile: '.local/office-corpus/playwright-results.json' }]],
  use: { baseURL: 'http://127.0.0.1:4183', trace: 'retain-on-failure' },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
  webServer: {
    command: 'node scripts/serve.mjs',
    url: 'http://127.0.0.1:4183',
    reuseExistingServer: !process.env.CI,
  },
});
