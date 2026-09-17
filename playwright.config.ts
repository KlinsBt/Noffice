import { defineConfig, devices } from '@playwright/test';
export default defineConfig({
  testDir: './tests',
  fullyParallel: false,
  // Office import/export fixtures are CPU and memory intensive in Chromium.
  workers: 4,
  timeout: 45000,
  use: { baseURL: 'http://127.0.0.1:4183', trace: 'retain-on-failure' },
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
  webServer: {
    command: 'node scripts/serve.mjs',
    url: 'http://127.0.0.1:4183',
    reuseExistingServer: !process.env.CI,
  },
});
