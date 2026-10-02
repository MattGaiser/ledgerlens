import { defineConfig } from '@playwright/test';
export default defineConfig({
  testDir: './tests/browser', fullyParallel: false, workers: 1, timeout: 30000,
  reporter: [['list'], ['json', { outputFile: 'artifacts/validation/browser-results.json' }]],
  use: { channel: 'msedge', headless: true, viewport: { width: 430, height: 980 }, screenshot: 'only-on-failure', trace: 'retain-on-failure' },
});
