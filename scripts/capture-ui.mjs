import { chromium } from '@playwright/test';
import fs from 'node:fs';
const endpoint = JSON.parse(
  fs
    .readFileSync(new URL('../.runtime/endpoint.json', import.meta.url), 'utf8')
    .replace(/^\uFEFF/, ''),
);
const browser = await chromium.launch({ channel: 'msedge', headless: true });
try {
  const page = await browser.newPage({
    viewport: { width: 430, height: 1060 },
    deviceScaleFactor: 1,
  });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(endpoint.BaseUrl + '/#session=' + endpoint.Token);
  await page.getByText('$281.7B', { exact: true }).waitFor();
  await page.screenshot({ path: 'artifacts/test-results/overview.png', fullPage: true });
  console.log(
    JSON.stringify({
      errors,
      heading: await page.locator('h1:visible').innerText(),
      horizontalOverflow: await page.evaluate(
        () => document.documentElement.scrollWidth > innerWidth,
      ),
    }),
  );
} finally {
  await browser.close();
}
