import fs from 'node:fs';
import { chromium } from '@playwright/test';
const svg = fs.readFileSync('web/assets/icon.svg', 'utf8');
const browser = await chromium.launch({ channel: 'msedge', headless: true });
try {
  for (const size of [32, 64, 80]) {
    const page = await browser.newPage({
      viewport: { width: size, height: size },
      deviceScaleFactor: 1,
    });
    await page.setContent(
      `<style>html,body{margin:0;width:100%;height:100%}svg{width:100%;height:100%}</style>${svg}`,
    );
    await page.screenshot({ path: `web/assets/icon-${size}.png`, omitBackground: true });
    await page.close();
  }
} finally {
  await browser.close();
}
