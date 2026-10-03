import { test, expect } from '@playwright/test';
import fs from 'node:fs';
const endpoint = JSON.parse(
  fs.readFileSync('.runtime/endpoint.json', 'utf8').replace(/^\uFEFF/, ''),
);
test.beforeEach(async ({ page }) => {
  await page.goto(endpoint.BaseUrl + '/#session=' + endpoint.Token);
  await expect(page.locator('#revenue-value')).toHaveText('$281.7B');
});
test('financial overview matches sourced values and switches company', async ({ page }) => {
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await expect(page.locator('#revenue-change')).toHaveText('+14.9% YoY');
  await page.getByRole('button', { name: 'AAPL', exact: true }).click();
  await expect(page.locator('#company-name')).toHaveText('Apple');
  await expect(page.locator('#revenue-value')).toHaveText('$416.2B');
  await page.getByRole('button', { name: 'NVDA', exact: true }).click();
  await expect(page.locator('#revenue-value')).toHaveText('$130.5B');
  expect(errors).toEqual([]);
});
test('source audit is keyboard dismissible and links only to SEC', async ({ page }) => {
  await page.getByRole('button', { name: /Operating income:.*Inspect source/ }).click();
  await expect(page.getByRole('dialog')).toBeVisible();
  await expect(page.locator('#source-details')).toContainText('us-gaap:OperatingIncomeLoss');
  await expect(page.getByRole('link', { name: /Open the original/ })).toHaveAttribute(
    'href',
    /^https:\/\/www\.sec\.gov\/Archives\//,
  );
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog')).not.toBeVisible();
});

test('workspace navigation works with keyboard focus and Enter', async ({ page }) => {
  const research = page.getByRole('button', { name: 'Research', exact: true });
  await research.focus();
  await page.keyboard.press('Enter');
  await expect(page.getByRole('button', { name: /Research this company/ })).toBeVisible();
  await page.keyboard.press('Tab');
  expect(await page.evaluate(() => document.activeElement?.tagName)).toBe('BUTTON');
});
test('calculated research produces traceable claims without calling AI', async ({ page }) => {
  await page.getByRole('button', { name: 'Research', exact: true }).click();
  await page.getByLabel('Use OpenAI').uncheck();
  await page.getByRole('button', { name: /Research this company/ }).click();
  await expect(page.locator('.answer-card')).toBeVisible();
  await expect(page.locator('.answer-meta')).toContainText('Calculated analysis');
  await expect(page.locator('.claim')).toHaveCount(3);
  await page.locator('.citation').first().click();
  await expect(page.getByRole('dialog')).toBeVisible();
});
test('failure scenarios are explicit and recover', async ({ page }) => {
  await page.getByRole('button', { name: 'Health', exact: true }).click();
  try {
    await page.getByRole('button', { name: 'Offline', exact: true }).click();
    await expect(page.locator('#connection-state')).toContainText('Offline cache');
    await expect(page.locator('[data-mode=offline]')).toHaveClass(/selected/);
    await page.getByRole('button', { name: 'Test notifications' }).click();
    await expect(page.locator('#event-list')).toContainText('Notification channel test received');
  } finally {
    await page.getByRole('button', { name: 'Online', exact: true }).click();
  }
  await expect(page.locator('#connection-state')).toContainText('Connected');
});
test('browser mode cannot pretend to write into Excel', async ({ page }) => {
  await expect(page.getByRole('button', { name: /Insert into Excel/ })).toBeDisabled();
  await page.getByRole('button', { name: 'Refresh', exact: true }).click();
  await expect(page.getByRole('button', { name: /Preview workbook changes/ })).toBeDisabled();
  await expect(page.locator('#refresh-hint')).toContainText('Open the LedgerLens pane in Excel');
});
test('320px pane has no horizontal overflow on any tab', async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 850 });
  for (const tab of ['Overview', 'Research', 'Refresh', 'Health']) {
    await page.getByRole('button', { name: tab, exact: true }).click();
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(
      true,
    );
  }
});
test('untrusted model output remains text, never executable markup', async ({ page }) => {
  await page.route('**/api/research', (route) =>
    route.fulfill({
      json: {
        headline: '<img src=x onerror="window.compromised=true">',
        summary: '<script>window.compromised=true</script>',
        claims: [{ text: '<iframe src="https://bad.example">', sourceIds: [] }],
        caveats: ['Untrusted content'],
        sources: [],
        provider: 'Test provider',
        model: 'Fixture',
        isAiGenerated: true,
        generatedAt: new Date().toISOString(),
        cached: false,
      },
    }),
  );
  await page.getByRole('button', { name: 'Research', exact: true }).click();
  await page.getByRole('button', { name: /Research this company/ }).click();
  await expect(page.locator('.answer-card h2')).toContainText('<img');
  expect(await page.evaluate(() => Boolean(window.compromised))).toBe(false);
  await expect(
    page.locator('.answer-card img,.answer-card iframe,.answer-card script'),
  ).toHaveCount(0);
});
