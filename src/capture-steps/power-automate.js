import path from 'path';

export async function capturePowerAutomate(ctx) {
  const { context, slide, screenshotsDir } = ctx;
  const page = await context.newPage();
  const result = { status: 'done', screenshot: null, clip: null };

  try {
    await page.goto(slide.url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForLoadState('networkidle').catch(() => {});
    await page.waitForTimeout(3000);

    // Check for auth redirect
    if (page.url().includes('login.microsoftonline.com')) {
      throw new Error('Auth expired — redirected to login.microsoftonline.com');
    }

    // Wait for flows LIST (table/grid), NOT individual flow diagram
    const flowListSelectors = [
      '[role="grid"]',
      '.ms-DetailsRow',
      '[data-automationid]',
      'table',
      '.ms-List',
      '.ms-DetailsList',
      '[data-testid="flow-list"]',
      '[class*="FlowList"]',
    ];

    for (const sel of flowListSelectors) {
      try {
        await page.waitForSelector(sel, { timeout: 10000 });
        break;
      } catch { /* try next */ }
    }

    // Additional wait for rendering
    await page.waitForTimeout(2000);

    // Mask sensitive data that might be visible
    const sensitiveSelectors = [
      '[data-testid="connection-string"]',
      '[class*="credential"]',
      '[class*="secret"]',
      'input[type="password"]',
    ];
    for (const sel of sensitiveSelectors) {
      const elements = await page.$$(sel);
      for (const el of elements) {
        await el.evaluate(node => { node.style.visibility = 'hidden'; });
      }
    }

    const screenshotPath = path.join(screenshotsDir, slide.screenshot_filename);
    await page.screenshot({ path: screenshotPath, fullPage: false });
    result.screenshot = screenshotPath;
  } finally {
    await page.close();
  }

  return result;
}
