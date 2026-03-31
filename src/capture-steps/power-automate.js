import path from 'path';

export async function capturePowerAutomate(ctx) {
  const { context, slide, screenshotsDir } = ctx;
  const page = await context.newPage();
  const result = { status: 'done', screenshot: null, clip: null };

  try {
    await page.goto(slide.url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(5000);

    // Check for auth redirect
    if (page.url().includes('login.microsoftonline.com')) {
      throw new Error('Auth expired — redirected to login.microsoftonline.com');
    }

    // Wait for flow diagram/steps panel
    const flowSelectors = [
      '[data-testid="flow-designer"]',
      '.ms-flow-designer',
      '[class*="FlowDesigner"]',
      '[class*="flow-detail"]',
      '.flow-header',
      '[data-automation-id="flow-details"]',
    ];

    for (const sel of flowSelectors) {
      try {
        await page.waitForSelector(sel, { timeout: 10000 });
        break;
      } catch { /* try next */ }
    }

    // Try to collapse expanded action cards for clean overview
    const expandedCards = await page.$$('[aria-expanded="true"]');
    for (const card of expandedCards) {
      try {
        await card.click();
        await page.waitForTimeout(300);
      } catch { /* ignore if click fails */ }
    }

    await page.waitForTimeout(2000);

    // IMPORTANT: Mask sensitive data that might be visible
    // Look for connection details, credentials, API keys
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
