import path from 'path';

export async function captureSharePoint(ctx) {
  const { context, slide, screenshotsDir } = ctx;
  const page = await context.newPage();
  const result = { status: 'done', screenshot: null, clip: null };

  try {
    await page.goto(slide.url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(3000);

    // Check for auth redirect
    if (page.url().includes('login.microsoftonline.com')) {
      throw new Error('Auth expired — redirected to login.microsoftonline.com');
    }

    // Check for 403
    const statusText = await page.textContent('body').catch(() => '');
    if (statusText && (statusText.includes('Access Denied') || statusText.includes('you need access'))) {
      throw new Error('403 — Demo account does not have access to this SharePoint site');
    }

    // Wait for list/table content to load
    const tableSelectors = [
      '[role="grid"]',
      '.ms-List',
      '.ms-DetailsList',
      'table',
      '[data-automationid="ListCell"]',
      '.od-ItemContent',
    ];

    for (const sel of tableSelectors) {
      try {
        await page.waitForSelector(sel, { timeout: 10000 });
        break;
      } catch { /* try next */ }
    }

    // Additional wait for rendering
    await page.waitForTimeout(2000);

    const screenshotPath = path.join(screenshotsDir, slide.screenshot_filename);
    await page.screenshot({ path: screenshotPath, fullPage: true });
    result.screenshot = screenshotPath;
  } finally {
    await page.close();
  }

  return result;
}
