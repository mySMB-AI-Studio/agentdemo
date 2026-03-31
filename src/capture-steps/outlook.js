import path from 'path';

export async function captureOutlook(ctx) {
  const { context, slide, screenshotsDir } = ctx;
  const page = await context.newPage();
  const result = { status: 'done', screenshot: null, clip: null };

  try {
    await page.goto(slide.url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(5000);

    if (page.url().includes('login.microsoftonline.com')) {
      throw new Error('Auth expired — redirected to login.microsoftonline.com');
    }

    // Wait for email list or body to load
    const outlookSelectors = [
      '[role="listbox"]',
      '[data-testid="MailList"]',
      '.customScrollBar',
      '[aria-label*="Message list"]',
      '[role="main"]',
      '.ReadingPaneContainerClass',
    ];

    for (const sel of outlookSelectors) {
      try {
        await page.waitForSelector(sel, { timeout: 15000 });
        break;
      } catch { /* try next */ }
    }

    await page.waitForTimeout(2000);

    const screenshotPath = path.join(screenshotsDir, slide.screenshot_filename);
    await page.screenshot({ path: screenshotPath, fullPage: false });
    result.screenshot = screenshotPath;
  } finally {
    await page.close();
  }

  return result;
}
