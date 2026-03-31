import path from 'path';

export async function captureTeams(ctx) {
  const { context, slide, screenshotsDir } = ctx;
  const page = await context.newPage();
  const result = { status: 'done', screenshot: null, clip: null };

  try {
    await page.goto(slide.url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(5000);

    if (page.url().includes('login.microsoftonline.com')) {
      throw new Error('Auth expired — redirected to login.microsoftonline.com');
    }

    // Wait for Teams messages to load
    const messageSelectors = [
      '[data-tid="messageBodyContent"]',
      '[role="main"]',
      '.message-body',
      '.ts-message-list-container',
      '[data-testid="message-list"]',
    ];

    for (const sel of messageSelectors) {
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
