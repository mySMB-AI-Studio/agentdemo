import path from 'path';

export async function captureXero(ctx) {
  const { context, slide, screenshotsDir } = ctx;
  const page = await context.newPage();
  const result = { status: 'done', screenshot: null, clip: null };

  try {
    await page.goto(slide.url, { waitUntil: 'networkidle', timeout: 60000 });
    await page.waitForTimeout(3000);

    if (page.url().includes('login.xero.com')) {
      throw new Error('Xero auth required — redirected to login.xero.com');
    }

    const screenshotPath = path.join(screenshotsDir, slide.screenshot_filename);
    await page.screenshot({ path: screenshotPath, fullPage: false });
    result.screenshot = screenshotPath;
  } finally {
    await page.close();
  }

  return result;
}
