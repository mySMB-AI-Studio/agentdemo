import path from 'path';

export async function captureGeneric(ctx) {
  const { context, slide, screenshotsDir } = ctx;
  const page = await context.newPage();
  const result = { status: 'done', screenshot: null, clip: null };

  try {
    await page.goto(slide.url, { waitUntil: 'networkidle', timeout: 30000 });
    await page.waitForTimeout(3000);

    if (page.url().includes('login.microsoftonline.com')) {
      throw new Error('Auth expired — redirected to login.microsoftonline.com');
    }

    const screenshotPath = path.join(screenshotsDir, slide.screenshot_filename);
    await page.screenshot({ path: screenshotPath, fullPage: true });
    result.screenshot = screenshotPath;
  } finally {
    await page.close();
  }

  return result;
}
