import path from 'path';

export async function captureOutlook(ctx) {
  const { context, slide, screenshotsDir } = ctx;
  const page = await context.newPage();
  const result = { status: 'done', screenshot: null, clip: null };

  // Keywords to match against email subject — from slide config or defaults
  const subjectKeywords = (slide.email_subject || slide.subject_keywords || '')
    .toLowerCase()
    .split(/[,|]/)
    .map(s => s.trim())
    .filter(Boolean);

  try {
    await page.goto(slide.url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(5000);

    if (page.url().includes('login.microsoftonline.com')) {
      throw new Error('Auth expired — redirected to login.microsoftonline.com');
    }

    // Wait for email list to load
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

    // Try to open an email — prefer one matching subject keywords, fall back to first
    const emailOpened = await page.evaluate((keywords) => {
      const rows = Array.from(document.querySelectorAll(
        '[role="option"], [role="listitem"], [data-convid], ' +
        '[class*="mailListItem" i], [class*="ms-List-cell"]'
      )).filter(el => el.offsetParent !== null);

      if (rows.length === 0) return false;

      // First pass: match subject keywords
      if (keywords.length > 0) {
        for (const row of rows) {
          const t = (row.textContent || '').toLowerCase();
          if (keywords.some(kw => kw && t.includes(kw))) {
            row.click();
            return true;
          }
        }
      }

      // Second pass: click the first visible row
      rows[0].click();
      return true;
    }, subjectKeywords);

    if (emailOpened) {
      // Wait for reading pane to load
      await page.waitForTimeout(3000);
      const bodySelectors = [
        '[data-testid="messageBody"]',
        '[aria-label*="Message body"]',
        '.allowTextSelection',
        '[class*="readingPane" i]',
        '[role="main"] [class*="body" i]',
      ];
      for (const sel of bodySelectors) {
        try { await page.waitForSelector(sel, { timeout: 5000 }); break; } catch { /* next */ }
      }
      await page.waitForTimeout(1000);
    }

    const screenshotPath = path.join(screenshotsDir, slide.screenshot_filename);
    await page.screenshot({ path: screenshotPath, fullPage: false });
    result.screenshot = screenshotPath;
  } finally {
    await page.close();
  }

  return result;
}
