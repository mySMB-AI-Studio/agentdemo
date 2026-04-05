import path from 'path';
import { execSync } from 'child_process';
import fs from 'fs';

export async function captureM365Copilot(ctx) {
  const { context, config, slide, screenshotsDir, clipsDir } = ctx;
  const url = slide.url || config.m365_copilot_url;
  const result = { status: 'done', screenshot: null, clip: null };

  const page = await context.newPage();

  try {
    // Navigate to M365 Copilot
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(5000);

    // Check for auth redirect
    if (page.url().includes('login.microsoftonline.com')) {
      throw new Error('Auth expired — redirected to login.microsoftonline.com');
    }

    // Wait for chat input
    const chatInputSelectors = [
      'textarea[placeholder*="message"]',
      'textarea[placeholder*="Message"]',
      'textarea[placeholder*="Ask"]',
      'textarea[placeholder*="ask"]',
      'textarea[aria-label*="chat"]',
      'textarea[aria-label*="Chat"]',
      '[data-testid="chat-input"]',
      'textarea',
    ];

    let chatInput = null;
    for (const sel of chatInputSelectors) {
      try {
        chatInput = await page.waitForSelector(sel, { timeout: 10000 });
        if (chatInput) break;
      } catch { /* try next */ }
    }

    if (!chatInput) {
      throw new Error('Could not find chat input element — Timeout waiting for selector');
    }

    const prompts = slide.sample_prompts || [];
    const clips = [];

    for (let i = 0; i < prompts.length; i++) {
      const prompt = prompts[i];
      const clipBasename = `${slide.id}-prompt-${i + 1}`;

      // Start video recording if enabled
      let videoPage = null;
      if (slide.record_clip) {
        // We use Playwright's video recording on a new context if needed
        // For simplicity, we'll use screenshot-based recording approach
        // and rely on page video recording
      }

      // Type prompt with realistic speed
      await chatInput.click();
      await chatInput.fill('');
      for (const char of prompt) {
        await page.keyboard.type(char, { delay: 50 });
      }

      // Take pre-send screenshot
      await page.waitForTimeout(500);

      // Press Enter
      await page.keyboard.press('Enter');

      // Wait for response to start
      await page.waitForTimeout(3000);

      // Wait for agent to finish responding (look for response indicators)
      const responseTimeout = 45000;
      const startTime = Date.now();
      let responseComplete = false;

      if (slide.wait_for_response) {
        while (Date.now() - startTime < responseTimeout) {
          // Check for various response completion indicators
          const isTyping = await page.$('[data-testid="typing-indicator"]') ||
                          await page.$('.typing-indicator') ||
                          await page.$('[aria-label*="typing"]');

          if (!isTyping) {
            // Additional wait to ensure response is fully rendered
            await page.waitForTimeout(2000);
            const stillTyping = await page.$('[data-testid="typing-indicator"]') ||
                               await page.$('.typing-indicator');
            if (!stillTyping) {
              responseComplete = true;
              break;
            }
          }
          await page.waitForTimeout(1000);
        }

        if (!responseComplete) {
          // Save what we have as partial
          const partialScreenshot = path.join(screenshotsDir, `${slide.id}-m365-error-state.png`);
          await page.screenshot({ path: partialScreenshot, fullPage: false });
          result.screenshot = partialScreenshot;
          result.status = 'partial';

          throw new Error(`agent response timeout — Agent did not finish responding within ${responseTimeout / 1000}s`);
        }
      } else {
        await page.waitForTimeout(5000);
      }

      // Check for error responses
      const pageContent = await page.textContent('body');
      const errorPhrases = [
        'Something went wrong',
        "I'm having trouble",
        'connection failed',
        "couldn't retrieve",
      ];
      for (const phrase of errorPhrases) {
        if (pageContent && pageContent.includes(phrase)) {
          const errorScreenshot = path.join(screenshotsDir, `${slide.id}-m365-error-state.png`);
          await page.screenshot({ path: errorScreenshot, fullPage: false });
          result.screenshot = errorScreenshot;
          result.status = 'needs-review';
          console.log(`⚠ Slide ${slide.id}: Agent returned an error message — review screenshot`);
          break;
        }
      }

      // Take final screenshot
      const screenshotPath = path.join(screenshotsDir, `${slide.id}-m365-final.png`);
      await page.screenshot({ path: screenshotPath, fullPage: false });
      result.screenshot = screenshotPath;

      // Re-find chat input for next prompt
      if (i < prompts.length - 1) {
        for (const sel of chatInputSelectors) {
          try {
            chatInput = await page.waitForSelector(sel, { timeout: 5000 });
            if (chatInput) break;
          } catch { /* try next */ }
        }
      }
    }

    // If no prompts, just take a screenshot
    if (prompts.length === 0) {
      const screenshotPath = path.join(screenshotsDir, slide.screenshot_filename);
      await page.screenshot({ path: screenshotPath, fullPage: false });
      result.screenshot = screenshotPath;
    }
  } finally {
    await page.close();
  }

  return result;
}
