import fs from 'fs';
import path from 'path';
import { getPlatformColor, PLATFORM_COLORS } from './config-parser.js';
import { generateAnnotationCSS, renderAnnotations } from './annotation-engine.js';

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function imageToBase64(filePath) {
  if (!fs.existsSync(filePath)) return null;
  const ext = path.extname(filePath).slice(1).toLowerCase();
  const mime = ext === 'png' ? 'image/png' : ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : 'image/png';
  const data = fs.readFileSync(filePath);
  return `data:${mime};base64,${data.toString('base64')}`;
}

export function buildSlideHTML(slide, config, demoDir, slideIndex, totalSlides) {
  const platformColor = getPlatformColor(slide.platform, config.brand_color);
  const screenshotPath = path.join(demoDir, 'screenshots', slide.screenshot_filename);
  const hasScreenshot = fs.existsSync(screenshotPath);
  const imageData = hasScreenshot ? imageToBase64(screenshotPath) : null;

  // Check for clip
  const clipFilename = slide.clip_filename;
  const clipPath = clipFilename ? path.join(demoDir, 'clips', clipFilename) : null;
  const hasClip = clipPath && fs.existsSync(clipPath);

  // Annotations
  const annotationsHTML = renderAnnotations(slide.annotations);

  // M365 Copilot chat panel
  let chatPanelHTML = '';
  if (slide.platform === 'm365-copilot' && slide.sample_prompts.length > 0) {
    const promptBubbles = slide.sample_prompts.map((p, i) => `
      <div class="chat-bubble user-bubble" style="animation-delay:${i * 2000 + 500}ms;">
        <div class="bubble-content">${escapeHtml(p)}</div>
      </div>
      <div class="chat-bubble thinking-bubble" style="animation-delay:${i * 2000 + 1500}ms;">
        <div class="thinking-dots"><span></span><span></span><span></span></div>
      </div>
      <div class="chat-bubble agent-bubble" style="animation-delay:${i * 2000 + 2500}ms;">
        <div class="bubble-content">Agent responded — see recording above</div>
      </div>
    `).join('');

    const playButton = hasClip
      ? `<button class="play-clip-btn" onclick="playClip(${slideIndex})">▶ Play Recording</button>`
      : '';

    chatPanelHTML = `
      <div class="chat-panel">
        <div class="chat-header" style="background:${platformColor};">
          <span class="chat-title">${escapeHtml(config.title)}</span>
        </div>
        <div class="chat-messages">
          ${promptBubbles}
        </div>
        ${playButton}
      </div>
    `;
  }

  // Screenshot or placeholder
  let mainContent;
  if (imageData) {
    mainContent = `
      <div class="screenshot-container">
        <img src="${imageData}" alt="Slide ${slide.id}" class="screenshot-img" />
        ${annotationsHTML}
      </div>
    `;
  } else {
    mainContent = `
      <div class="placeholder-card">
        <h3>${escapeHtml(slide.story_label)}</h3>
        <p>${escapeHtml(slide.narrative)}</p>
        <p class="placeholder-notice">Screenshot not captured — run <code>agentdemo resume</code> to complete</p>
      </div>
    `;
  }

  // Video element (relative path, not inlined)
  let videoHTML = '';
  if (hasClip) {
    const relClipPath = `clips/${clipFilename}`;
    videoHTML = `
      <div class="video-container" id="video-${slideIndex}" style="display:none;">
        <video class="clip-video" data-slide="${slideIndex}" muted>
          <source src="${relClipPath}" type="video/mp4">
        </video>
        <button class="replay-btn" onclick="replayClip(${slideIndex})" style="display:none;">↻ Replay</button>
      </div>
    `;
  }

  return `
    <div class="slide" data-slide="${slideIndex}" data-platform="${slide.platform}">
      <h2 class="story-label">${escapeHtml(slide.story_label)}</h2>
      <div class="slide-body ${slide.platform === 'm365-copilot' ? 'with-chat' : ''}">
        <div class="main-content">
          ${videoHTML}
          ${mainContent}
        </div>
        ${chatPanelHTML}
      </div>
      <div class="narrative" style="border-left-color:${platformColor};">
        <p>${escapeHtml(slide.narrative)}</p>
      </div>
    </div>
  `;
}

export function getPlatformBadge(platform, brandColor) {
  const color = getPlatformColor(platform, brandColor);
  const icons = {
    'm365-copilot': '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="7" stroke="currentColor" stroke-width="1.5" fill="none"/><path d="M5 8h6M8 5v6" stroke="currentColor" stroke-width="1.5"/></svg>',
    'sharepoint': '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="6" r="4" stroke="currentColor" stroke-width="1.5" fill="none"/><circle cx="5" cy="10" r="3" stroke="currentColor" stroke-width="1.5" fill="none"/></svg>',
    'power-automate': '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><path d="M3 12L8 2l2 5h3L8 14l-2-5H3z" stroke="currentColor" stroke-width="1.5" fill="none"/></svg>',
    'teams': '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="2" y="3" width="12" height="10" rx="2" stroke="currentColor" stroke-width="1.5" fill="none"/><path d="M6 7h4M6 9h3" stroke="currentColor" stroke-width="1.5"/></svg>',
    'outlook': '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="2" y="3" width="12" height="10" rx="1" stroke="currentColor" stroke-width="1.5" fill="none"/><path d="M2 4l6 4 6-4" stroke="currentColor" stroke-width="1.5"/></svg>',
    'xero': '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="8" cy="8" r="6" stroke="currentColor" stroke-width="1.5" fill="none"/><path d="M5 5l6 6M11 5l-6 6" stroke="currentColor" stroke-width="1.5"/></svg>',
    'custom': '<svg width="16" height="16" viewBox="0 0 16 16" fill="none"><rect x="3" y="3" width="10" height="10" rx="2" stroke="currentColor" stroke-width="1.5" fill="none"/></svg>',
  };

  const platformLabels = {
    'm365-copilot': 'M365 Copilot',
    'sharepoint': 'SharePoint',
    'power-automate': 'Power Automate',
    'teams': 'Teams',
    'outlook': 'Outlook',
    'xero': 'Xero',
    'custom': 'Custom',
  };

  const icon = icons[platform] || icons.custom;
  const label = platformLabels[platform] || platform;

  return `<span class="platform-badge" style="color:${color};border-color:${color};">${icon} ${escapeHtml(label)}</span>`;
}
