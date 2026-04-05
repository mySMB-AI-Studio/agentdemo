import fs from 'fs';
import path from 'path';
import { parseConfig, getDemoDir, getPlatformColor, PLATFORM_COLORS } from './config-parser.js';
import { generateAnnotationCSS } from './annotation-engine.js';
import { buildSlideHTML, getPlatformBadge } from './slide-builder.js';

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export async function runGenerate(opts) {
  const configPath = opts.config;
  const config = parseConfig(configPath);
  const demoDir = getDemoDir(configPath);
  const outputDir = path.join(demoDir, 'output');

  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const totalSlides = config.slides.length;

  // Build all slide HTML
  const slidesHTML = config.slides.map((slide, i) =>
    buildSlideHTML(slide, config, demoDir, i, totalSlides)
  ).join('\n');

  // Build platform badge data for JS
  const platformBadges = config.slides.map(s => ({
    platform: s.platform,
    color: getPlatformColor(s.platform, config.brand_color),
    badge: getPlatformBadge(s.platform, config.brand_color),
  }));

  // Agent icon
  let iconHTML = '';
  if (config.agent_icon) {
    const iconPath = path.resolve(demoDir, config.agent_icon);
    if (fs.existsSync(iconPath)) {
      const ext = path.extname(iconPath).slice(1);
      const mime = ext === 'svg' ? 'image/svg+xml' : `image/${ext}`;
      const data = fs.readFileSync(iconPath).toString('base64');
      iconHTML = `<img src="data:${mime};base64,${data}" alt="icon" class="agent-icon" />`;
    }
  }
  if (!iconHTML) {
    iconHTML = `<div class="agent-icon-placeholder" style="background:${config.brand_color};">${escapeHtml(config.title.charAt(0))}</div>`;
  }

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${escapeHtml(config.title)} — Interactive Demo</title>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    background: #f5f5f5;
    color: #1a1a1a;
    overflow: hidden;
    height: 100vh;
  }

  /* ── Top bar ── */
  .top-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 12px 24px;
    background: #fff;
    border-bottom: 1px solid #e0e0e0;
    z-index: 100;
    position: relative;
  }
  .top-left {
    display: flex;
    align-items: center;
    gap: 12px;
  }
  .agent-icon {
    width: 32px;
    height: 32px;
    border-radius: 8px;
  }
  .agent-icon-placeholder {
    width: 32px;
    height: 32px;
    border-radius: 8px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #fff;
    font-weight: 700;
    font-size: 16px;
  }
  .agent-title {
    font-size: 18px;
    font-weight: 700;
  }
  .slide-counter {
    font-size: 14px;
    color: #666;
    font-weight: 500;
  }

  /* ── Progress bar ── */
  .progress-bar {
    height: 3px;
    background: #e0e0e0;
    position: relative;
  }
  .progress-fill {
    height: 100%;
    background: ${config.brand_color};
    transition: width 350ms ease;
  }

  /* ── Slides container ── */
  .slides-wrapper {
    position: relative;
    height: calc(100vh - 110px);
    overflow: hidden;
  }
  .slides-track {
    display: flex;
    transition: transform 350ms ease-in-out;
    height: 100%;
  }
  .slide {
    min-width: 100%;
    padding: 32px 48px;
    display: flex;
    flex-direction: column;
    align-items: center;
    overflow-y: auto;
    height: 100%;
  }

  /* ── Story label ── */
  .story-label {
    font-size: 28px;
    font-weight: 700;
    text-align: center;
    margin-bottom: 20px;
    color: #1a1a1a;
  }

  /* ── Slide body ── */
  .slide-body {
    display: flex;
    gap: 24px;
    align-items: flex-start;
    justify-content: center;
    width: 100%;
    max-width: 1200px;
  }
  .slide-body.with-chat .main-content {
    flex: 1;
    max-width: calc(100% - 344px);
  }

  .main-content {
    position: relative;
    width: 100%;
  }

  /* ── Screenshot ── */
  .screenshot-container {
    position: relative;
    display: inline-block;
    width: 100%;
  }
  .screenshot-img {
    max-height: 65vh;
    max-width: 100%;
    border-radius: 8px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.1);
    display: block;
    margin: 0 auto;
  }

  /* ── Placeholder ── */
  .placeholder-card {
    background: #fff;
    border: 2px dashed #ccc;
    border-radius: 12px;
    padding: 48px;
    text-align: center;
    max-width: 600px;
    margin: 0 auto;
  }
  .placeholder-card h3 { margin-bottom: 12px; }
  .placeholder-notice {
    margin-top: 16px;
    font-size: 13px;
    color: #999;
  }
  .placeholder-notice code {
    background: #f0f0f0;
    padding: 2px 6px;
    border-radius: 4px;
    font-size: 12px;
  }

  /* ── Narrative ── */
  .narrative {
    border-left: 4px solid #00C9A7;
    padding: 16px 24px;
    background: #fff;
    border-radius: 0 8px 8px 0;
    margin-top: 24px;
    max-width: 680px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
  }
  .narrative p {
    font-style: italic;
    color: #444;
    line-height: 1.6;
    font-size: 15px;
  }

  /* ── Chat panel ── */
  .chat-panel {
    width: 320px;
    min-width: 320px;
    background: #fff;
    border-radius: 12px;
    box-shadow: 0 4px 20px rgba(0,0,0,0.1);
    overflow: hidden;
    max-height: 65vh;
    display: flex;
    flex-direction: column;
  }
  .chat-header {
    padding: 14px 16px;
    color: #fff;
    font-weight: 600;
    font-size: 14px;
  }
  .chat-messages {
    flex: 1;
    padding: 16px;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 12px;
  }
  .chat-bubble {
    opacity: 0;
    animation: bubbleFadeIn 400ms ease forwards;
  }
  @keyframes bubbleFadeIn { to { opacity: 1; } }

  .user-bubble .bubble-content {
    background: #00C9A7;
    color: #fff;
    padding: 10px 14px;
    border-radius: 16px 16px 4px 16px;
    font-size: 13px;
    line-height: 1.4;
    margin-left: 40px;
  }
  .agent-bubble .bubble-content {
    background: #f0f0f0;
    color: #333;
    padding: 10px 14px;
    border-radius: 16px 16px 16px 4px;
    font-size: 13px;
    line-height: 1.4;
    margin-right: 40px;
  }
  .thinking-bubble { margin-right: 40px; }
  .thinking-dots {
    display: flex;
    gap: 4px;
    padding: 10px 14px;
    background: #f0f0f0;
    border-radius: 16px;
    width: fit-content;
  }
  .thinking-dots span {
    width: 6px; height: 6px;
    background: #999;
    border-radius: 50%;
    animation: dotPulse 1.2s ease-in-out infinite;
  }
  .thinking-dots span:nth-child(2) { animation-delay: 0.2s; }
  .thinking-dots span:nth-child(3) { animation-delay: 0.4s; }
  @keyframes dotPulse {
    0%, 80%, 100% { opacity: 0.3; transform: scale(0.8); }
    40% { opacity: 1; transform: scale(1); }
  }

  .play-clip-btn {
    margin: 12px 16px;
    padding: 10px;
    background: #00C9A7;
    color: #fff;
    border: none;
    border-radius: 8px;
    cursor: pointer;
    font-size: 13px;
    font-weight: 600;
  }
  .play-clip-btn:hover { background: #005a9e; }

  /* ── Video ── */
  .video-container {
    text-align: center;
    margin-bottom: 16px;
  }
  .clip-video {
    max-height: 60vh;
    max-width: 100%;
    border-radius: 8px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.1);
  }
  .replay-btn {
    margin-top: 8px;
    padding: 8px 20px;
    background: #eee;
    border: 1px solid #ccc;
    border-radius: 6px;
    cursor: pointer;
    font-size: 13px;
  }

  /* ── Bottom bar ── */
  .bottom-bar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 10px 24px;
    background: #fff;
    border-top: 1px solid #e0e0e0;
    position: fixed;
    bottom: 0;
    left: 0;
    right: 0;
    z-index: 100;
  }
  .nav-btn {
    padding: 8px 24px;
    border: 1px solid #ccc;
    border-radius: 6px;
    background: #fff;
    cursor: pointer;
    font-size: 14px;
    font-weight: 500;
    transition: all 150ms;
  }
  .nav-btn:hover { background: #f0f0f0; }
  .nav-btn:disabled { opacity: 0.4; cursor: default; }
  .nav-btn.primary {
    background: ${config.brand_color};
    color: #fff;
    border-color: ${config.brand_color};
  }
  .nav-btn.primary:hover { opacity: 0.9; }

  .platform-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 4px 14px;
    border: 1.5px solid;
    border-radius: 20px;
    font-size: 13px;
    font-weight: 600;
    transition: all 350ms ease;
  }
  .platform-badge svg { width: 14px; height: 14px; }

  ${generateAnnotationCSS()}
</style>
</head>
<body>

<div class="top-bar">
  <div class="top-left">
    ${iconHTML}
    <span class="agent-title">${escapeHtml(config.title)}</span>
  </div>
  <span class="slide-counter" id="slideCounter">1 of ${totalSlides}</span>
</div>

<div class="progress-bar">
  <div class="progress-fill" id="progressFill" style="width:${(1 / totalSlides) * 100}%"></div>
</div>

<div class="slides-wrapper">
  <div class="slides-track" id="slidesTrack">
    ${slidesHTML}
  </div>
</div>

<div class="bottom-bar">
  <button class="nav-btn" id="prevBtn" onclick="navigate(-1)" disabled>← Back</button>
  <span id="platformBadge">${platformBadges[0]?.badge || ''}</span>
  <button class="nav-btn primary" id="nextBtn" onclick="navigate(1)">Next →</button>
</div>

<script>
  const totalSlides = ${totalSlides};
  let currentSlide = 0;

  const platformBadges = ${JSON.stringify(platformBadges.map(b => b.badge))};

  function navigate(dir) {
    const next = currentSlide + dir;
    if (next < 0 || next >= totalSlides) return;
    currentSlide = next;
    updateView();
  }

  function updateView() {
    const track = document.getElementById('slidesTrack');
    track.style.transform = 'translateX(-' + (currentSlide * 100) + '%)';

    document.getElementById('slideCounter').textContent = (currentSlide + 1) + ' of ' + totalSlides;
    document.getElementById('progressFill').style.width = ((currentSlide + 1) / totalSlides * 100) + '%';
    document.getElementById('prevBtn').disabled = currentSlide === 0;
    document.getElementById('nextBtn').disabled = currentSlide === totalSlides - 1;
    document.getElementById('nextBtn').textContent = currentSlide === totalSlides - 1 ? 'Done ✓' : 'Next →';

    document.getElementById('platformBadge').innerHTML = platformBadges[currentSlide] || '';

    // Autoplay video if present
    const videos = document.querySelectorAll('.clip-video');
    videos.forEach(v => { v.pause(); v.currentTime = 0; });

    const currentVid = document.querySelector('.slide[data-slide="' + currentSlide + '"] .clip-video');
    if (currentVid) {
      currentVid.parentElement.style.display = 'block';
      currentVid.play().catch(() => {});
    }
  }

  function playClip(slideIndex) {
    const container = document.getElementById('video-' + slideIndex);
    if (!container) return;
    container.style.display = 'block';
    const video = container.querySelector('video');
    if (video) {
      video.currentTime = 0;
      video.play().catch(() => {});
      video.onended = () => {
        container.querySelector('.replay-btn').style.display = 'inline-block';
      };
    }
  }

  function replayClip(slideIndex) {
    const container = document.getElementById('video-' + slideIndex);
    if (!container) return;
    const video = container.querySelector('video');
    if (video) {
      video.currentTime = 0;
      video.play().catch(() => {});
    }
  }

  // Keyboard navigation
  document.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowRight' || e.key === ' ') { e.preventDefault(); navigate(1); }
    if (e.key === 'ArrowLeft') { e.preventDefault(); navigate(-1); }
  });

  updateView();
</script>
</body>
</html>`;

  const outputPath = path.join(outputDir, 'demo.html');
  fs.writeFileSync(outputPath, html);

  // Update session meta
  const metaPath = path.join(demoDir, '.session-meta.json');
  if (fs.existsSync(metaPath)) {
    const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'));
    meta.last_generated = new Date().toISOString();
    fs.writeFileSync(metaPath, JSON.stringify(meta, null, 2));
  }

  console.log(`\n✓ Demo generated: ${outputPath}`);
  console.log(`  Open this file in a browser to view the interactive demo.`);

  // Note about clips
  const hasClips = config.slides.some(s =>
    s.clip_filename && fs.existsSync(path.join(demoDir, 'clips', s.clip_filename))
  );
  if (hasClips) {
    console.log(`\n  Note: Video clips are referenced as relative paths.`);
    console.log(`  Keep the clips/ folder alongside demo.html when sharing.`);
  }
}
