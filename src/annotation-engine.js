import { getPlatformColor } from './config-parser.js';

export function generateAnnotationCSS() {
  return `
    .annotation-overlay {
      position: absolute;
      top: 0; left: 0; right: 0; bottom: 0;
      pointer-events: none;
      z-index: 10;
    }
    .annotation {
      position: absolute;
      pointer-events: auto;
      cursor: pointer;
      opacity: 0;
      animation: annotationFadeIn 300ms ease forwards;
    }
    @keyframes annotationFadeIn {
      to { opacity: 1; }
    }
    .annotation-box {
      border: 2px dashed var(--ann-color);
      background: var(--ann-bg);
      border-radius: 6px;
      min-width: 120px;
      min-height: 60px;
      padding: 8px 12px;
    }
    .annotation-box .ann-label {
      font-size: 12px;
      font-weight: 600;
      color: var(--ann-color);
    }
    .annotation-arrow {
      width: 0; height: 0;
    }
    .annotation-arrow::after {
      content: '';
      display: block;
      width: 40px;
      height: 2px;
      background: var(--ann-color);
      position: relative;
      animation: arrowPulse 1.5s ease-in-out infinite;
    }
    .annotation-arrow::before {
      content: '';
      display: block;
      width: 0; height: 0;
      border-left: 8px solid var(--ann-color);
      border-top: 5px solid transparent;
      border-bottom: 5px solid transparent;
      position: absolute;
      right: -2px;
      top: 50%;
      transform: translateY(-50%);
    }
    @keyframes arrowPulse {
      0%, 100% { opacity: 1; }
      50% { opacity: 0.5; }
    }
    .annotation-badge {
      background: var(--ann-color);
      color: #fff;
      padding: 4px 14px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 600;
      white-space: nowrap;
      box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }
    .annotation-spotlight {
      position: absolute;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(0,0,0,0.55);
      pointer-events: none;
    }
    .annotation-spotlight .cutout {
      position: absolute;
      border-radius: 8px;
      box-shadow: 0 0 0 9999px rgba(0,0,0,0.55);
      background: transparent;
    }
    .annotation-spotlight .spot-label {
      position: absolute;
      background: #fff;
      color: #333;
      padding: 6px 14px;
      border-radius: 6px;
      font-size: 13px;
      font-weight: 600;
      box-shadow: 0 2px 12px rgba(0,0,0,0.2);
      white-space: nowrap;
    }
    .annotation-tooltip {
      display: none;
      position: absolute;
      background: #fff;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
      padding: 12px 16px;
      max-width: 260px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 20;
      pointer-events: none;
    }
    .annotation:hover .annotation-tooltip {
      display: block;
    }
    .annotation-tooltip .tooltip-label {
      font-weight: 700;
      font-size: 13px;
      margin-bottom: 4px;
    }
    .annotation-tooltip .tooltip-desc {
      font-size: 12px;
      color: #555;
      line-height: 1.4;
    }
  `;
}

export function renderAnnotation(annotation, index) {
  const { type, position, label, description, highlight_color } = annotation;
  const x = position.x;
  const y = position.y;
  const color = highlight_color || '#0078D4';
  const bgColor = color + '1A'; // 10% opacity

  const tooltip = `
    <div class="annotation-tooltip" style="top: -60px; left: 0;">
      <div class="tooltip-label">${escapeHtml(label)}</div>
      ${description ? `<div class="tooltip-desc">${escapeHtml(description)}</div>` : ''}
    </div>
  `;

  const delay = 300 + index * 100;

  switch (type) {
    case 'box':
      return `
        <div class="annotation annotation-box"
             style="left:${x}%;top:${y}%;--ann-color:${color};--ann-bg:${bgColor};animation-delay:${delay}ms;">
          <span class="ann-label">${escapeHtml(label)}</span>
          ${tooltip}
        </div>`;

    case 'arrow':
      return `
        <div class="annotation annotation-arrow"
             style="left:${x}%;top:${y}%;--ann-color:${color};animation-delay:${delay}ms;">
          <span class="ann-label" style="position:absolute;top:-20px;left:0;font-size:11px;font-weight:600;color:${color};white-space:nowrap;">
            ${escapeHtml(label)}
          </span>
          ${tooltip}
        </div>`;

    case 'badge':
      return `
        <div class="annotation annotation-badge"
             style="left:${x}%;top:${y}%;--ann-color:${color};animation-delay:${delay}ms;">
          ${escapeHtml(label)}
          ${tooltip}
        </div>`;

    case 'spotlight':
      const cutoutW = 20;
      const cutoutH = 15;
      return `
        <div class="annotation annotation-spotlight" style="animation-delay:${delay}ms;">
          <div class="cutout" style="left:${x - cutoutW / 2}%;top:${y - cutoutH / 2}%;width:${cutoutW}%;height:${cutoutH}%;"></div>
          <div class="spot-label" style="left:${x + cutoutW / 2 + 2}%;top:${y}%;">
            ${escapeHtml(label)}
          </div>
          ${tooltip}
        </div>`;

    default:
      return '';
  }
}

export function renderAnnotations(annotations) {
  if (!annotations || annotations.length === 0) return '';
  const items = annotations.map((ann, i) => renderAnnotation(ann, i)).join('\n');
  return `<div class="annotation-overlay">${items}</div>`;
}

function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
