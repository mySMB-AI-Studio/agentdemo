import yaml from 'js-yaml';
import fs from 'fs';
import path from 'path';

const VALID_PLATFORMS = [
  'm365-copilot', 'sharepoint', 'power-automate',
  'teams', 'outlook', 'xero', 'custom',
];

const VALID_ANNOTATION_TYPES = ['box', 'arrow', 'badge', 'spotlight'];

const HEX_COLOR_RE = /^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/;

export function parseConfig(configPath) {
  const absPath = path.resolve(configPath);
  if (!fs.existsSync(absPath)) {
    throw new Error(`Config file not found: ${absPath}`);
  }
  const raw = fs.readFileSync(absPath, 'utf8');
  const doc = yaml.load(raw);
  validate(doc);
  return normalize(doc);
}

function validate(doc) {
  const errors = [];

  if (!doc || !doc.demo) {
    throw new Error('Config must have a top-level "demo" key');
  }

  const d = doc.demo;

  if (!d.title || typeof d.title !== 'string') {
    errors.push('demo.title is required and must be a string');
  }
  if (!d.description || typeof d.description !== 'string') {
    errors.push('demo.description is required and must be a string');
  }
  if (d.brand_color && !HEX_COLOR_RE.test(d.brand_color)) {
    errors.push(`demo.brand_color must be a valid hex color, got: ${d.brand_color}`);
  }
  if (!d.m365_copilot_url || typeof d.m365_copilot_url !== 'string') {
    errors.push('demo.m365_copilot_url is required');
  }

  if (!Array.isArray(d.slides) || d.slides.length === 0) {
    errors.push('demo.slides must be a non-empty array');
  } else {
    const ids = new Set();
    for (const slide of d.slides) {
      if (slide.id == null) {
        errors.push('Every slide must have an id');
      } else if (ids.has(slide.id)) {
        errors.push(`Duplicate slide id: ${slide.id}`);
      } else {
        ids.add(slide.id);
      }

      if (!VALID_PLATFORMS.includes(slide.platform)) {
        errors.push(`Slide ${slide.id}: invalid platform "${slide.platform}". Valid: ${VALID_PLATFORMS.join(', ')}`);
      }

      if (!slide.story_label) {
        errors.push(`Slide ${slide.id}: story_label is required`);
      }
      if (!slide.narrative) {
        errors.push(`Slide ${slide.id}: narrative is required`);
      }

      if (slide.platform !== 'm365-copilot' && !slide.url) {
        errors.push(`Slide ${slide.id}: url is required for platform "${slide.platform}"`);
      }

      if (slide.annotations) {
        for (let i = 0; i < slide.annotations.length; i++) {
          const ann = slide.annotations[i];
          if (!VALID_ANNOTATION_TYPES.includes(ann.type)) {
            errors.push(`Slide ${slide.id}, annotation ${i}: invalid type "${ann.type}"`);
          }
          if (!ann.position || ann.position.x == null || ann.position.y == null) {
            errors.push(`Slide ${slide.id}, annotation ${i}: position with x and y is required`);
          }
        }
      }
    }
  }

  if (errors.length > 0) {
    throw new Error(`Config validation errors:\n  - ${errors.join('\n  - ')}`);
  }
}

function normalize(doc) {
  const d = doc.demo;
  const configDir = '';

  d.brand_color = d.brand_color || '#00C9A7';
  d.agent_icon = d.agent_icon || null;

  for (const slide of d.slides) {
    // Default url for m365-copilot slides
    if (slide.platform === 'm365-copilot' && !slide.url) {
      slide.url = d.m365_copilot_url;
    }

    // Default record_clip for m365-copilot
    if (slide.platform === 'm365-copilot' && slide.record_clip == null) {
      slide.record_clip = true;
    }

    if (slide.wait_for_response == null) {
      slide.wait_for_response = true;
    }

    slide.sample_prompts = slide.sample_prompts || [];
    slide.annotations = slide.annotations || [];

    // Auto-set filenames
    if (!slide.screenshot_filename) {
      slide.screenshot_filename = `${slide.id}-${slide.platform}-final.png`;
    }
    if (!slide.clip_filename && slide.record_clip) {
      slide.clip_filename = `${slide.id}-prompt-1.mp4`;
    }

    // Normalize annotations
    for (const ann of slide.annotations) {
      ann.highlight_color = ann.highlight_color || getPlatformColor(slide.platform, d.brand_color);
      ann.label = ann.label || '';
      ann.description = ann.description || '';
    }
  }

  return d;
}

export const PLATFORM_COLORS = {
  'm365-copilot': '#00C9A7',
  'sharepoint': '#036C70',
  'power-automate': '#0066FF',
  'teams': '#5558AF',
  'outlook': '#0072C6',
  'xero': '#13B5EA',
  'custom': '#00C9A7',
};

export function getPlatformColor(platform, fallback) {
  return PLATFORM_COLORS[platform] || fallback || '#00C9A7';
}

export function getDemoDir(configPath) {
  return path.dirname(path.resolve(configPath));
}
