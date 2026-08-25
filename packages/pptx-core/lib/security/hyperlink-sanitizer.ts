import type { PptxHyperlinkAction } from '../text';

/**
 * Control characters forbidden in safe URLs and ScreenTips.
 * Matches ASCII 0x00 - 0x1F and 0x7F.
 */
// eslint-disable-next-line no-control-regex
export const CONTROL_CHARS_REGEX = /[\x00-\x1F\x7F]/g;

/**
 * Dangerous URI schemes that can lead to code execution, script injection (XSS), or local file access.
 */
export const DISALLOWED_SCHEMES_REGEX = /^(javascript|vbscript|data|file|ms-msdt|search-ms|powershell|cmd|bash|rundll32|shell):/i;

/**
 * Safe URI schemes allowed for external hyperlinks.
 */
export const SAFE_SCHEMES_REGEX = /^(https?|mailto|ftp|ftps|tel):/i;

/**
 * Standard PowerPoint slide show jump action strings.
 */
export const STANDARD_ACTIONS: readonly PptxHyperlinkAction[] = [
  'endShow',
  'firstSlide',
  'lastSlide',
  'nextSlide',
  'previousSlide',
] as const;

export const MAX_TOOLTIP_LENGTH = 2048;
export const MIN_SLIDE_INDEX = 1;
export const MAX_SLIDE_INDEX = 50_000;

/**
 * Sanitizes and validates an external hyperlink URL.
 *
 * Strips whitespace, control characters, and blocks dangerous URI schemes
 * (e.g. `javascript:`, `vbscript:`, `data:`, `file:`, `ms-msdt:`).
 * @param url Raw input URL.
 * @returns Clean, safe URL string, or `undefined` if empty or dangerous.
 * @example
 * ```ts
 * sanitizeHyperlinkUrl('https://example.com'); // 'https://example.com'
 * sanitizeHyperlinkUrl('javascript:alert(1)'); // undefined
 * ```
 */
export function sanitizeHyperlinkUrl(url?: string): string | undefined {
  if (!url || typeof url !== 'string') return undefined;

  // Strip leading/trailing whitespaces and control characters
  const clean = url.trim().replace(CONTROL_CHARS_REGEX, '');
  if (!clean) return undefined;

  // Check for disallowed dangerous schemes
  if (DISALLOWED_SCHEMES_REGEX.test(clean)) {
    return undefined;
  }

  // Allow safe protocols, relative paths, query params, or anchor fragments
  if (
    SAFE_SCHEMES_REGEX.test(clean)
    || clean.startsWith('#')
    || clean.startsWith('/')
    || clean.startsWith('./')
    || clean.startsWith('../')
    || clean.startsWith('?')
  ) {
    return clean;
  }

  // If no scheme is specified (e.g. "www.example.com" or "example.com/page"), allow it
  if (!clean.includes(':')) {
    return clean;
  }

  // Disallow any unknown custom protocol schemes by default for security
  return undefined;
}

/**
 * Sanitizes ScreenTip / hover tooltip text.
 *
 * Strips control characters (including line breaks that cause XML breakout) and enforces a maximum length limit.
 * @param tooltip Raw tooltip string.
 * @param maxLength Maximum allowed length (default: 2048).
 * @returns Sanitized tooltip or `undefined`.
 */
export function sanitizeHyperlinkTooltip(tooltip?: string, maxLength = MAX_TOOLTIP_LENGTH): string | undefined {
  if (!tooltip || typeof tooltip !== 'string') return undefined;

  const clean = tooltip.replace(CONTROL_CHARS_REGEX, '').trim();
  if (!clean) return undefined;

  return clean.length > maxLength ? clean.slice(0, maxLength) : clean;
}

/**
 * Validates and sanitizes a target slide index for internal jump navigation.
 *
 * Enforces integer bounds between 1 and 50,000, rejecting NaN, Infinity, negative values, and non-numeric inputs.
 * @param slideIndex Target slide index.
 * @returns Safe integer slide index or `undefined`.
 */
export function sanitizeSlideIndex(slideIndex?: null | number | string): number | undefined {
  if (slideIndex === undefined || slideIndex === null) return undefined;

  let num: number | undefined;
  if (typeof slideIndex === 'number') {
    num = slideIndex;
  } else if (typeof slideIndex === 'string') {
    num = parseInt(slideIndex, 10);
  }

  if (num === undefined || !Number.isSafeInteger(num)) return undefined;
  if (num < MIN_SLIDE_INDEX || num > MAX_SLIDE_INDEX) return undefined;

  return num;
}

/**
 * Validates and normalizes slide show action strings.
 *
 * Supports standard actions ('firstSlide', 'lastSlide', 'nextSlide', 'previousSlide', 'endShow')
 * and validates custom `ppaction://` URIs without dangerous payloads.
 * @param action Raw action string.
 * @returns Validated action or `undefined`.
 */
export function sanitizeHyperlinkAction(action?: string): PptxHyperlinkAction | undefined {
  if (!action || typeof action !== 'string') return undefined;

  const clean = action.trim().replace(CONTROL_CHARS_REGEX, '');
  if (!clean) return undefined;

  // Match known standard action names (case-insensitive check)
  const lower = clean.toLowerCase();
  if (lower === 'firstslide') return 'firstSlide';
  if (lower === 'lastslide') return 'lastSlide';
  if (lower === 'nextslide') return 'nextSlide';
  if (lower === 'prevslide' || lower === 'previousslide') return 'previousSlide';
  if (lower === 'endshow') return 'endShow';

  // Custom action URI (must be ppaction:// and safe)
  if (clean.startsWith('ppaction://') && !DISALLOWED_SCHEMES_REGEX.test(clean)) {
    return clean;
  }

  return undefined;
}
