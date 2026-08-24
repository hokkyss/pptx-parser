import { describe, expect, it } from 'vitest';
import {
  sanitizeHyperlinkAction,
  sanitizeHyperlinkTooltip,
  sanitizeHyperlinkUrl,
  sanitizeSlideIndex,
} from '../../lib/security/hyperlink-sanitizer';

describe('Hyperlink Security Sanitizer (@hokkyss/pptx-core)', () => {
  describe('sanitizeHyperlinkUrl', () => {
    it('allows valid HTTPS, HTTP, and Mailto URLs', () => {
      expect(sanitizeHyperlinkUrl('https://example.com')).toBe('https://example.com');
      expect(sanitizeHyperlinkUrl('http://insecure.org/path?q=1#top')).toBe('http://insecure.org/path?q=1#top');
      expect(sanitizeHyperlinkUrl('mailto:developer@example.com?subject=Hello')).toBe('mailto:developer@example.com?subject=Hello');
      expect(sanitizeHyperlinkUrl('ftp://ftp.example.com')).toBe('ftp://ftp.example.com');
      expect(sanitizeHyperlinkUrl('tel:+1234567890')).toBe('tel:+1234567890');
    });

    it('allows relative URLs and anchor fragments', () => {
      expect(sanitizeHyperlinkUrl('#section-2')).toBe('#section-2');
      expect(sanitizeHyperlinkUrl('/docs/overview')).toBe('/docs/overview');
      expect(sanitizeHyperlinkUrl('./relative/path')).toBe('./relative/path');
      expect(sanitizeHyperlinkUrl('www.google.com')).toBe('www.google.com');
    });

    it('neutralizes dangerous script injection schemes (XSS)', () => {
      expect(sanitizeHyperlinkUrl('javascript:alert(1)')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('JAVASCRIPT:alert(document.cookie)')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('vbscript:msgbox(1)')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('data:text/html;base64,PHNjcmlwdD5hbGVydCgxKTwvc2NyaXB0Pg==')).toBeUndefined();
    });

    it('neutralizes local file and protocol execution exploits', () => {
      expect(sanitizeHyperlinkUrl('file:///C:/Windows/System32/calc.exe')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('ms-msdt:/id PCWDiagnostic')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('search-ms:query=calc.exe')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('powershell:-Command "calc"')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('cmd:/c dir')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('rundll32:shell32.dll,Control_RunDLL')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('unknown-custom-scheme:payload')).toBeUndefined();
    });

    it('strips ASCII control characters and CRLF injection attempts', () => {
      expect(sanitizeHyperlinkUrl('https://example.com\r\n\0/attack')).toBe('https://example.com/attack');
      expect(sanitizeHyperlinkUrl('  https://example.com  ')).toBe('https://example.com');
    });

    it('returns undefined for invalid or empty inputs', () => {
      expect(sanitizeHyperlinkUrl(undefined)).toBeUndefined();
      expect(sanitizeHyperlinkUrl('')).toBeUndefined();
      expect(sanitizeHyperlinkUrl('   ')).toBeUndefined();
    });
  });

  describe('sanitizeHyperlinkTooltip', () => {
    it('preserves clean tooltip text', () => {
      expect(sanitizeHyperlinkTooltip('Visit Official Website')).toBe('Visit Official Website');
    });

    it('strips control characters, newlines, and null bytes', () => {
      expect(sanitizeHyperlinkTooltip('Malicious\r\nHeader\0Injection')).toBe('MaliciousHeaderInjection');
    });

    it('truncates oversized tooltips to prevent memory exhaustion', () => {
      const hugeTooltip = 'A'.repeat(5000);
      const sanitized = sanitizeHyperlinkTooltip(hugeTooltip, 100);
      expect(sanitized).toHaveLength(100);
    });

    it('returns undefined for empty inputs', () => {
      expect(sanitizeHyperlinkTooltip(undefined)).toBeUndefined();
      expect(sanitizeHyperlinkTooltip('')).toBeUndefined();
    });
  });

  describe('sanitizeSlideIndex', () => {
    it('accepts valid integer slide indexes', () => {
      expect(sanitizeSlideIndex(1)).toBe(1);
      expect(sanitizeSlideIndex(42)).toBe(42);
      expect(sanitizeSlideIndex('10')).toBe(10);
    });

    it('rejects out-of-bounds, negative, NaN, and float numbers', () => {
      expect(sanitizeSlideIndex(0)).toBeUndefined();
      expect(sanitizeSlideIndex(-5)).toBeUndefined();
      expect(sanitizeSlideIndex(999_999)).toBeUndefined();
      expect(sanitizeSlideIndex(NaN)).toBeUndefined();
      expect(sanitizeSlideIndex(Infinity)).toBeUndefined();
      expect(sanitizeSlideIndex('invalid')).toBeUndefined();
      expect(sanitizeSlideIndex('../../etc/passwd')).toBeUndefined();
    });
  });

  describe('sanitizeHyperlinkAction', () => {
    it('normalizes standard PowerPoint actions', () => {
      expect(sanitizeHyperlinkAction('nextSlide')).toBe('nextSlide');
      expect(sanitizeHyperlinkAction('previousSlide')).toBe('previousSlide');
      expect(sanitizeHyperlinkAction('firstSlide')).toBe('firstSlide');
      expect(sanitizeHyperlinkAction('lastSlide')).toBe('lastSlide');
      expect(sanitizeHyperlinkAction('endShow')).toBe('endShow');
      expect(sanitizeHyperlinkAction('NEXTSLIDE')).toBe('nextSlide');
      expect(sanitizeHyperlinkAction('prevSlide')).toBe('previousSlide');
    });

    it('validates safe custom ppaction URIs', () => {
      expect(sanitizeHyperlinkAction('ppaction://customaction')).toBe('ppaction://customaction');
    });

    it('rejects non-ppaction or dangerous action payloads', () => {
      expect(sanitizeHyperlinkAction('javascript:alert(1)')).toBeUndefined();
      expect(sanitizeHyperlinkAction('customAction')).toBeUndefined();
      expect(sanitizeHyperlinkAction('')).toBeUndefined();
    });
  });
});
