---
title: "Security & Sanitization"
description: "Hyperlink sanitization, XML safety, and resource protections in @hokkyss/pptx-core."
order: 3
package: "@hokkyss/pptx-core"
section: "pptx-core"
---

# Security & Sanitization

The `@hokkyss/pptx-core` security module safeguards against document vulnerabilities, including Cross-Site Scripting (XSS), XML External Entity (XXE) injection, malicious protocol URI execution, control character breakouts, and ZIP bomb memory exhaustion.

```typescript
import {
  sanitizeHyperlinkUrl,
  sanitizeHyperlinkTooltip,
  sanitizeSlideIndex,
  sanitizeHyperlinkAction,
  SAFE_SCHEMES_REGEX,
  DISALLOWED_SCHEMES_REGEX
} from '@hokkyss/pptx-core';
```

---

## `sanitizeHyperlinkUrl(url)`

Validates and sanitizes an external URL before it is written into OpenXML DrawingML hyperlink relationships (`r:hlinkClick`). Strips leading/trailing whitespace, null bytes, and control characters (ASCII 0x00–0x1F and 0x7F).

```typescript
export function sanitizeHyperlinkUrl(url?: string): string | undefined
```

### Protocol Validation

- **Allowed Protocols (`SAFE_SCHEMES_REGEX`)**: `https://`, `http://`, `mailto:`, `ftp://`, `ftps://`, `tel:`, relative paths (`/`, `./`, `../`), query params (`?`), or anchor fragments (`#`).
- **Disallowed Protocols (`DISALLOWED_SCHEMES_REGEX`)**: `javascript:`, `vbscript:`, `data:`, `file:`, `ms-msdt:`, `search-ms:`, `powershell:`, `cmd:`, `bash:`, `rundll32:`, `shell:`.
- **Return Value**: Returns the cleaned, safe URL string, or `undefined` if empty, malformed, or matching a dangerous scheme (preventing script execution without throwing unhandled exceptions).

```typescript
sanitizeHyperlinkUrl('https://example.com/report.pdf'); // 'https://example.com/report.pdf'
sanitizeHyperlinkUrl('mailto:sales@acme.com');          // 'mailto:sales@acme.com'
sanitizeHyperlinkUrl('javascript:alert(1)');            // undefined
sanitizeHyperlinkUrl('file:///C:/Windows/System32');    // undefined
```

---

## `sanitizeHyperlinkTooltip(tooltip, maxLength?)`

Sanitizes ScreenTip hover tooltip text, removing control characters and newlines that could cause XML attribute breakout.

```typescript
export function sanitizeHyperlinkTooltip(
  tooltip?: string,
  maxLength = 2048
): string | undefined
```

---

## `sanitizeSlideIndex(slideIndex)`

Validates and sanitizes integer slide target indices for internal slide show jump hyperlinks (`action="ppaction://hlinksldjump"`).

```typescript
export function sanitizeSlideIndex(slideIndex?: unknown): number | undefined
```

- Enforces integer bounds between `1` (`MIN_SLIDE_INDEX`) and `50,000` (`MAX_SLIDE_INDEX`).
- Safely rejects `NaN`, `Infinity`, negative numbers, or non-numeric values by returning `undefined`.

---

## `sanitizeHyperlinkAction(action)`

Validates and normalizes slide show action navigation strings.

```typescript
export function sanitizeHyperlinkAction(action?: string): PptxHyperlinkAction | undefined
```

- Standard actions supported: `'firstSlide'`, `'lastSlide'`, `'nextSlide'`, `'previousSlide'`, `'endShow'`.
- Validates custom `ppaction://` URIs, ensuring no disallowed protocol payloads are embedded.

---

## XML & Archive Protections

1. **No External DTD / XXE Resolution**: XML parsers disable external entity resolution to prevent server-side request forgery (SSRF) and local file disclosures.
2. **Decompression Ratio Limits**: `pptx-reader` enforces maximum expansion ratio checks on ZIP entries to prevent ZIP bomb denial of service (DoS).
