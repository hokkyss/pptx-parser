---
title: "Text & Multi-Level Bullets"
description: "API specification for rich text runs, multi-level bullets, and paragraph formatting in @hokkyss/pptx."
order: 3
package: "@hokkyss/pptx"
section: "pptx"
---

# Text & Multi-Level Bullets

The text module in `@hokkyss/pptx` creates rich DrawingML text bodies (`<p:txBody>`) with hierarchical indentation levels, custom bullet types, character formatting, and inline hyperlink runs.

```typescript
import { buildTextBody, type ParagraphConfig, type TextRunConfig } from '@hokkyss/pptx';
```

---

## Data Structures

### `TextRunConfig`

Defines an inline span of text with uniform styling:

```typescript
export interface TextRunConfig {
  /** The text string content */
  text: string;
  /** Font family name */
  font?: string;
  /** Font size in points */
  fontSize?: Points;
  /** Color hex string (e.g. '0284C7') */
  color?: string;
  /** Bold font weight */
  bold?: boolean;
  /** Italic font style */
  italic?: boolean;
  /** Underline decoration style */
  underline?: boolean | 'dash' | 'dbl' | 'dotted' | 'heavy' | 'sng' | 'wave';
  /** Strikethrough decoration */
  strikethrough?: boolean | 'dblStrike' | 'sngStrike';
  /** Subscript baseline shift */
  subscript?: boolean;
  /** Superscript baseline shift */
  superscript?: boolean;
  /** Baseline shift percentage (e.g. 30000 for +30%) */
  baseline?: number;
  /** Clickable hyperlink URL or action */
  hyperlink?: PptxHyperlink | string;
}
```

---

### `ParagraphConfig`

Defines a paragraph containing one or more inline text runs, bullet formatting, margins, and alignment:

```typescript
export interface ParagraphConfig {
  /** Text content shorthand (string, TextRunConfig, or array of runs) */
  text?: (string | TextRunConfig)[] | string | TextRunConfig;
  /** Array of constituent text runs */
  runs?: (string | TextRunConfig)[];
  /** 0-based indentation level (0 to 8) */
  level?: number;
  /** Bullet formatting */
  bullet?: BulletInput;
  /** Horizontal alignment */
  align?: 'center' | 'justify' | 'left' | 'right';
  /** Space before paragraph in points */
  spaceBefore?: Points;
  /** Space after paragraph in points */
  spaceAfter?: Points;
  /** Line spacing height in points */
  lineSpacing?: Points;
  /** Default typography overrides for all runs in paragraph */
  font?: string;
  fontSize?: Points;
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | 'dash' | 'dbl' | 'dotted' | 'heavy' | 'sng' | 'wave';
  strikethrough?: boolean | 'dblStrike' | 'sngStrike';
  subscript?: boolean;
  superscript?: boolean;
  baseline?: number;
  hyperlink?: PptxHyperlink | string;
}
```

---

### `BulletInput`

```typescript
export type BulletInput =
  | boolean                     // true: '•', false: no bullet
  | 'bullet'                    // '•' bullet
  | 'number'                    // '1.' auto-numbered period
  | 'none'                      // no bullet
  | string                      // custom character e.g. '→', '✓', '★'
  | PptxBullet;                 // full AST object
```

---

## Utility Functions

### `buildTextBody(content, options?)`

Converts a string, ParagraphConfig array, or TextRunConfig array into a strongly-typed `PptxTextBody` AST node.

```typescript
export function buildTextBody(
  content: (ParagraphConfig | string | TextRunConfig)[] | string,
  options?: TextOptions
): PptxTextBody
```
