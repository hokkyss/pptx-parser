---
title: "Text & Multilevel Bullets"
description: "Rich text formatting, hierarchical bullet lists (levels 0-8), superscript/subscript, alignment, and line spacing."
order: 3
section: "authoring"
---

# Text & Multilevel Bullets

`@hokkyss/pptx` supports both simple strings and advanced **multilevel hierarchical bullet matrices** with inline text runs.

```typescript
import { points, inches } from '@hokkyss/pptx';

slide.addText([
  {
    // Level 0 (outermost bullet)
    level: 0,
    runs: [
      { text: 'Core Architecture Pillar: ', bold: true },
      { text: 'High Availability', bold: true, color: '0284C7' },
      // Shift+Enter behavior: soft line break within the same bullet
      { break: true },
      { text: 'Continuous operations across primary and secondary regions.' },
    ],
  },
  {
    // Level 1 (Tab - one level deeper)
    level: 1,
    text: 'Active-active multi-region replication across 3 availability zones.',
  },
  {
    // Level 2 (Tab Tab - two levels deeper)
    level: 2,
    runs: [
      { text: 'Latency: O(1) in-memory cache lookup profile: 10' },
      { text: '-6', superscript: true },
      { text: ' seconds' },
    ],
  },
  {
    // Level 0 (Shift+Tab back to root)
    level: 0,
    text: 'Secondary Architecture Pillar',
  },
], {
  x: inches(1),
  y: inches(1.5),
  w: inches(11.33),
  h: inches(5.0),
  lineSpacing: points(24),
  spaceBefore: points(8),
  spaceAfter: points(8),
});
```

### Shift + Enter Behavior (Soft Line Breaks)

In PowerPoint desktop, pressing `Shift + Enter` inside a bulleted paragraph breaks to a new line **without spawning a new bullet point**.

To achieve this in `@hokkyss/pptx`, place a `{ break: true }` entry in your paragraph's `runs` array. It is serialized as an OpenXML `<a:br>` element:

```typescript
slide.addText([
  {
    level: 0,
    runs: [
      { text: 'First line with bullet' },
      { break: true },
      { text: 'Second line indented under same bullet (no new bullet point)' },
    ],
  },
]);
```

