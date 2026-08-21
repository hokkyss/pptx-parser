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
    level: 0,
    runs: [
      { text: 'Core Architecture Pillar: ', bold: true },
      { text: 'High Availability', bold: true, color: '0284C7' },
    ],
  },
  {
    level: 1,
    text: 'Active-active multi-region replication across 3 availability zones.',
  },
  {
    level: 2,
    runs: [
      { text: 'Latency: O(1) in-memory cache lookup profile: 10' },
      { text: '-6', superscript: true },
      { text: ' seconds' },
    ],
  },
  {
    level: 2,
    runs: [
      { text: 'Legacy synchronous RPC: ' },
      { text: 'deprecated in v2.0', strikethrough: true },
    ],
  },
], {
  x: inches(1),
  y: inches(1.5),
  w: inches(11.33),
  h: inches(5.0),
  lineSpacing: 1.25,
});
```
