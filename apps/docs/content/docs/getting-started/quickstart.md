---
title: "Quick Start"
description: "Create your first PowerPoint presentation with custom themes, styled text, and bullet lists in under 30 seconds."
order: 3
section: "getting-started"
---

# Quick Start

Here is how to create a complete, styled presentation from scratch with **`@hokkyss/pptx`** in under 30 seconds.

```typescript
import { Presentation, inches, points } from '@hokkyss/pptx';

// 1. Initialize presentation (16:9 Widescreen by default)
const pres = Presentation.create({
  title: 'Quarterly Cloud Architecture Keynote',
  author: 'Platform Engineering Team',
  company: 'Enterprise Cloud Corp',
});

// 2. Configure Theme Colors & Fonts
pres
  .setThemeName('Modern Cobalt Theme')
  .setThemeColors({
    accent1: '#0284C7', // Sky Blue
    accent2: '#6366F1', // Indigo
    dk1: '#0F172A',     // Slate 900
    lt1: '#FFFFFF',     // White
  })
  .setThemeFonts({
    major: 'Inter',
    minor: 'Roboto',
  });

// 3. Add a Slide with Dark Background
const slide = pres.addSlide();
slide.setBackground('0F172A');

// 4. Add Title Header
slide.addText('Distributed Edge Compute', {
  bold: true,
  color: '38BDF8',
  fontSize: points(32),
  x: inches(1.0),
  y: inches(1.0),
  w: inches(11.33),
  h: inches(0.8),
});

// 5. Add Multilevel Bullet List
slide.addText([
  { level: 0, text: '1. Event-Driven Edge Runtime' },
  { level: 1, text: '1.1 Sub-millisecond latency profile at 300+ PoPs' },
  {
    level: 2,
    runs: [
      { text: '1.1.1 Time Complexity: O(log ' },
      { text: 'N', italic: true },
      { text: ')' },
    ],
  },
  { level: 0, text: '2. Zero Native C++ Bindings' },
  { level: 1, text: '2.1 100% pure TypeScript OpenXML codecs' },
], {
  x: inches(1.0),
  y: inches(2.0),
  w: inches(11.33),
  h: inches(4.5),
});

// 6. Save to disk
await pres.save('cloud_keynote.pptx');
console.log('✨ Generated cloud_keynote.pptx successfully!');
```
