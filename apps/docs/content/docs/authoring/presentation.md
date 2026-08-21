---
title: "Presentations & Masters"
description: "Create presentations, configure slide dimensions (16:9, 4:3), set metadata, and manage slide masters."
order: 1
section: "authoring"
---

# Presentations & Masters

The `Presentation` class is the top-level entry point in `@hokkyss/pptx`.

```typescript
import { Presentation, inches } from '@hokkyss/pptx';

// 1. Initialize presentation
const pres = Presentation.create({
  title: 'Executive Board Meeting',
  author: 'Chief Architect',
  company: 'Global Fintech Inc',
  revision: 1,
});

// 2. Custom Slide Dimensions (Default is 16:9 Widescreen: 13.33 x 7.5 inches)
pres.setSize({
  width: inches(13.33),
  height: inches(7.5),
});

// 3. Add slides
const slide1 = pres.addSlide();
const slide2 = pres.addSlide();

// 4. Output results
await pres.save('board_deck.pptx');
const arrayBuffer = await pres.toArrayBuffer();
const ast = pres.toAst();
```
