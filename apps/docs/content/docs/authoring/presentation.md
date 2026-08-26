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

// 1. Initialize presentation with dimensions and metadata
const pres = Presentation.create({
  title: 'Executive Board Meeting',
  author: 'Chief Architect',
  company: 'Global Fintech Inc',
  width: inches(13.333),
  height: inches(7.5),
  firstSlideNumber: 1,
});

// 2. Add slides
const slide1 = pres.addSlide();
const slide2 = pres.addSlide();

// 3. Serialize output
const buffer = await pres.toBuffer();
const arrayBuffer = await pres.toArrayBuffer();
const ast = pres.ast;
```
