---
title: "Slide Masters & Layouts"
description: "Inspect slide masters, discover child layouts, and instantiate slides bound to master templates."
order: 2
section: "authoring"
---

# Slide Masters & Layouts

Slide Masters and Slide Layouts allow defining reusable templates with semantic placeholders that slides can instantiate and populate.

```typescript
import { Presentation } from '@hokkyss/pptx';

const pres = await Presentation.load(templateBuffer);

// 1. Inspect masters and child layouts
const master = pres.getMaster('Office Theme');
const layouts = master?.getLayouts();

// 2. Instantiate slide bound to a specific layout
const slide = pres.addSlide({
  master,
  layout: 'Title and Content',
});

// 3. Populate semantic placeholders
slide.addText('Annual Strategy', { placeholder: 'title' });
slide.addText('• Expansion into APAC region
• Launching v2 API', { placeholder: 'body' });
```
