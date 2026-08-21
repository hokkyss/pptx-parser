---
title: "Template Mutation & Placeholders"
description: "Load existing corporate .pptx decks and populate text/image placeholders programmatically."
order: 3
section: "parsing-and-mutation"
---

# Template Mutation & Placeholders

Load existing corporate presentation templates, find placeholders, and inject dynamic content:

```typescript
import { Presentation } from '@hokkyss/pptx';

// 1. Load corporate master template
const pres = await Presentation.load(templateBuffer);

// 2. Add slide bound to a layout
const slide = pres.addSlide({ layout: 'Title and Content' });

// 3. Populate title and body placeholders
slide.addText('Quarterly Sales Report', { placeholder: 'title' });
slide.addText([
  { level: 0, text: '• Revenue up 34% YoY' },
  { level: 0, text: '• Net Retention at 128%' },
], { placeholder: 'body' });

// 4. Export mutated presentation
const buffer = await pres.toBuffer();
```
