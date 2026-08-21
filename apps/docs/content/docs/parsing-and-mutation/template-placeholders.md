---
title: "Template Mutation & Placeholders"
description: "Load existing corporate .pptx decks and populate text/image placeholders programmatically."
order: 3
section: "parsing-and-mutation"
---

# Template Mutation & Placeholders

```typescript
import { Presentation } from '@hokkyss/pptx';

const pres = await Presentation.load(templateBuffer);
const slide = pres.addSlide({ layout: 'master:title-with-body' });
slide.addText('Quarterly Sales Report', { placeholder: 'title' });
slide.addText('• Revenue up 34% YoY\n• Net Retention at 128%', { placeholder: 'body' });
await pres.save('populated_deck.pptx');
```
