---
title: "Browser Client-Side Generation"
description: "Generate and download PowerPoint decks directly in the browser client."
order: 2
section: "runtimes-and-deploy"
---

# Browser Client-Side Generation

```typescript
import { Presentation } from '@hokkyss/pptx';

async function exportDeck() {
  const pres = Presentation.create({ title: 'Client Export' });
  const slide = pres.addSlide();
  slide.addText('Client-Side PPTX');

  const buffer = await pres.toArrayBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });
  
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'deck.pptx';
  a.click();
}
```
