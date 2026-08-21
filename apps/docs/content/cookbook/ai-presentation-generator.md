---
title: "AI Presentation Generator"
description: "Translate LLM JSON outputs directly into styled PowerPoint decks."
category: "AI & LLM"
difficulty: "Advanced"
---

# AI Presentation Generator

```typescript
import { Presentation, inches, points } from '@hokkyss/pptx';

interface AiSlideSpec {
  title: string;
  bullets: string[];
}

export async function buildPresentationFromAi(deckSpec: { title: string; slides: AiSlideSpec[] }) {
  const pres = Presentation.create({ title: deckSpec.title });

  for (const item of deckSpec.slides) {
    const slide = pres.addSlide();
    slide.addText(item.title, { x: inches(1), y: inches(1), fontSize: points(28), bold: true });
    slide.addText(item.bullets.map(b => ({ level: 0, text: b })), { x: inches(1), y: inches(2.2), w: inches(11), h: inches(4.5) });
  }

  return await pres.toArrayBuffer();
}
```
