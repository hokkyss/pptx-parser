---
title: "Executive Pitch Deck"
description: "Build an executive pitch deck with dark theme styling, KPI callouts, and drop shadows."
category: "Design & Presentation"
difficulty: "Beginner"
---

# Executive Pitch Deck

Build a sleek executive pitch deck with dark theme styling, ambient glows, and metric cards.

```typescript
import { Presentation, inches, points, degrees } from '@hokkyss/pptx';

const pres = Presentation.create({ title: 'Executive Keynote' });
pres.setThemeColors({ accent1: '#0284C7', dk1: '#0F172A', lt1: '#FFFFFF' });

const slide = pres.addSlide();
slide.setBackground('0F172A');

// Ambient Accent Pill
slide.addShape('roundRect', {
  x: inches(4.5),
  y: inches(1.2),
  w: inches(4.33),
  h: inches(0.45),
  fill: '1E293B',
  line: { color: '38BDF8', width: inches(0.015) },
  text: '⚡ NEXT-GEN ISOMORPHIC ENGINE',
  textOptions: { color: '38BDF8', bold: true, fontSize: points(10), align: 'center' },
});

// Title
slide.addText('Enterprise Scalability Deck', {
  x: inches(1),
  y: inches(2.0),
  w: inches(11.33),
  h: inches(1.2),
  fontSize: points(40),
  bold: true,
  color: 'FFFFFF',
  align: 'center',
});

await pres.save('pitch_deck.pptx');
```
