---
title: "Shapes & Gradients"
description: "Vector auto-shapes, DrawingML linear & path gradients, borders, shadows, and rotation."
order: 4
section: "authoring"
---

# Shapes & Gradients

Add vector geometries with solid colors or rich **DrawingML Linear & Path Gradients**.

```typescript
import { inches, degrees, points } from '@hokkyss/pptx';

slide.addShape('roundRect', {
  x: inches(2),
  y: inches(2),
  w: inches(4.5),
  h: inches(2.5),
  gradient: {
    type: 'linear',
    angle: degrees(90),
    stops: [
      { position: 0, color: '0284C7', opacity: 1.0 },
      { position: 100, color: '0F172A', opacity: 0.85 },
    ],
  },
  line: {
    color: '38BDF8',
    width: inches(0.02),
    dashStyle: 'solid',
  },
  shadow: {
    color: '000000',
    opacity: 0.35,
    blur: inches(0.15),
    distance: inches(0.05),
    direction: degrees(90),
  },
  text: 'Cloud Native Microservice',
  textOptions: {
    bold: true,
    color: 'FFFFFF',
    fontSize: points(14),
    align: 'center',
  },
});
```
