---
title: "Connectors & Vector Lines"
description: "Add straight, bent, or curved connector lines with custom stroke colors and widths."
order: 5
section: "authoring"
---

# Connectors & Vector Lines

Add vector connectors and lines between coordinates on the slide canvas:

```typescript
import { inches } from '@hokkyss/pptx';

// 1. Add Source Card
slide.addShape('roundRect', {
  id: 'api-gateway',
  text: 'API Gateway',
  x: inches(1),
  y: inches(2.5),
  w: inches(3),
  h: inches(1.5),
});

// 2. Add Target Card
slide.addShape('roundRect', {
  id: 'auth-service',
  text: 'Auth Service',
  x: inches(6),
  y: inches(2.5),
  w: inches(3),
  h: inches(1.5),
});

// 3. Connect from right edge of Gateway to left edge of Auth Service
slide.addConnector({
  from: { x: inches(4), y: inches(3.25) },
  to: { x: inches(6), y: inches(3.25) },
  color: '0284C7',
  width: inches(0.025),
  shapeType: 'bentConnector2',
  dashStyle: 'solid',
});
```
