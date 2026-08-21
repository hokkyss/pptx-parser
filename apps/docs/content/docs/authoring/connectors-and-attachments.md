---
title: "Connectors & Shape Attachment"
description: "Attach connectors to shapes via top/bottom/left/right connection sites with arrowheads."
order: 5
section: "authoring"
---

# Connectors & Shape Attachment

Attach connectors between shapes using OpenXML Connection Sites (`<a:stCxn>` and `<a:endCxn>`). PowerPoint, Keynote, and Google Slides "glue" lines to shapes so they stay connected when moved.

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

// 3. Connect Right side of Gateway to Left side of Auth Service
slide.addConnector({
  from: { shapeId: 'api-gateway', position: 'right' },
  to: { shapeId: 'auth-service', position: 'left' },
  color: '0284C7',
  width: inches(0.025),
  shapeType: 'bentConnector2',
  endArrow: 'triangle',
});
```
