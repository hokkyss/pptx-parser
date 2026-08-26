---
title: "Architecture Diagrams & Flowcharts"
description: "Link microservice cards with connected arrows that glue to top, bottom, left, and right sites."
category: "Diagrams"
difficulty: "Intermediate"
---

# Architecture Diagrams & Flowcharts

```typescript
import { Presentation, inches } from '@hokkyss/pptx';

const pres = Presentation.create({ title: 'Microservice Pipeline' });
const slide = pres.addSlide();

slide.addShape('roundRect', { id: 'gateway', text: 'Gateway', x: inches(1), y: inches(3), w: inches(2.5), h: inches(1.5) });
slide.addShape('roundRect', { id: 'service-a', text: 'Auth API', x: inches(5), y: inches(1.5), w: inches(2.5), h: inches(1.5) });
slide.addShape('roundRect', { id: 'service-b', text: 'Orders API', x: inches(5), y: inches(4.5), w: inches(2.5), h: inches(1.5) });

slide.addConnector({ from: { shapeId: 'gateway', position: 'right' }, to: { shapeId: 'service-a', position: 'left' }, endArrow: 'triangle', shapeType: 'bentConnector2' });
slide.addConnector({ from: { shapeId: 'gateway', position: 'right' }, to: { shapeId: 'service-b', position: 'left' }, endArrow: 'triangle', shapeType: 'bentConnector2' });

await pres.save('architecture.pptx');
```
