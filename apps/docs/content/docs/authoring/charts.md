---
title: "Native OpenXML Charts"
description: "Bar, Column, Line, Area, Doughnut, Pie, and Radar charts with series styling and <c:dPt> data point custom colors."
order: 8
section: "authoring"
---

# Native OpenXML Charts

Generate native PowerPoint charts rendered by Office Open XML DrawingML chart spaces (`<c:chartSpace>`).

```typescript
import { inches } from '@hokkyss/pptx';

slide.addChart({
  chartType: 'bar',
  title: 'Throughput by Runtime Engine (k ops/s)',
  categories: ['Parser Core', 'XML Serializer', 'Full Deck Gen'],
  series: [
    {
      name: '@hokkyss/pptx',
      values: [9.38, 5070, 0.358],
      color: '0284C7',
    },
    {
      name: 'Legacy Competitor',
      values: [1.2, 850, 0.045],
      color: '64748B',
    },
  ],
  x: inches(1),
  y: inches(2),
  w: inches(11.33),
  h: inches(4.5),
  legend: { position: 'bottom' },
  showGridlines: true,
});
```
