---
title: "Enterprise Tables"
description: "Multi-column data tables with column widths, row heights, headers, and cell formatting."
order: 6
section: "authoring"
---

# Enterprise Tables

Create multi-column tables with custom headers, column widths, text alignment, and cell styling.

```typescript
import { inches, points } from '@hokkyss/pptx';

slide.addTable([
  ['Microservice Component', 'P99 Latency', 'Throughput (RPS)', 'Health SLA'],
  ['Edge API Gateway', '0.12 ms', '125,000 rps', '99.999%'],
  ['Authentication Engine', '0.45 ms', '95,000 rps', '99.99%'],
  ['Distributed Session Store', '0.28 ms', '180,000 rps', '99.999%'],
], {
  x: inches(1),
  y: inches(1.5),
  w: inches(11.33),
  h: inches(3.5),
  colWidths: [inches(3.5), inches(2.5), inches(2.5), inches(2.83)],
  header: {
    fill: '0284C7',
    color: 'FFFFFF',
    bold: true,
    fontSize: points(12),
    align: 'center',
  },
});
```
