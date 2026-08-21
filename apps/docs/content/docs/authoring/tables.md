---
title: "Enterprise Tables"
description: "Multi-column data tables with column widths, row heights, headers, zebra striping, and cell padding."
order: 6
section: "authoring"
---

# Enterprise Tables

Create multi-column tables with custom headers, alternating row styles, text alignment, and cell padding.

```typescript
import { inches, points } from '@hokkyss/pptx';

slide.addTable({
  x: inches(1),
  y: inches(1.5),
  w: inches(11.33),
  columns: [
    { width: inches(3.5) },
    { width: inches(2.5) },
    { width: inches(2.5) },
    { width: inches(2.83) },
  ],
  headers: ['Microservice Component', 'P99 Latency', 'Throughput (RPS)', 'Health SLA'],
  headerStyle: {
    background: '0284C7',
    color: 'FFFFFF',
    bold: true,
    fontSize: points(12),
    align: 'center',
  },
  rows: [
    ['Edge API Gateway', '0.12 ms', '125,000 rps', '99.999%'],
    ['Authentication Engine', '0.45 ms', '95,000 rps', '99.99%'],
    ['Distributed Session Store', '0.28 ms', '180,000 rps', '99.999%'],
  ],
  alternateRowBackground: 'F1F5F9',
  padding: inches(0.1),
});
```
