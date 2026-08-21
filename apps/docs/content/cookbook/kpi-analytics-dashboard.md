---
title: "KPI Analytics Dashboard"
description: "Pair multi-column data tables with OpenXML Bar and Line charts for financial reporting."
category: "Analytics"
difficulty: "Intermediate"
---

# KPI Analytics Dashboard

```typescript
import { Presentation, inches, points } from '@hokkyss/pptx';

const pres = Presentation.create({ title: 'Q4 Financial Dashboard' });
const slide = pres.addSlide();

// Left Table
slide.addTable({
  x: inches(0.8),
  y: inches(1.5),
  w: inches(5.5),
  columns: [{ width: inches(2.5) }, { width: inches(1.5) }, { width: inches(1.5) }],
  headers: ['Region', 'Revenue ($M)', 'Growth'],
  rows: [
    ['North America', '$142.5', '+28%'],
    ['EMEA', '$98.2', '+19%'],
    ['APAC', '$64.1', '+42%'],
  ],
  headerStyle: { background: '0284C7', color: 'FFFFFF', bold: true },
});

// Right Chart
slide.addChart('col', {
  x: inches(6.8),
  y: inches(1.5),
  w: inches(5.8),
  h: inches(4.5),
  title: 'Revenue Comparison ($M)',
  categories: ['NA', 'EMEA', 'APAC'],
  series: [{ name: '2026', values: [142.5, 98.2, 64.1], color: '0284C7' }],
});

await pres.save('kpi_dashboard.pptx');
```
