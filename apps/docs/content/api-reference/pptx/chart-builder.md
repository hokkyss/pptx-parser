---
title: "ChartBuilder & addChart"
description: "API specification for native OpenXML DrawingML charts in @hokkyss/pptx."
order: 6
package: "@hokkyss/pptx"
section: "pptx"
---

# ChartBuilder & addChart

The `addChart` API creates native OpenXML DrawingML chart parts (`ppt/charts/chart*.xml`) embedded within graphic frames.

```typescript
import { inches, points } from '@hokkyss/pptx';

slide.addChart({
  chartType: 'col',
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: '2025 Actual', values: [32, 45, 58, 71], color: '2563EB' },
    { name: '2026 Target', values: [40, 55, 70, 90], color: '10B981' }
  ],
  x: inches(1),
  y: inches(1.8),
  w: inches(11.3),
  h: inches(4.5),
  title: 'Revenue Growth by Quarter ($M)',
  legend: { position: 'bottom' },
  showGridlines: true
});
```

---

## Function Signatures

```typescript
// Slide canvas method (returns this)
addChart(options: AddChartOptions): this

// Low-level builder function
export function buildChartElement(
  options: AddChartOptions,
  counter?: number
): PptxElement
```

---

## Chart Options (`AddChartOptions`)

```typescript
export interface AddChartOptions {
  /** Category labels array */
  categories: string[];
  /** Data series definitions */
  series: ChartSeriesConfig[];
  /** Chart topology type */
  chartType?: ({} & string) | PptxChartType;
  /** Category axis styling */
  catAxis?: PptxChartAxis;
  /** Value axis styling */
  valAxis?: PptxChartAxis;
  /** Axis text and line color */
  axisColor?: string;
  /** Default series colors array */
  colors?: string[];
  /** Data labels configuration */
  dataLabels?: PptxChartDataLabels;
  /** Gridline color */
  gridColor?: string;
  /** Bar/Column grouping */
  grouping?: 'clustered' | 'percentStacked' | 'stacked' | 'standard';
  /** Height */
  h?: Inches;
  /** Doughnut chart inner hole size percentage (10 to 90) */
  holeSize?: number;
  /** Custom element identifier */
  id?: string;
  /** Legend configuration */
  legend?: {
    color?: string;
    fontSize?: Points;
    overlay?: boolean;
    position?: 'bottom' | 'left' | 'right' | 'top' | 'topRight';
  };
  /** Custom element name */
  name?: string;
  /** Whether to show value axis gridlines */
  showGridlines?: boolean;
  /** Whether line charts are smoothed with splines */
  smooth?: boolean;
  /** Global chart text color */
  textColor?: string;
  /** Chart title text */
  title?: string;
  /** Width */
  w?: Inches;
  /** X coordinate */
  x?: Inches;
  /** Y coordinate */
  y?: Inches;
  /** Z-index order */
  zIndex?: number;
}
```

---

## Series Configuration (`ChartSeriesConfig`)

```typescript
export interface ChartSeriesConfig {
  /** Series display name */
  name?: string;
  /** Numeric values array */
  values: number[];
  /** Series solid color hex */
  color?: string;
  /** Multi-colors */
  colors?: string[];
  /** Per-datapoint slice color overrides (<c:dPt>) for Pie / Doughnut charts */
  dataPointColors?: string[];
  /** Custom PptxFill or color string */
  fill?: PptxFill | string;
}
```

---

## Supported Chart Topologies

| Type String | OpenXML Chart Element | Description |
| :--- | :--- | :--- |
| `'bar'` | `<c:barChart>` | Horizontal bar chart |
| `'col'` / `'column'` | `<c:barChart>` (`barDir="col"`) | Vertical column chart |
| `'line'` | `<c:lineChart>` | Line chart |
| `'area'` | `<c:areaChart>` | Area chart |
| `'pie'` | `<c:pieChart>` | Pie chart |
| `'doughnut'` / `'donut'` | `<c:doughnutChart>` | Doughnut chart with inner `holeSize` |
| `'radar'` | `<c:radarChart>` | Radar/Spider chart |
| `'scatter'` | `<c:scatterChart>` | XY scatter chart |
| `'bubble'` | `<c:bubbleChart>` | Bubble chart |
| `'stackedBar'` | `<c:barChart>` (`grouping="stacked"`) | Stacked bar chart |
| `'stackedColumn'` | `<c:barChart>` (`barDir="col" grouping="stacked"`) | Stacked column chart |
