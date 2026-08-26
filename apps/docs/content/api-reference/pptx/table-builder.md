---
title: "TableBuilder & addTable"
description: "API specification for OpenXML tables, styling, and cell matrixes in @hokkyss/pptx."
order: 5
package: "@hokkyss/pptx"
section: "pptx"
---

# TableBuilder & addTable

The `addTable` API generates structured OpenXML presentation tables with customizable column widths, cell fills, alignments, and typography.

```typescript
import { inches, points } from '@hokkyss/pptx';

slide.addTable([
  ['Feature', 'Standard', 'Enterprise'],
  ['Isomorphic Engine', 'Yes', 'Yes'],
  ['Sub-millisecond Parsing', 'Yes', 'Yes'],
  ['Round-trip Fidelity', '100%', '100%']
], {
  x: inches(1),
  y: inches(1.8),
  w: inches(11.3),
  h: inches(3.5),
  colWidths: [inches(4.5), inches(3.4), inches(3.4)],
  header: { fill: '0F172A', color: 'FFFFFF', bold: true }
});
```

---

## Function Signatures

```typescript
// Slide helper (returns this)
addTable(
  dataOrBuilder: ((builder: TableBuilder) => void) | TableBuilder | TableMatrix,
  options?: AddTableOptions
): this

// Fluent Builder
const table = new TableBuilder(options)
  .setColWidths([inches(3), inches(4)])
  .addRow({ cells: ['A1', 'B1'], h: inches(0.5) })
  .build();
```

---

## Table Options (`AddTableOptions`)

```typescript
export interface AddTableOptions {
  /** Array of column widths */
  colWidths?: Inches[];
  /** Alternative column definitions */
  columns?: Array<{ w?: Inches } | Inches>;
  /** Total table height */
  h?: Inches;
  /** Style defaults applied to all cells in the first row */
  header?: Partial<CellConfig>;
  /** Custom element identifier */
  id?: string;
  /** Custom element name */
  name?: string;
  /** Placeholder index or type */
  placeholder?: number | string;
  /** Total table width */
  w?: Inches;
  /** X coordinate */
  x?: Inches;
  /** Y coordinate */
  y?: Inches;
  /** Z-index rendering order */
  zIndex?: number;
}
```

---

## Cell Configuration (`CellConfig`)

Individual cells in the table matrix can be passed as strings or as `CellConfig` objects for fine-grained styling:

```typescript
export interface CellConfig {
  align?: 'center' | 'left' | 'right';
  bold?: boolean;
  color?: string;
  colSpan?: number;
  fill?: PptxFill | string;
  font?: string;
  fontSize?: Points;
  italic?: boolean;
  rowSpan?: number;
  text: ParagraphConfig[] | string | TextRunConfig[];
  verticalAlign?: 'bottom' | 'middle' | 'top';
}
```

### Example with Cell Overrides

```typescript
slide.addTable([
  ['Region', 'Q1 Revenue', 'Status'],
  ['North America', '$4.2M', { text: 'Exceeded', fill: 'DCFCE7', color: '166534', bold: true }],
  ['EMEA', '$3.1M', { text: 'On Track', fill: 'FEF3C7', color: '92400E' }]
], {
  x: inches(1),
  y: inches(2),
  colWidths: [inches(4), inches(3.5), inches(3.5)],
  header: { fill: '0F172A', color: 'FFFFFF', bold: true }
});
```
