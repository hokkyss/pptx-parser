---
title: "AST Schemas"
description: "Comprehensive TypeScript schemas and discriminated union specifications for the OpenXML Abstract Syntax Tree in @hokkyss/pptx-core."
order: 1
package: "@hokkyss/pptx-core"
section: "pptx-core"
---

# AST Schemas

The `@hokkyss/pptx-core` package defines the universal Abstract Syntax Tree (AST) representing an ECMA-376 OpenXML presentation.

All visual canvas objects are modeled using a strict TypeScript **Discriminated Union** (`PptxElement`), discriminated by the `elementType` property (`'shape'`, `'picture'`, `'table'`, `'chart'`, `'connector'`, `'group'`).

```typescript
import type {
  PptxDocument,
  PptxSlide,
  PptxElement,
  PptxShapeElement,
  PptxPictureElement,
  PptxTableElement,
  PptxChartElement,
  PptxConnectorElement,
  PptxGroupElement
} from '@hokkyss/pptx-core';
```

---

## Root Container: `PptxDocument`

```typescript
export interface PptxDocument {
  /** Custom XML data parts. OpenXML: customXml/* */
  customXml: PptxCustomXmlPart[];
  /** Embedded media assets. OpenXML: ppt/media/* */
  media: PptxMediaAsset[];
  /** Document metadata. OpenXML: docProps/core.xml & ppt/presentation.xml */
  metadata: PptxMetadata;
  /** Reusable slide layout templates (ppt/slideLayouts/slideLayout*.xml) */
  slideLayouts: PptxSlideLayout[];
  /** Slide Masters defining layout hierarchies (ppt/slideMasters/slideMaster*.xml) */
  slideMasters: PptxSlideMaster[];
  /** Chronological array of presentation slides (ppt/slides/slide*.xml) */
  slides: PptxSlide[];
  /** Themes with color palettes and font schemes (ppt/theme/theme*.xml) */
  themes: PptxTheme[];
}
```

---

## Slide Model: `PptxSlide`

```typescript
export interface PptxSlide {
  /** 1-based slide number */
  slideNumber: number;
  /** Internal OpenXML relationship ID (e.g. 'rId2') */
  slideId: string;
  /** Reference to parent layout ID (e.g. 'slideLayout1') */
  layoutId?: string;
  /** Array of all visual canvas elements on this slide */
  elements: PptxElement[];
  /** Slide canvas background fill */
  background?: PptxBackground;
  /** Plain text speaker notes */
  notes?: string;
  /** Structured rich text body for speaker notes */
  notesBody?: PptxTextBody;
  /** Slide transition animation effect */
  transition?: PptxTransition;
  /** Shape entrance/exit animations */
  animations: PptxAnimation[];
  /** Shapes on the slide (alias for elements) */
  shapes: PptxShape[];
  /** Optional raw OpenXML content for exact round-trip fidelity */
  rawXml?: string;
  /** Optional raw relationship XML content (.rels) */
  relsXml?: string;
}
```

---

## Universal `PptxBaseElement` Properties

Every visual item on a slide canvas inherits these foundational properties:

```typescript
export interface PptxBaseElement {
  /** Unique element identifier (OpenXML: <p:cNvPr @_id>) */
  id: string;
  /** Human-readable element name (OpenXML: <p:cNvPr @_name>) */
  name: string;
  /** Position and dimensions in EMUs (OpenXML: <a:off @_x @_y>, <a:ext @_cx @_cy>) */
  position: {
    x: Emu;
    y: Emu;
    cx: Emu; // width
    cy: Emu; // height
  };
  /** Rotation angle in EmuDegrees (OpenXML: <a:xfrm @_rot>) */
  rotation: EmuDegree;
  /** 0-based z-index rendering order within container */
  zIndex: number;
  /** Whether the element is visible in PowerPoint (OpenXML: <p:cNvPr @_hidden="1">) */
  isVisible: boolean;
  /** Detailed lock settings (noMove, noResize, noRot, noSelect) */
  locks?: PptxShapeLocks;
  /** Clickable hyperlink relationship */
  hyperlink?: PptxHyperlink | string;
  /** Outer drop shadow effect (OpenXML: <a:outerShdw>) */
  shadow?: PptxShadow;
  /** Low-level OpenXML container tag */
  type: 'connector' | 'graphicFrame' | 'group' | 'picture' | 'shape';
}
```

---

## Detailed `PptxElement` Discriminated Union

```typescript
export type PptxElement =
  | PptxShapeElement      // elementType: 'shape'
  | PptxPictureElement    // elementType: 'picture'
  | PptxTableElement      // elementType: 'table'
  | PptxChartElement      // elementType: 'chart'
  | PptxConnectorElement  // elementType: 'connector'
  | PptxGroupElement;     // elementType: 'group'
```

---

### 1. `PptxShapeElement` (`elementType: 'shape'`)

Represents geometric vector AutoShapes and Text Boxes (`<p:sp>`).

```typescript
export interface PptxShapeElement extends PptxBaseElement {
  elementType: 'shape';
  type: 'shape';
  /** Preset geometry name (e.g. 'rect', 'roundRect', 'ellipse') */
  shapeType?: string;
  /** Geometry definition with adjustments or custom path commands */
  geometry?: PptxGeometry;
  /** Fill styling (solid color, linear gradient, radial gradient, transparent) */
  fill?: PptxFill;
  /** Border outline stroke (color, width, dash style) */
  line?: PptxLine;
  /** Formatted rich text body */
  textBody?: PptxTextBody;
  /** Whether the shape is explicitly configured as a pure text box (<p:cNvSpPr @_txBox="1">) */
  isTextBox?: boolean;
}
```

---

### 2. `PptxPictureElement` (`elementType: 'picture'`)

Represents an embedded bitmap image or vector picture (`<p:pic>`).

```typescript
export interface PptxPictureElement extends PptxBaseElement {
  elementType: 'picture';
  type: 'picture';
  /** Relationship ID of embedded media asset in ppt/media/ (OpenXML: <a:blip @_r:embed>) */
  blipEmbedId?: string;
  /** Picture fill and cropping parameters */
  picture: {
    mediaId: string;
    alpha?: ThousandthsPercent;
    blipFill?: Record<string, unknown>;
    crop?: {
      top?: ThousandthsPercent;
      bottom?: ThousandthsPercent;
      left?: ThousandthsPercent;
      right?: ThousandthsPercent;
    };
  };
}
```

---

### 3. `PptxTableElement` (`elementType: 'table'`)

Represents an OpenXML structured table grid rendered inside a graphic frame (`<p:graphicFrame><a:tbl>`).

```typescript
export interface PptxTableElement extends PptxBaseElement {
  elementType: 'table';
  type: 'graphicFrame';
  /** Structured table data */
  table: {
    /** Column widths in EMU */
    columnWidths: Emu[];
    /** Rows and constituent cells */
    rows: Array<{
      height: Emu;
      cells: Array<{
        colSpan?: number;
        rowSpan?: number;
        properties?: {
          fill?: PptxFill;
          verticalAlignment?: 'bottom' | 'middle' | 'top';
          topInset?: Emu;
          bottomInset?: Emu;
          leftInset?: Emu;
          rightInset?: Emu;
        };
        textBody?: PptxTextBody;
      }>;
    }>;
  };
}
```

---

### 4. `PptxChartElement` (`elementType: 'chart'`)

Represents an interactive DrawingML chart part rendered inside a graphic frame (`<p:graphicFrame><c:chartSpace>`).

```typescript
export interface PptxChartElement extends PptxBaseElement {
  elementType: 'chart';
  type: 'graphicFrame';
  /** Embedded chart configuration */
  chart: {
    /** Chart topology (bar, col, line, area, pie, doughnut, radar, scatter, etc.) */
    chartType: ({} & string) | PptxChartType;
    /** Category labels along category axis */
    categories: string[];
    /** Numerical data series */
    series: Array<{
      name: string;
      values: number[];
      index: number;
      order: number;
      fill?: PptxFill;
      /** Per-datapoint slice color overrides for Pie/Doughnut charts (<c:dPt>) */
      dataPointColors?: string[];
    }>;
    /** Chart title */
    title?: string;
    /** Legend settings */
    legend?: {
      color?: string;
      fontSize?: Points;
      overlay?: boolean;
      position?: 'bottom' | 'left' | 'right' | 'top' | 'topRight';
    };
    /** Doughnut inner hole size (10% to 90%) */
    holeSize?: number;
    /** Grouping style ('clustered' | 'stacked' | 'percentStacked' | 'standard') */
    grouping?: 'clustered' | 'percentStacked' | 'stacked' | 'standard';
    /** Category and Value axes */
    catAxis?: PptxChartAxis;
    valAxis?: PptxChartAxis;
    /** Data labels */
    dataLabels?: PptxChartDataLabels;
  };
}
```

---

### 5. `PptxConnectorElement` (`elementType: 'connector'`)

Represents an OpenXML connector line (`<p:cxnSp>`).

```typescript
export interface PptxConnectorElement extends PptxBaseElement {
  elementType: 'connector';
  type: 'connector';
  /** Stroke color, width, and dash style */
  line?: PptxLine;
}
```

---

### 6. `PptxGroupElement` (`elementType: 'group'`)

Represents a compound group shape container (`<p:grpSp>`) housing multiple nested child elements.

```typescript
export interface PptxGroupElement extends PptxBaseElement {
  elementType: 'group';
  type: 'group';
  /** Nested child elements rendered within the group's coordinate frame */
  children: PptxElement[];
}
```
