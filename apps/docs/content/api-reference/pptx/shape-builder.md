---
title: "ShapeBuilder & addShape"
description: "API specification for vector shapes, gradient fills, borders, shadows, and text formatting in @hokkyss/pptx."
order: 4
package: "@hokkyss/pptx"
section: "pptx"
---

# ShapeBuilder & addShape

The `addShape` API adds geometric shapes to a slide canvas with fills (solid, gradient, transparent), border lines, outer drop shadows, and rich text bodies.

```typescript
import { inches, points, degrees } from '@hokkyss/pptx';

slide.addShape('roundRect', {
  x: inches(1),
  y: inches(2),
  w: inches(4),
  h: inches(2),
  fill: {
    type: 'linear',
    angle: degrees(45),
    stops: [
      { position: 0, color: '38BDF8' },
      { position: 1, color: '6366F1' }
    ]
  },
  line: {
    color: 'FFFFFF',
    width: inches(0.02),
    dashStyle: 'solid'
  },
  shadow: {
    blur: inches(0.1),
    distance: inches(0.05),
    direction: degrees(90),
    color: '000000',
    opacity: 0.25
  },
  text: 'Serverless Edge Node',
  textOptions: {
    bold: true,
    color: 'FFFFFF',
    fontSize: points(16),
    align: 'center'
  }
});
```

---

## Function Signatures

```typescript
// Slide canvas method
addShape(shapeType: string, options: AddShapeOptions): this

// Low-level builder function
export function buildShapeElement(
  shapeType: string,
  options: AddShapeOptions,
  counter?: number
): PptxShapeElement
```

---

## Shape Options (`AddShapeOptions`)

```typescript
export interface AddShapeOptions {
  /** X position */
  x: Inches;
  /** Y position */
  y: Inches;
  /** Width */
  w: Inches;
  /** Height */
  h: Inches;
  /** Solid hex, GradientFillInput, or PptxFill */
  fill?: FillInput;
  /** Border outline stroke styling */
  line?: {
    color?: string;
    dashStyle?: string;
    width?: Inches;
  };
  /** Rotation angle */
  rotation?: Degrees;
  /** Drop shadow styling */
  shadow?: ShapeShadowOptions;
  /** Plain text string, TextRunConfig array, or ParagraphConfig array */
  text?: ParagraphConfig[] | string | TextRunConfig[];
  /** Global text styling options */
  textOptions?: TextOptions;
  /** Clickable hyperlink URL or internal slide jump */
  hyperlink?: PptxHyperlink | string;
  /** Custom element identifier */
  id?: string;
  /** Custom element name */
  name?: string;
  /** 0-based z-index order */
  zIndex?: number;
}
```

---

## Fill Types (`FillInput`)

### 1. Hex Color String
```typescript
fill: '38BDF8'
```

### 2. Gradient Fill (`GradientFillInput`)
```typescript
export interface GradientFillInput {
  angle?: Degrees | number;
  flip?: 'none' | 'x' | 'xy' | 'y';
  pathBounds?: {
    bottom?: number;
    left?: number;
    right?: number;
    top?: number;
  };
  rotateWithShape?: boolean;
  stops: (GradientStopInput | string)[];
  type?: 'linear' | 'path' | 'radial';
}
```

### 3. Transparent Fill
```typescript
fill: { type: 'none' }
```

---

## Shadow Options (`ShapeShadowOptions`)

```typescript
export interface ShapeShadowOptions {
  alignment?: string;
  blur?: Inches;
  color?: string;
  direction?: Degrees;
  distance?: Inches;
  opacity?: number;
  rotateWithShape?: boolean;
}
```

---

## Supported Geometries

Supports all standard ECMA-376 OpenXML preset geometries including: `rect`, `roundRect`, `ellipse`, `triangle`, `diamond`, `parallelogram`, `trapezoid`, `hexagon`, `octagon`, `star4`, `star5`, `can`, `cube`, `heart`, `cloudCallout`, `wedgeRoundRectCallout`, and arrow presets.
