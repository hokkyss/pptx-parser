# @hokkyss/pptx-core

Core type definitions, OpenXML AST schemas, branded units, and conversion utilities powering the entire PPTX parser and generation ecosystem.

Zero runtime dependencies. Pure TypeScript. Marked `"sideEffects": false`.

---

## Overview

`@hokkyss/pptx-core` is the foundational contract layer shared across:
- **`@hokkyss/pptx-reader`** (parser)
- **`@hokkyss/pptx-writer`** (serializer)
- **`@hokkyss/pptx`** (high-level fluent SDK)

It provides compile-time type safety for OpenXML PresentationML/DrawingML data structures, branded physical and angular units, and discriminated unions for slide canvas elements.

---

## Installation

```bash
pnpm add @hokkyss/pptx-core
```

---

## 1. Type-Safe Branded Unit System (`/units`)

OpenXML expresses distances in **EMUs** (English Metric Units, $1\text{ in} = 914,400\text{ EMU}$), font sizes in **Hundredths of a Point** ($1\text{ pt} = 100\text{ hundredths}$), angles in **EMU Degrees** ($1^\circ = 60,000\text{ units}$), and alpha/opacity in **Thousandths of a Percent** ($100\% = 100,000$).

`@hokkyss/pptx-core` enforces compile-time branded types so raw numbers cannot be accidentally passed without explicit conversion:

```typescript
import {
  emu,
  inches,
  points,
  px,
  emuDegree,
  inchesToEmu,
  pointsToEmu,
  pxToEmu,
  emuToInches,
  emuToPoints,
  emuToPx,
} from '@hokkyss/pptx-core';

// Branded Unit Factories
const width = inches(13.333); // Inches
const height = inches(7.5);   // Inches
const fontSize = points(24);  // Points
const rotation = emuDegree(45); // EmuDegree (45°)

// Exact EMU Converters
const widthEmu = inchesToEmu(width); // Emu (12192000)
const fontEmu = pointsToEmu(fontSize); // Emu
const pxEmu = pxToEmu(px(100)); // Emu

// Reverse Converters
const inInches = emuToInches(widthEmu); // 13.333
const inPoints = emuToPoints(fontEmu);  // 24
const inPixels = emuToPx(pxEmu);        // 100
```

---

## 2. Universal Presentation AST (`/ast`)

The root `PptxDocument` interface models the full OpenXML presentation archive:

```typescript
import type {
  PptxDocument,
  PptxSlide,
  PptxElement,
  PptxShapeElement,
  PptxTableElement,
  PptxPictureElement,
  PptxGroupElement,
  PptxConnectorElement,
  PptxSlideMaster,
  PptxSlideLayout,
  PptxMetadata,
} from '@hokkyss/pptx-core';
```

### Discriminated Union: `PptxElement`

All visual elements on slides, layouts, and slide masters implement `PptxElement` with `elementType`:

```typescript
function renderElement(el: PptxElement) {
  switch (el.elementType) {
    case 'shape':
      // Text boxes, rectangles, geometric shapes
      console.log('Shape text:', el.textBody?.paragraphs.map(p => p.runs.map(r => r.text).join('')));
      break;

    case 'table':
      // Data table with rows, cells, columnWidths
      console.log(`Table: ${el.table.rows.length} rows x ${el.table.columnWidths.length} cols`);
      break;

    case 'picture':
      // Embedded bitmap/vector image with mediaId reference
      console.log('Image mediaId:', el.picture.mediaId);
      break;

    case 'connector':
      // Line or connector between shapes
      console.log('Connector shapeType:', el.shapeType);
      break;

    case 'group':
      // Group container containing nested PptxElement[]
      console.log(`Group with ${el.children.length} child elements`);
      break;

    case 'chart':
      // Embedded chart data
      console.log('Chart title:', el.chart.title);
      break;
  }
}
```

---

## 3. Rich Text & Multilevel Bullets (`/text`)

Models OpenXML `<p:txBody>`, `<a:p>`, `<a:pPr>`, `<a:r>`, and `<a:rPr>`:

```typescript
import type {
  PptxTextBody,
  PptxParagraph,
  PptxParagraphProperties,
  PptxTextRun,
  PptxTextRunProperties,
  PptxBullet,
} from '@hokkyss/pptx-core';

// Text Run Properties (Formatting Modifiers)
interface PptxTextRunProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  strikethrough?: boolean | string;
  baseline?: number; // >0 for superscript (e.g. 30000), <0 for subscript (e.g. -25000)
  superscript?: boolean;
  subscript?: boolean;
  fontSize?: HundredthsPoint;
  fontFamily?: string;
  color?: string; // Hex color (e.g. '0284C7')
}

// Paragraph Properties
interface PptxParagraphProperties {
  alignment?: 'left' | 'center' | 'right' | 'justify';
  level?: number; // 0 to 8 (0 = top-level bullet, 1 = sub-bullet, etc.)
  bullet?: PptxBullet; // Char, Auto-Numbered, or None
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
}
```

---

## 4. Theme & Color Schemes (`/theme`)

Models OpenXML `<a:theme>`, `<a:clrScheme>`, `<a:fontScheme>`, and theme modification inputs:

```typescript
import type {
  PptxTheme,
  PptxColorScheme,
  PptxFontScheme,
  PptxFormatScheme,
  ThemeColorInput,
  ThemeFontInput,
} from '@hokkyss/pptx-core';

// Partial color input supporting '#hex' and raw hex for all 12 slots:
const colors: ThemeColorInput = {
  accent1: '#0284C7',
  accent2: '#6366F1',
  accent3: '#10B981',
  accent4: '#F59E0B',
  accent5: '#EF4444',
  accent6: '#8B5CF6',
  dk1: '#0F172A',
  lt1: '#FFFFFF',
};

// Font scheme input:
const fonts: ThemeFontInput = {
  major: 'Inter',
  minor: 'Roboto',
  name: 'Enterprise Typography',
};
```

---

## Unified Main Export

All contracts, AST schemas, branded units, color definitions, and utilities are available via the main package entry point:

```typescript
import {
  // Units & Converters
  inches,
  points,
  emu,
  inchesToEmu,
  pointsToEmu,
  
  // AST Types
  type PptxDocument,
  type PptxSlide,
  type PptxElement,
  type PptxShapeElement,
  type PptxTableElement,
  type PptxPictureElement,

  // Theme & Colors
  type PptxTheme,
  type ThemeColorInput,
  type ThemeFontInput,

  // Text & Modifiers
  type PptxTextBody,
  type PptxParagraph,
  type PptxTextRun,
} from '@hokkyss/pptx-core';
```

---

## OpenXML Specifications & Schema Standards

AST contracts and units in `@hokkyss/pptx-core` are mapped 1-to-1 against ECMA-376 and ISO/IEC 29500 standards:

| Domain | Specification Reference | Schema & Elements |
| :--- | :--- | :--- |
| **PresentationML Package** | [ECMA-376 Part 1 (PresentationML)](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) | `PptxDocument`, `PptxSlide`, `PptxSlideLayout`, `PptxSlideMaster` |
| **Physical & Angular Units** | [OpenXML Unit Specification (EMU)](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk) | `Emu` (914,400/in), `Points` (12,700 EMU/pt), `EmuDegree` (60,000/deg) |
| **DrawingML Theme Engine** | [DrawingML Themes (`<a:theme>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.theme?view=openxml-3.0.1) | `PptxTheme`, `PptxColorScheme`, `PptxFontScheme`, `PptxFormatScheme` |
| **DrawingML Typography** | [DrawingML Font Schemes (`<a:fontScheme>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fontscheme?view=openxml-3.0.1) | `majorFont` (`+mj-lt`), `minorFont` (`+mn-lt`) |
| **DrawingML Shapes & Geometry** | [DrawingML Shapes (`<p:sp>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.shapeproperties?view=openxml-3.0.1) | `PptxShapeElement`, `PptxGeometry`, `PptxPosition` |
| **DrawingML Tables** | [DrawingML Tables (`<a:tbl>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.table?view=openxml-3.0.1) | `PptxTableElement`, `PptxTableRow`, `PptxTableCell` |
| **DrawingML Charts** | [DrawingML Charts (`<c:chartSpace>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts?view=openxml-3.0.1) | `PptxChart`, `PptxChartSeries`, `PptxChartType` |

---

## License

MIT License. Copyright (c) 2026 hokkyss.

