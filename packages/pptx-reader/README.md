# @hokkyss/pptx-reader

A high-performance, isomorphic, and pure TypeScript OpenXML PowerPoint (`.pptx`) presentation parser for Node.js, Web Browsers, Cloudflare Workers, Deno, and Bun.

---

## Features

- ⚡ **Pure Isomorphic & Asynchronous**: Works seamlessly across all modern JavaScript runtimes with non-blocking async decompression powered by `fflate`.
- 🌲 **Tree-Shakeable**: Pure functional architecture with factory closures, zero native dependencies, and marked `"sideEffects": false`.
- 📐 **3-Tier Slide Layer Composition**: Resolves and composes OpenXML rendering layers (**Master** $\rightarrow$ **Layout** $\rightarrow$ **Slide**) in exact back-to-front depth order via `resolveSlideLayers()`.
- 🎨 **Theme & Typography Cascading**: Full OpenXML style cascade support, theme color/font schemes, and `<a:defRPr>` paragraph default run properties fallback.
- 🖼️ **Media Extraction**: Supports both eager binary extraction and zero-memory lazy getters (`lazyGetter` / `getData()`) for embedded images, videos, and audio.
- 📊 **Exhaustive OpenXML AST**: Strongly-typed discriminated unions for Shapes, Text Boxes, Tables, Charts, Group Shapes, Connectors, Animations, Transitions, Lock states (`isLocked`, `locks`), and Visibility (`isVisible`).
- 📏 **Branded Types & Units**: Interoperable with `@hokkyss/pptx-core` branded units (`Emu`, `Inches`, `Points`, `EmuDegree`).

---

## Installation

```bash
pnpm add @hokkyss/pptx-reader
# or
npm install @hokkyss/pptx-reader
```

---

## Quick Start

```typescript
import { parsePptx } from '@hokkyss/pptx-reader';

// Pass ArrayBuffer, Uint8Array, or Node Buffer
const buffer = await fetch('/presentation.pptx').then((res) => res.arrayBuffer());
const doc = await parsePptx(buffer);

console.log(`Title: ${doc.metadata.title}`);
console.log(`Total Slides: ${doc.slides.length}`);

for (const slide of doc.slides) {
  console.log(`Slide #${slide.slideNumber} has ${slide.elements.length} elements`);
}
```

---

## Parse Options

The main `parsePptx()` entry point accepts an options object:

```typescript
const doc = await parsePptx(buffer, {
  includeMedia: true,       // Extract media files (default: true)
  lazyMedia: true,          // Defer buffer allocation via lazy getters (default: false)
  parseAnimations: true,    // Parse animation sequences and timelines (default: true)
  parseTransitions: true,   // Parse slide transitions and durations (default: true)
  customXml: false,         // Preserve customXml/* parts (default: false)
});
```

| Option | Type | Default | Description |
|---|---|---|---|
| `includeMedia` | `boolean` | `true` | Whether to extract embedded media files from `ppt/media/`. |
| `lazyMedia` | `boolean` | `false` | When `true`, media binaries are not allocated in memory upfront and are loaded on-demand via `lazyGetter()`. |
| `parseAnimations` | `boolean` | `true` | Whether to parse slide animation sequences and timelines. |
| `parseTransitions` | `boolean` | `true` | Whether to parse slide transition effects and durations. |
| `customXml` | `boolean` | `false` | Whether to preserve custom XML data parts in `doc.customXml`. |

---

## Main Exports

All parser functions, layer resolvers, and utilities are exported directly from `@hokkyss/pptx-reader`:

```typescript
import {
  parsePptx,
  parseShapes,
  resolveSlideLayers,
  extractMedia,
  createZipReader,
  createXmlParser,
  createRelationshipResolver,
  createThemeResolver,
  parseTextBody,
  parseAnimations,
  parseTransition,
  parseChart,
  parseTable,
} from '@hokkyss/pptx-reader';
```

---

## 3-Tier Slide Layer Composition

PowerPoint renders slides in 3 distinct stacked visual layers:
1. **Master Layer** (`masterElements`): Background graphics and corporate branding from `slideMaster*.xml`.
2. **Layout Layer** (`layoutElements`): Template layout frames and placeholders from `slideLayout*.xml`.
3. **Slide Layer** (`slideElements`): Slide-specific foreground shapes, text, tables, and media from `slide*.xml`.

Use `resolveSlideLayers()` to compose all 3 layers in correct back-to-front rendering order:

```typescript
import { parsePptx, resolveSlideLayers } from '@hokkyss/pptx-reader';

const doc = await parsePptx(buffer);
const layers = resolveSlideLayers(doc, 1); // by 1-based slide number or slideId string

if (layers) {
  // Render elements in exact back-to-front visual order
  for (const element of layers.allElementsInRenderOrder) {
    console.log(`[${element.layerSource}] zIndex:${element.zIndex} name:"${element.name}" visible:${element.isVisible}`);
  }
}
```

---

## Discriminated Union AST

Every visual element in `slide.elements` (and `slide.shapes`) is a strongly-typed `PptxElement` discriminated union:

```typescript
import type { PptxElement } from '@hokkyss/pptx-reader';

function renderElement(element: PptxElement) {
  switch (element.elementType) {
    case 'shape':
      // Text box or AutoShape (element.textBody, element.isTextBox, element.geometry)
      break;

    case 'picture':
      // Picture element (element.picture.mediaId, element.blipEmbedId)
      break;

    case 'table':
      // Embedded Table (element.table.rows, element.table.columnWidths)
      break;

    case 'chart':
      // Embedded Chart (element.chart.series, element.chart.chartType, element.chart.categories)
      break;

    case 'group':
      // Group container (element.children recursively)
      break;

    case 'connector':
      // Connector line (element.line)
      break;
  }
}
```

---

## Working with Branded Units

OpenXML uses specialized integer units for coordinates, angles, and typography. Import type-safe branded unit helpers directly from `@hokkyss/pptx-reader`:

```typescript
import {
  emu,
  inches,
  points,
  emuToInches,
  inchesToEmu,
  emuToPoints,
  pointsToEmu,
  rotationToDegrees,
  degreesToRotation,
} from '@hokkyss/pptx-reader';

// Conversions
const widthInches = emuToInches(shape.position.cx); // Returns Inches (e.g. 10.0)
const emuVal = inchesToEmu(inches(2.5));            // Returns Emu (2,286,000)
const deg = rotationToDegrees(shape.rotation);       // Returns Degrees (90)
```

---

## CLI Inspector

A built-in inspection CLI script is included to quickly inspect presentation metadata, themes, layouts, masters, media files, and layer hierarchies:

```bash
pnpm inspect path/to/presentation.pptx
```

---

## OpenXML Specifications & Schema Standards

Parsing in `@hokkyss/pptx-reader` maps directly from ECMA-376 and ISO/IEC 29500 schemas into structured TypeScript ASTs:

| OpenXML Part | Schema & Specification Link | Parser Implementation |
| :--- | :--- | :--- |
| **Presentation & Structure** | [PresentationML Document Schema](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document) | `parser.ts`, `slide-parser.ts`, `slide-layout-parser.ts`, `slide-master-parser.ts` |
| **Themes & Color Schemes** | [DrawingML `<a:theme>`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.theme?view=openxml-3.0.1) | `theme-parser.ts` |
| **Shapes & Geometry** | [DrawingML Shapes (`<p:sp>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.shapeproperties?view=openxml-3.0.1) | `shape-parser.ts` |
| **DrawingML Text & Bullets** | [DrawingML Text (`<a:txBody>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textbody?view=openxml-3.0.1) | `text-parser.ts` |
| **DrawingML Tables** | [DrawingML Tables (`<a:tbl>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.table?view=openxml-3.0.1) | `table-parser.ts` |
| **DrawingML Charts** | [DrawingML Charts (`<c:chartSpace>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts?view=openxml-3.0.1) | `chart-parser.ts` |
| **Animations & Timing** | [PresentationML Timing (`<p:timing>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.timing?view=openxml-3.0.1) | `animation-parser.ts` |
| **Slide Transitions** | [PresentationML Transitions (`<p:transition>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.presentation.transition?view=openxml-3.0.1) | `transition-parser.ts` |

---

## License

MIT License. Copyright (c) 2026 hokkyss.
