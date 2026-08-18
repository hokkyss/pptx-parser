# @hokkyss/pptx

A modern, high-performance, and type-safe PowerPoint (`.pptx`) authoring and manipulation SDK for TypeScript and JavaScript.

Generate, load, mutate, style, and save presentations with a fluent API across Node.js, browsers, Cloudflare Workers, Deno, and Bun.

---

## Highlights

- 🚀 **Fluent Presentation API**: Chainable methods for creating slides, formatting text, drawing shapes, populating tables, and configuring themes.
- 🎨 **Theme & Typography Engine**: Programmatically customize theme color schemes (`setThemeColors`), font schemes (`setThemeFonts`), and theme display names (`setThemeName`).
- 📑 **Multilevel Bullet Lists**: Full support for hierarchical lists (levels 0 to 8) with automatic bullets, custom indents, and rich inline text modifiers (`bold`, `italic`, `underline`, `strikethrough`, `superscript`, `subscript`).
- 🎯 **Template & Placeholder Replacement**: Load corporate `.pptx` templates, instantiate slides from layouts (`master:cover-1`, `master:title-with-body`), and replace placeholder text by semantic type (`'title'`, `'body'`) or placeholder shape name.
- 📊 **Declarative Tables**: Create enterprise data tables with column widths, custom headers, alternating row styles, text alignment, and cell padding.
- 📐 **Slide CRUD & Re-Indexing**: Easily duplicate, move, and remove slides with automatic 1-based slide re-numbering.
- 🔄 **100% Round-Trip Fidelity**: High-fidelity serialization powered by `@hokkyss/pptx-writer` and isomorphic parsing powered by `@hokkyss/pptx-reader`.

---

## Installation

```bash
pnpm add @hokkyss/pptx
# or
npm install @hokkyss/pptx
```

---

## Quick Start

```typescript
import { Presentation, inches, points } from '@hokkyss/pptx';

// 1. Create a 16:9 widescreen presentation
const pres = Presentation.create({
  title: 'Cloud Architecture Deck',
  author: 'Platform Team',
});

// 2. Customize Theme Colors & Fonts
pres
  .setThemeName('Cloud Theme')
  .setThemeColors({
    accent1: '#0284C7', // Sky Blue
    accent2: '#6366F1', // Indigo
  })
  .setThemeFonts({
    major: 'Inter',
    minor: 'Roboto',
  });

// 3. Add Slide with Title & Multilevel Bullets
const slide = pres.addSlide();
slide.setBackground('0F172A'); // Dark Navy

slide.addText('Distributed Edge Architecture', {
  bold: true,
  color: '38BDF8',
  fontSize: points(28),
  x: inches(1),
  y: inches(1),
  w: inches(11.33),
  h: inches(0.8),
});

slide.addText([
  { level: 0, text: '1. Core Microservices Engine' },
  { level: 1, text: '1.1 Zero-Cold-Start Microsecond Execution' },
  {
    level: 2,
    runs: [
      { text: '1.1.1 Sub-millisecond P99 response time: O(n' },
      { text: '2', superscript: true },
      { text: ')' },
    ],
  },
  {
    level: 2,
    runs: [
      { text: '1.1.2 Legacy socket polling: ' },
      { text: 'deprecated', strikethrough: true },
    ],
  },
], {
  x: inches(1),
  y: inches(2),
  w: inches(11.33),
  h: inches(4.5),
});

// 4. Export binary buffer
const buffer = await pres.toBuffer(); // Uint8Array
const arrayBuffer = await pres.toArrayBuffer(); // ArrayBuffer
```

---

## API Reference

### 1. Presentation Lifecycle

#### `Presentation.create(options?)`
Creates a new blank presentation from scratch.
```typescript
const pres = Presentation.create({
  title: 'Quarterly Review',
  author: 'Jane Doe',
  width: inches(13.333), // Default: 16:9 widescreen (13.333 in)
  height: inches(7.5),    // Default: 7.5 in
});
```

#### `Presentation.load(buffer)`
Parses and loads an existing `.pptx` binary (`Uint8Array`, `ArrayBuffer`, or Node `Buffer`).
```typescript
const pres = await Presentation.load(bytes);
```

#### `pres.toBuffer()` & `pres.toArrayBuffer()`
Serializes the presentation in-memory to binary buffers for HTTP responses, S3/R2 uploads, or browser downloads.
```typescript
// Get binary Uint8Array (Cloudflare Workers, Express, Fastify, S3 uploads)
const uint8 = await pres.toBuffer();

// Get standard ArrayBuffer (Web APIs, Blob, Response)
const arrayBuffer = await pres.toArrayBuffer();
```

---

### 2. Theme Customization

#### `pres.setThemeColors(colors)`
Partially merges color scheme overrides into the primary theme. Accepts `#hex` or raw hex strings.
```typescript
pres.setThemeColors({
  accent1: '#0284C7',
  accent2: '#6366F1',
  accent3: '#10B981',
  accent4: '#F59E0B',
  accent5: '#EF4444',
  accent6: '#8B5CF6',
  dk1: '#0F172A',
  lt1: '#FFFFFF',
});
```

#### `pres.setThemeFonts(fonts)`
Configures headings (`major`), body (`minor`), and font scheme `name`:
```typescript
pres.setThemeFonts({
  major: 'Inter',
  minor: 'Roboto',
  name: 'Modern Web Scheme',
});
```

#### `pres.setThemeName(name)`
Sets the theme display name visible in PowerPoint's Theme/Color schemes picker:
```typescript
pres.setThemeName('Corporate Brand 2026');
```

---

### 3. Slide CRUD & Navigation

```typescript
// Add new slide (with optional layout and speaker notes)
const slide1 = pres.addSlide({
  layout: 'Title and Content', // layout name or layout ID
  notes: 'Key points to mention during presentation',
});

// Access slides
const total = pres.slides.length;
const firstSlide = pres.getSlide(1); // 1-based index
const byId = pres.getSlide('rId2');   // by slide ID

// Duplicate slide (deep-clones all elements, shapes, text, tables)
const duplicate = pres.duplicateSlide(1);

// Move slide position (1-based index)
pres.moveSlide(3, 1); // Move slide 3 to slide 1

// Remove slide (auto re-numbers subsequent slides)
pres.removeSlide(2);
```

---

### 4. Adding Text & Multilevel Hierarchies

#### Simple Text Box
```typescript
slide.addText('Executive Summary', {
  x: inches(1),
  y: inches(1),
  w: inches(8),
  h: inches(1),
  fontSize: points(24),
  bold: true,
  color: '0F172A',
  align: 'left', // 'left' | 'center' | 'right' | 'justify'
});
```

#### Multilevel Bullet Lists (Levels 0–8)
Pass an array of paragraph definitions with `level` (0 = main bullet, 1 = sub-bullet, etc.):
```typescript
slide.addText([
  { level: 0, text: 'Phase 1: Foundation' },
  { level: 1, text: 'Migrate infrastructure to Kubernetes' },
  { level: 1, text: 'Implement OpenTelemetry distributed tracing' },
  { level: 2, text: 'Configure Prometheus scrape targets' },
  { level: 0, text: 'Phase 2: Scale & Optimize' },
], {
  x: inches(1),
  y: inches(2),
  w: inches(10),
  h: inches(4),
});
```

#### Inline Formatting Modifiers
Use `runs` within paragraph definitions to apply granular formatting:
```typescript
slide.addText([
  {
    level: 0,
    runs: [
      { text: 'Formatted Text: ' },
      { text: 'Bold', bold: true },
      { text: ', ' },
      { text: 'Italic', italic: true },
      { text: ', ' },
      { text: 'Underlined', underline: true },
      { text: ', ' },
      { text: 'Strikethrough', strikethrough: true },
      { text: ', ' },
      { text: 'E = mc' },
      { text: '2', superscript: true },
      { text: ', and Chemical formula H' },
      { text: '2', subscript: true },
      { text: 'O.' },
    ],
  },
], {
  x: inches(1),
  y: inches(2),
  w: inches(10),
  h: inches(2),
});
```

---

### 5. Template Placeholder Replacement

When working with corporate templates, use `placeholder` in `addText()` to automatically inherit position, geometry, and styling from the slide layout:

```typescript
const pres = await Presentation.load(templateBuffer);

const slide = pres.addSlide({
  layout: 'master:title-with-body',
});

// Target title placeholder
slide.addText('Q3 Platform Architecture Review', {
  placeholder: 'title', // or exact name 'placeholder:slide-title'
});

// Target body placeholder
slide.addText([
  { level: 0, text: 'Key Milestone Deliverables' },
  { level: 1, text: 'API Gateway latency reduced by 40%' },
  { level: 1, text: 'Automated failover verified across 3 regions' },
], {
  placeholder: 'body', // or exact name 'placeholder:slide-content'
});
```

---

### 6. Data Tables

```typescript
slide.addTable(
  [
    ['Metric', 'Target SLA', 'Actual Q1', 'Variance'],
    ['P99 Latency', '< 50ms', '34ms', '-32%'],
    ['Throughput', '25,000 RPS', '32,100 RPS', '+28%'],
    ['Error Rate', '< 0.01%', '0.002%', '-80%'],
  ],
  {
    x: inches(1),
    y: inches(2),
    w: inches(11.33),
    h: inches(3),
    colWidths: [inches(3.5), inches(2.5), inches(2.5), inches(2.83)],
    header: {
      bold: true,
      fill: '0F172A',
      color: 'FFFFFF',
      fontSize: points(13),
    },
  },
);
```

---

### 7. Geometric Shapes & Connectors

```typescript
// Rounded Rectangle
slide.addShape('roundRect', {
  x: inches(1),
  y: inches(1),
  w: inches(3),
  h: inches(1.5),
  fill: '0284C7',
  text: 'Microservice API',
  color: 'FFFFFF',
  bold: true,
});

// Rectangle with Border
slide.addShape('rect', {
  x: inches(5),
  y: inches(1),
  w: inches(3),
  h: inches(1.5),
  fill: '1E293B',
  line: {
    color: '38BDF8',
    width: points(2),
  },
});

// Connector line with customizable start and end arrowheads
slide.addConnector({
  from: { x: inches(1), y: inches(3) },
  to: { x: inches(8), y: inches(3) },
  color: '0284C7',
  width: inches(0.03),
  endArrow: 'triangle', // or { type: 'stealth', width: 'lg', length: 'lg' }
  startArrow: { type: 'oval', width: 'med', length: 'med' },
  shapeType: 'bentConnector2', // 'line' | 'straightConnector1' | 'bentConnector2' | 'curvedConnector3'
});

// Attaching connectors directly to shapes (PowerPoint / Keynote glued connection sites)
slide.addShape('roundRect', { id: 'step-1', x: inches(1), y: inches(5), w: inches(3), h: inches(1.5), fill: '0284C7' });
slide.addShape('roundRect', { id: 'step-2', x: inches(6), y: inches(5), w: inches(3), h: inches(1.5), fill: '6366F1' });

slide.addConnector({
  from: { shapeId: 'step-1', position: 'right' }, // 'top' | 'bottom' | 'left' | 'right'
  to: { shapeId: 'step-2', position: 'left' },
  endArrow: 'triangle',
  color: '0284C7',
});
```

---

## Working with Units

```typescript
// Import physical units directly from @hokkyss/pptx
import { inches, points, px, emu } from '@hokkyss/pptx';

// Or conversion helpers from @hokkyss/pptx-core
import { inchesToEmu, pointsToEmu } from '@hokkyss/pptx-core';
```

---

## OpenXML Specifications & Schema Standards

The `@hokkyss/pptx` SDK builds presentation components against official ECMA-376 and ISO/IEC 29500 schemas:

| Component | Specification Reference | Schema & Elements |
| :--- | :--- | :--- |
| **Presentation & Master Hierarchy** | [PresentationML Document Structure](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document) | `Presentation`, `Slide`, `SlideMaster`, `SlideLayout` |
| **Theme & Color Customization** | [DrawingML `<a:theme>` & `<a:extraClrSchemeLst>`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.theme?view=openxml-3.0.1) | `setThemeName`, `setThemeColors` |
| **Typography & Font Schemes** | [DrawingML `<a:fontScheme>`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fontscheme?view=openxml-3.0.1) | `setThemeFonts` (`major` / `+mj-lt`, `minor` / `+mn-lt`) |
| **DrawingML Tables** | [DrawingML Table (`<a:tbl>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.table?view=openxml-3.0.1) | `TableBuilder`, `addTable` |
| **DrawingML Charts** | [DrawingML ChartSpace (`<c:chartSpace>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts?view=openxml-3.0.1) | `addChart` (Bar, Column, Line, Area, Pie, Doughnut, Radar) |

---

## License

MIT License. Copyright (c) 2026 hokkyss.
