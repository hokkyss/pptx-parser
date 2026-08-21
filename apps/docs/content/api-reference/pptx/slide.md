---
title: "Slide"
description: "Comprehensive API specification and method reference for the Slide class in @hokkyss/pptx."
order: 2
package: "@hokkyss/pptx"
section: "pptx"
---

# Slide

The `Slide` class models an individual presentation canvas. It wraps the OpenXML slide part (`ppt/slides/slide*.xml`) and provides a fluent, chainable API for placing text boxes, vector shapes, data tables, native charts, images, groups, and smart connectors.

```typescript
const slide = pres.addSlide();
```

---

## Slide Properties

| Property | Type | Description |
| :--- | :--- | :--- |
| `slideNumber` | `number` | 1-based chronological slide number within the deck. |
| `slideId` | `string` | Internal OpenXML relationship ID (e.g. `rId2`). |
| `layoutId` | `string \| undefined` | Associated slide layout ID (e.g. `slideLayout1`). |
| `notes` | `string \| undefined` | Plain text speaker notes. |
| `notesBody` | `PptxTextBody \| undefined` | Structured rich text body for speaker notes. |
| `ast` | `PptxSlide` | Direct reference to the underlying `PptxSlide` AST node. |

---

## Canvas Element Methods

### `addText(content, options?)`

Places a styled text box or populates a layout placeholder on the slide. Returns `this` for fluent chaining.

```typescript
addText(
  content: ParagraphConfig[] | string | TextRunConfig[],
  options?: AddTextOptions
): this
```

#### `AddTextOptions`

`AddTextOptions` extends all standard typography properties (`TextOptions`) with layout geometry and element metadata:

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `x` | `Inches` | `inches(1)` | X coordinate from left canvas edge. |
| `y` | `Inches` | `inches(1)` | Y coordinate from top canvas edge. |
| `w` | `Inches` | `inches(6)` | Text box width. |
| `h` | `Inches` | `inches(1)` | Text box height. |
| `color` | `string` | Theme `dk1` | Font color hex string (e.g. `'0F172A'`). |
| `fontSize` | `Points` | `points(18)` | Font size in points. |
| `font` | `string` | Theme font | Custom font family name (e.g. `'Inter'`, `'Arial'`). |
| `bold` | `boolean` | `false` | Bold font weight. |
| `italic` | `boolean` | `false` | Italic font style. |
| `underline` | `boolean | 'dash' | 'dbl' | 'dotted' | 'heavy' | 'sng' | 'wave'` | `false` | Underline decoration style. |
| `strikethrough` | `boolean | 'dblStrike' | 'sngStrike'` | `false` | Strikethrough decoration. |
| `subscript` | `boolean` | `false` | Subscript baseline shift. |
| `superscript` | `boolean` | `false` | Superscript baseline shift. |
| `align` | `'center' | 'justify' | 'left' | 'right'` | `'left'` | Horizontal text alignment. |
| `verticalAlignment` | `'bottom' | 'middle' | 'top'` | `'top'` | Vertical alignment inside container. |
| `lineSpacing` | `Points` | Auto | Explicit line spacing height in points. |
| `spaceBefore` | `Points` | `0` | Paragraph space before in points. |
| `spaceAfter` | `Points` | `0` | Paragraph space after in points. |
| `bullet` | `BulletInput` | `false` | Bullet point styling (`true`, `'number'`, `'none'`, custom character `'•'`, or `PptxBullet`). |
| `fill` | `PptxFill | string` | Transparent | Background fill for the text box shape. |
| `rotation` | `Degrees` | `degrees(0)` | Text box rotation angle. |
| `hyperlink` | `PptxHyperlink | string` | `undefined` | Clickable hyperlink URL or internal slide jump. |
| `placeholder` | `number | string` | `undefined` | Target layout placeholder index or type (`'title'`, `'body'`, `'subTitle'`, `'ctrTitle'`). |
| `id` | `string` | Auto-increment | Custom element identifier. |
| `name` | `string` | Auto | Custom element name. |
| `zIndex` | `number` | Sequential | 0-based z-index order. |

#### Example: Multi-Level Bullets & Rich Typography

```typescript
import { inches, points } from '@hokkyss/pptx';

slide.addText([
  {
    text: 'Key Architectural Principles',
    bold: true,
    fontSize: points(20),
    color: '0284C7'
  },
  {
    text: 'Zero-Native Codecs: 100% pure TypeScript',
    level: 0,
    bullet: true,
    fontSize: points(15)
  },
  {
    text: 'Runs identically in Node.js, Browsers, and Cloudflare Workers',
    level: 1,
    bullet: true,
    fontSize: points(13),
    color: '64748B'
  },
  {
    text: 'Branded Coordinates: Eliminates EMU math conversion bugs',
    level: 0,
    bullet: true,
    fontSize: points(15)
  }
], {
  x: inches(1),
  y: inches(1.5),
  w: inches(11.3),
  h: inches(4.5)
});
```

---

### `addShape(shapeType, options)`

Places a geometric vector shape onto the canvas with solid, gradient, or transparent fills, border line styling, drop shadows, and internal text.

```typescript
addShape(shapeType: string, options: AddShapeOptions): this
```

#### Example

```typescript
import { inches, points, degrees } from '@hokkyss/pptx';

slide.addShape('roundRect', {
  id: 'card-auth',
  x: inches(1),
  y: inches(2),
  w: inches(3.5),
  h: inches(2),
  fill: {
    type: 'linear',
    angle: degrees(45),
    stops: [
      { position: 0, color: '0284C7' },
      { position: 1, color: '0369A1' }
    ]
  },
  line: { color: 'FFFFFF', width: inches(0.02) },
  shadow: { blur: inches(0.1), distance: inches(0.05), opacity: 0.2 },
  text: 'Authentication Layer',
  textOptions: { bold: true, color: 'FFFFFF', align: 'center' }
});
```

---

### `addImage(imageData, options?)`

Embeds a binary picture asset (`Uint8Array` or `ArrayBuffer`) into the presentation archive and renders it on the slide.

```typescript
addImage(
  imageData: ArrayBuffer | Uint8Array,
  options?: AddImageOptions
): this
```

#### `AddImageOptions`

| Option | Type | Description |
| :--- | :--- | :--- |
| `x` | `Inches` | X coordinate on canvas. |
| `y` | `Inches` | Y coordinate on canvas. |
| `w` | `Inches` | Image display width. |
| `h` | `Inches` | Image display height. |
| `mediaId` | `string` | Custom media asset ID for de-duplication. |
| `fileName` | `string` | Target file name in `ppt/media/` (e.g. `'hero.png'`). |
| `alpha` | `ThousandthsPercent` | Transparency level ($100\% = 100{,}000$). |
| `rotation` | `Degrees` | Image rotation angle. |
| `placeholder` | `number | string` | Picture placeholder target. |
| `hyperlink` | `PptxHyperlink | string` | Clickable hyperlink URL. |

---

### `addTable(dataOrBuilder, options?)`

Renders an OpenXML table grid from a 2D matrix or a fluent `TableBuilder`.

```typescript
addTable(
  dataOrBuilder: ((builder: TableBuilder) => void) | TableBuilder | TableMatrix,
  options?: AddTableOptions
): this
```

#### Example

```typescript
import { inches } from '@hokkyss/pptx';

slide.addTable([
  ['Service', 'SLA', 'Status'],
  ['Edge Gateway', '99.99%', 'Operational'],
  ['Vector Index', '99.95%', 'Operational']
], {
  x: inches(1),
  y: inches(2),
  w: inches(11.3),
  colWidths: [inches(4.5), inches(3.4), inches(3.4)],
  header: { fill: '0F172A', color: 'FFFFFF', bold: true }
});
```

---

### `addChart(options)`

Embeds a native OpenXML DrawingML chart.

```typescript
addChart(options: AddChartOptions): this
```

#### Example

```typescript
import { inches } from '@hokkyss/pptx';

slide.addChart({
  chartType: 'bar',
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Revenue', values: [100, 140, 180, 220], color: '0284C7' }
  ],
  x: inches(1),
  y: inches(2),
  w: inches(11.3),
  h: inches(4.5),
  title: 'Quarterly Growth ($k)'
});
```

---

### `addConnector(options)`

Adds a vector connector line between two coordinate points.

```typescript
addConnector(options: AddConnectorOptions): this
```

#### `AddConnectorOptions`

| Option | Type | Description |
| :--- | :--- | :--- |
| `from` | `{ x: Inches; y: Inches }` | Start point coordinates. |
| `to` | `{ x: Inches; y: Inches }` | End point coordinates. |
| `color` | `string` | Stroke color hex string. |
| `width` | `Inches` | Stroke line width. |
| `dashStyle` | `string` | Dash style (`'solid'`, `'dash'`, `'dot'`, `'dashDot'`, `'lgDash'`). |
| `shapeType` | `string` | Connector geometry (`'line'`, `'straightConnector1'`, `'bentConnector3'`). |

---

### `addGroup(options, callback)`

Groups multiple child elements into a composite container with local coordinate transforms.

```typescript
slide.addGroup({
  x: inches(1),
  y: inches(1),
  w: inches(6),
  h: inches(4)
}, (group) => {
  group.addShape('rect', { x: inches(0), y: inches(0), w: inches(6), h: inches(4), fill: 'F1F5F9' });
  group.addText('Grouped Card Heading', { x: inches(0.5), y: inches(0.5), w: inches(5), h: inches(1), bold: true });
});
```

---

## Slide Settings & Animations

### `setBackground(fill)`

Sets the background fill for this slide canvas.

```typescript
slide.setBackground('0F172A'); // Solid hex
slide.setBackground({
  type: 'linear',
  angle: degrees(135),
  stops: [{ position: 0, color: '0F172A' }, { position: 1, color: '1E293B' }]
});
```

---

### `setNotes(notes, options?)`

Sets speaker notes for the slide with support for plain text or structured rich text runs.

```typescript
slide.setNotes('Discuss multi-region latency benchmarks during this slide.');
```

---

### `setTransition(transition, options?)`

Configures the slide presentation mode transition effect.

```typescript
slide.setTransition('fade', { durationMs: 500 });
slide.setTransition('wipe', { direction: 'right', speed: 'fast' });
slide.setTransition('push', { direction: 'up', advanceAfterMs: 3000 });
```
