---
title: "Presentation"
description: "Comprehensive API specification and method reference for the Presentation class in @hokkyss/pptx."
order: 1
package: "@hokkyss/pptx"
section: "pptx"
---

# Presentation

The `Presentation` class is the top-level entry point for programmatically creating, inspecting, mutating, and serializing PowerPoint presentations (`.pptx`).

```typescript
import { Presentation, inches } from '@hokkyss/pptx';

// 1. Create a new presentation
const pres = Presentation.create({
  title: 'Quarterly Business Review',
  author: 'Engineering Team',
  width: inches(13.333),
  height: inches(7.5)
});

// 2. Add slides and export
const slide = pres.addSlide();
slide.addText('Hello OpenXML', { x: inches(1), y: inches(1) });

const buffer = await pres.toBuffer();
```

---

## Static Factory Methods

### `Presentation.create(options?)`

Instantiates a new, blank presentation configured with a default Slide Master (`slideMaster1.xml`), Slide Layout (`slideLayout1.xml`), DrawingML Theme (`theme1.xml`), and document metadata.

```typescript
static create(options?: CreatePresentationOptions): Presentation
```

#### `CreatePresentationOptions`

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `title` | `string` | undefined | Document title stored in `docProps/core.xml`. |
| `author` | `string` | `'Pptx SDK'` | Creator and last modified author. |
| `company` | `string` | undefined | Company name in `docProps/app.xml`. |
| `width` | `Inches` | `inches(13.333)` | Slide canvas width ($12{,}192{,}000\text{ EMU}$). |
| `height` | `Inches` | `inches(7.5)` | Slide canvas height ($6{,}858{,}000\text{ EMU}$). |
| `firstSlideNumber` | `number` | `1` | Starting slide number in `ppt/presentation.xml` (set to `0` for cover slide decks). |

---

### `Presentation.load(input)`

Decompresses and parses an existing `.pptx` binary into a live, mutable `Presentation` instance.

```typescript
static async load(input: ArrayBuffer | Uint8Array): Promise<Presentation>
```

---

## Instance Properties

| Property | Type | Description |
| :--- | :--- | :--- |
| `ast` | `PptxDocument` | Direct reference to the root Abstract Syntax Tree. |
| `metadata` | `PptxMetadata` | Document metadata (title, author, slide count, dimensions). |
| `slides` | `Slide[]` | Array of all `Slide` instances in presentation order. |
| `firstSlideNumber` | `number \| undefined` | Starting slide number. |

---

## Slide Lifecycle Methods

### `addSlide(options?)`

Appends a new slide to the deck and returns its `Slide` instance.

```typescript
addSlide(options?: AddSlideOptions): Slide
```

#### `AddSlideOptions`

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `layout` | `string` | First master layout | Layout name (e.g. `'Title and Content'`, `'Blank'`) or ID (`'slideLayout1'`). |
| `layoutId` | `string` | undefined | Explicit layout ID. |
| `master` | `SlideMaster \| string` | Primary master | Associated master instance, master name, or ID. |
| `notes` | `string` | undefined | Plain text speaker notes. |

---

### `getSlide(indexOrId)`

Retrieves a slide by its 1-based index or internal OpenXML relationship ID (e.g. `'rId2'`).

```typescript
getSlide(indexOrId: number | string): Slide | undefined
```

---

### `removeSlide(indexOrId)`

Deletes a slide from the presentation and automatically re-indexes remaining slides. Returns `true` if deleted.

```typescript
removeSlide(indexOrId: number | string): boolean
```

---

### `duplicateSlide(indexOrId)`

Performs a deep clone of the target slide—including shapes, text runs, tables, charts, notes, and animations—and appends it to the deck.

```typescript
duplicateSlide(indexOrId: number | string): Slide
```

---

### `moveSlide(fromIndex, toIndex)`

Reorders a slide from one 1-based index position to another.

```typescript
moveSlide(fromIndex: number, toIndex: number): void
```

---

## Master & Theming Methods

### `getMasters()`

Returns all `SlideMaster` instances in the presentation.

```typescript
getMasters(): SlideMaster[]
```

---

### `getMaster(nameOrIdOrIndex)`

Finds a Slide Master by name, ID, or 1-based index.

```typescript
getMaster(nameOrIdOrIndex: number | string): SlideMaster | undefined
```

---

### `setThemeColors(colors, name?)`

Updates the 12 DrawingML color slots in `ppt/theme/theme1.xml`. All elements referencing theme tokens update automatically.

```typescript
setThemeColors(colors: ThemeColorInput, name?: string): this
```

#### `ThemeColorInput` Slots

| Slot | Semantic Role | OpenXML Tag | Default |
| :--- | :--- | :--- | :--- |
| `accent1` | Primary accent color | `<a:accent1>` | `2563EB` |
| `accent2` | Secondary accent color | `<a:accent2>` | `10B981` |
| `accent3` | Tertiary accent color | `<a:accent3>` | `F59E0B` |
| `accent4` | Warning accent color | `<a:accent4>` | `EF4444` |
| `accent5` | Info accent color | `<a:accent5>` | `8B5CF6` |
| `accent6` | Highlight accent color | `<a:accent6>` | `EC4899` |
| `dk1` | Primary dark text | `<a:dk1>` | `000000` |
| `dk2` | Secondary dark subtitles | `<a:dk2>` | `1F2937` |
| `lt1` | Main light canvas background | `<a:lt1>` | `FFFFFF` |
| `lt2` | Secondary light cards | `<a:lt2>` | `F3F4F6` |
| `hlink` | Hyperlink text | `<a:hlink>` | `2563EB` |
| `folHlink` | Visited hyperlink text | `<a:folHlink>` | `6D28D9` |

---

### `setThemeFonts(fonts)`

Configures the Major (headings) and Minor (body) font typography scheme.

```typescript
setThemeFonts(fonts: ThemeFontInput): this
```

---

### `setThemeName(name)`

Sets the display name of the primary theme.

```typescript
setThemeName(name: string): this
```

---

### `setFirstSlideNumber(num)`

Sets the starting slide number in `ppt/presentation.xml`.

```typescript
setFirstSlideNumber(num: number): this
```

---

## Serialization Methods

### `toBuffer(options?)`

Serializes the presentation into a binary `Uint8Array` archive.

```typescript
async toBuffer(options?: WritePptxOptions): Promise<Uint8Array>
```

---

### `toArrayBuffer(options?)`

Serializes the presentation into an `ArrayBuffer`.

```typescript
async toArrayBuffer(options?: WritePptxOptions): Promise<ArrayBuffer>
```

#### `WritePptxOptions`

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `mode` | `'lenient' \| 'strict'` | `'lenient'` | In `'lenient'` mode, automatically injects missing layouts and themes. In `'strict'` mode, throws errors on missing structures. |
