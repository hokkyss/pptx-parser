---
title: "SlideMaster & Layouts"
description: "API specification for SlideMaster, Layout discovery, and master theming in @hokkyss/pptx."
order: 7
package: "@hokkyss/pptx"
section: "pptx"
---

# SlideMaster & Layouts

The `SlideMaster` class represents an OpenXML Slide Master part (`ppt/slideMasters/slideMaster*.xml`). It provides methods to inspect associated layouts and create slides bound to specific master templates.

```typescript
const master = pres.getMaster('Office Theme');
const layouts = master.getLayouts();
```

---

## Properties

| Property | Type | Description |
| :--- | :--- | :--- |
| `id` | `string` | Slide Master ID (e.g. `'slideMaster1'`). |
| `name` | `string | undefined` | Slide Master name (e.g. `'Office Theme'`). |
| `ast` | `PptxSlideMaster` | Direct reference to underlying AST node. |
| `presentation` | `Presentation` | Presentation instance owning this master. |

---

## Methods

### `getLayouts()`

Returns an array of all `PptxSlideLayout` objects associated with this Slide Master.

```typescript
getLayouts(): PptxSlideLayout[]
```

---

### `getLayout(nameOrIdOrIndex)`

Finds a specific child layout by name, ID, or 1-based index.

```typescript
getLayout(nameOrIdOrIndex: number | string): PptxSlideLayout | undefined
```

---

### `addSlide(options?)`

Creates and returns a new `Slide` bound directly to this Slide Master.

```typescript
addSlide(options?: Omit<AddSlideOptions, 'master'>): Slide
```
