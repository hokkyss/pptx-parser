---
title: "Text & Multilevel Bullets"
description: "Rich text formatting, hierarchical bullet lists (levels 0-8), superscript/subscript, alignment, and line spacing."
order: 3
section: "authoring"
---

# Text & Multilevel Bullets

`@hokkyss/pptx` supports both simple strings and advanced **multilevel hierarchical bullet matrices** with full Microsoft PowerPoint desktop parity:

- **Granular Indentation Levels (`0` to `8`)**: Tab / Shift+Tab cycling in PowerPoint desktop.
- **Shift + Enter Soft Line Breaks**: Soft line breaks (`<a:br>`) beneath a bullet without spawning a new bullet glyph.
- **PowerPoint Desktop Editing Parity**: Full 9-level tab stops registered across `<p:defaultTextStyle>` in `presentation.xml` and `<p:bodyStyle>` in `slideMaster1.xml`.
- **Text Boxes vs Content Placeholders**: Works seamlessly for both standalone text shapes and slide master layout placeholders (`<p:ph idx="..."/>`).

---

## Multilevel Hierarchical Lists

```typescript
import { points, inches } from '@hokkyss/pptx';

slide.addText([
  {
    // Level 0 (outermost bullet)
    level: 0,
    bullet: true,
    runs: [
      { text: 'Core Architecture Pillar: ', bold: true },
      { text: 'High Availability', bold: true, color: '0284C7' },
      // Shift+Enter behavior: soft line break within the same bullet
      { break: true },
      { text: '↳ Continuous operations across primary and secondary regions.' },
    ],
  },
  {
    // Level 1 (Tab - one level deeper)
    level: 1,
    bullet: true,
    text: 'Active-active multi-region replication across 3 availability zones.',
  },
  {
    // Level 2 (Tab Tab - two levels deeper)
    level: 2,
    bullet: true,
    runs: [
      { text: 'Latency: O(1) in-memory cache lookup profile: 10' },
      { text: '-6', superscript: true },
      { text: ' seconds' },
    ],
  },
  {
    // Level 0 (Shift+Tab back to root)
    level: 0,
    bullet: true,
    text: 'Secondary Architecture Pillar',
  },
], {
  x: inches(1),
  y: inches(1.5),
  w: inches(11.33),
  h: inches(5.0),
  lineSpacing: points(24),
  spaceBefore: points(8),
  spaceAfter: points(8),
});
```

---

## PowerPoint Desktop Keyboard Parity

### 1. `Shift + Enter` (Soft Line Breaks)

In PowerPoint desktop, pressing `Shift + Enter` inside a bulleted paragraph breaks to a new line **without spawning a new bullet point**.

In `@hokkyss/pptx`, specify `{ break: true }` within a paragraph's `runs` array. It is serialized as an OpenXML `<a:br>` element with preserved formatting runs:

```typescript
slide.addText([
  {
    level: 0,
    bullet: true,
    runs: [
      { text: 'First line with bullet glyph' },
      { break: true },
      { text: 'Second line indented cleanly under the same bullet' },
    ],
  },
]);
```

### 2. `Tab` and `Shift + Tab` (Promote & Demote Levels)

When opening generated decks in PowerPoint desktop, placing the cursor at the start of any bulleted line and pressing `Tab` increases the level deeper (0 ➔ 1 ➔ 2), while `Shift + Tab` decreases the level back toward the root (2 ➔ 1 ➔ 0).

`@hokkyss/pptx` achieves this by writing full 9-level OpenXML list level definitions (`lvl1pPr` through `lvl9pPr`) to:
1. **`<p:defaultTextStyle>` in `ppt/presentation.xml`**: Controls standalone text boxes and regular shapes.
2. **`<p:bodyStyle>` in `ppt/slideMasters/slideMaster1.xml`**: Controls Content Placeholders (`<p:ph idx="1"/>`).
3. **`<p:otherStyle>` in `ppt/slideMasters/slideMaster1.xml`**: Master-level non-placeholder text fallbacks.

---

## Bullet Types & Numbering

`@hokkyss/pptx` supports standard character bullets, custom glyphs, and auto-numbering schemes:

```typescript
// Standard character bullet (•)
slide.addText('Bulleted point', { bullet: true });

// Custom character bullet
slide.addText('Custom arrow bullet', {
  bullet: { type: 'char', char: '➔' },
});

// Auto-numbered list (1., 2., 3.)
slide.addText([
  { text: 'Phase 1: Ingestion', bullet: { type: 'autoNum', style: 'arabicPeriod' } },
  { text: 'Phase 2: Processing', bullet: { type: 'autoNum', style: 'arabicPeriod' } },
]);
```

---

## Content Placeholders vs Text Boxes

| Feature | Standalone Text Box (`slide.addText`) | Content Placeholder (`placeholder: 'body'`) |
| :--- | :--- | :--- |
| **OpenXML Element** | `<p:sp>` with `<p:cNvSpPr txBox="1"/>` | `<p:sp>` with `<p:ph idx="1"/>` |
| **Positioning** | Explicit `x, y, w, h` | Inherits coordinates from Slide Layout |
| **Style Source** | `<p:defaultTextStyle>` in `presentation.xml` | `<p:bodyStyle>` in `slideMaster1.xml` |
| **Tab / Shift+Tab** | Supported (via `defaultTextStyle` tab stops) | Supported (via `bodyStyle` layout inheritance) |
| **Soft Breaks** | Supported (`<a:br>`) | Supported (`<a:br>`) |


