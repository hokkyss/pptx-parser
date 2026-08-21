---
title: "Coordinates & Branded Units"
description: "Master PowerPoint coordinate space, origin (0,0) Top-Left, and compile-time branded units (Inches, Points, Degrees, Emu)."
order: 2
section: "core-concepts"
---

# Coordinates & Branded Units

In OpenXML DrawingML and PowerPoint:
- **Origin `(0, 0)`** is at the **top-left** corner of the slide.
- **X axis**: Increases to the **right** ($+X$).
- **Y axis**: Increases **downward** ($+Y$).
- A shape's bounding box is placed at top-left `(x, y)` with width `w` and height `h`.

```
(0,0) Slide Top-Left
  ┌─────────────────────────────────────────────────────────────┐
  │                                                             │
  │     (x, y) ───────────── Top (x + w/2, y) ──────── (x + w, y)│
  │       │                                                │    │
  │  Left (x, y + h/2)          Center                Right (x + w, y + h/2)
  │       │                                                │    │
  │     (x, y + h) ──────── Bottom (x + w/2, y + h) ─── (x + w, y + h)
  │                                                             │
  └─────────────────────────────────────────────────────────────┘
                                                              (Slide Width, Slide Height)
```

---

## Branded Types

To prevent coordinate unit math bugs, `@hokkyss/pptx` utilizes TypeScript **Branded Types**:

```typescript
import { inches, points, degrees, emu } from '@hokkyss/pptx';

// Physical dimensions
const x = inches(1.5);      // Inches
const fontSize = points(16); // Points
const rotation = degrees(45);// Degrees
const rawEmu = emu(914400);  // 1 Inch = 914,400 EMUs
```

### Unit Conversion Table

| Unit | Value in EMUs | Constructor Function | Converter Utility |
| :--- | :--- | :--- | :--- |
| **Inches** | `914,400` EMUs | `inches(val)` | `toInches(u)` |
| **Points** | `12,700` EMUs | `points(val)` | `toPoints(u)` |
| **Degrees** | `60,000` EMU Degrees | `degrees(val)` | `toDegrees(u)` |
| **EMU** | `1` EMU | `emu(val)` | `toEmu(u)` |
