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

To prevent coordinate unit math bugs, `@hokkyss/pptx` and `@hokkyss/pptx-core` provide TypeScript **Branded Types**:

```typescript
import { inches, points, degrees, emu } from '@hokkyss/pptx';

// Physical dimensions
const x = inches(1.5);       // Inches
const fontSize = points(16); // Points
const rotation = degrees(45);// Degrees
const rawEmu = emu(914400);  // 1 Inch = 914,400 EMUs
```

### Unit Conversion Table

| Unit | Value in EMUs | Constructor Function | To EMU Conversion | From EMU Conversion |
| :--- | :--- | :--- | :--- | :--- |
| **Inches** | `914,400` EMUs | `inches(val)` | `inchesToEmu(inchVal)` | `emuToInches(emuVal)` |
| **Points** | `12,700` EMUs | `points(val)` | `pointsToEmu(ptVal)` | `emuToPoints(emuVal)` |
| **Centimeters** | `360,000` EMUs | `centimeters(val)` | `cmToEmu(cmVal)` | `emuToCm(emuVal)` |
| **Pixels** | Depends on DPI (96) | `pixels(val)` | `pxToEmu(pxVal, 96)` | `emuToPx(emuVal, 96)` |
| **Degrees** | `60,000` EMU Degrees | `degrees(val)` | `degreesToEmuDegree(degVal)` | `rotationToDegrees(rotVal)` |
| **EMU** | `1` EMU | `emu(val)` | Identity | Identity |
| **HundredthsPoint** | `100` = 1 pt | `hundredthsPoint(val)` | `pointsToHundredthsPoint(pt)` | `hundredthsPointToPoints(cpt)` |
| **ThousandthsPercent** | `100,000` = 100% | `thousandthsPercent(val)` | `percentToThousandthsPercent(pct)` | `thousandthsPercentToPercent(tpct)` |
