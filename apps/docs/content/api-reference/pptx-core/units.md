---
title: "Units System"
description: "Branded physical units, identity constructors, and metric conversion functions in @hokkyss/pptx-core."
order: 2
package: "@hokkyss/pptx-core"
section: "pptx-core"
---

# Units System

PowerPoint's ECMA-376 OpenXML standard internally expresses layout metrics in **EMUs (English Metric Units)**, **EmuDegrees**, **Hundredths of a Point**, and **Thousandths of a Percent**.

The `@hokkyss/pptx-core` units module provides **compile-time Branded Types**, identity constructors, and high-precision conversion helpers to prevent coordinate drift and type mixing.

```typescript
import {
  emu,
  inches,
  points,
  centimeters,
  pixels,
  degrees,
  emuDegree,
  hundredthsPoint,
  thousandthsPercent,
  percent,
  inchesToEmu,
  emuToInches,
  pointsToEmu,
  emuToPoints,
  cmToEmu,
  emuToCm,
  pxToEmu,
  emuToPx,
  degreesToEmuDegree,
  rotationToDegrees,
  type Emu,
  type Inches,
  type Points,
  type Centimeters,
  type Pixels,
  type Degrees,
  type EmuDegree,
  type HundredthsPoint,
  type ThousandthsPercent,
  type Percent
} from '@hokkyss/pptx-core';
```

---

## Branded Identity Constructors

Identity constructors validate that inputs are finite numbers and tag them with their corresponding compile-time brand:

| Constructor | Branded Type | Description | Usage Example |
| :--- | :--- | :--- | :--- |
| `inches(val)` | `Inches` | Anglo-American physical inches | `inches(13.333)` |
| `points(val)` | `Points` | Typographic desktop publishing points (1/72 inch) | `points(18)` |
| `centimeters(val)` | `Centimeters` | Metric centimeters | `centimeters(5)` |
| `pixels(val)` | `Pixels` | Screen display pixels (default 96 DPI) | `pixels(1920)` |
| `degrees(val)` | `Degrees` | Standard rotational degrees (0° to 360°) | `degrees(45)` |
| `emu(val)` | `Emu` | Raw OpenXML English Metric Units | `emu(914400)` |
| `emuDegree(val)` | `EmuDegree` | OpenXML rotational angle units ($1^\circ = 60{,}000$) | `emuDegree(2700000)` |
| `hundredthsPoint(val)` | `HundredthsPoint` | OpenXML DrawingML font metric ($1\text{ pt} = 100$) | `hundredthsPoint(1800)` |
| `thousandthsPercent(val)` | `ThousandthsPercent` | OpenXML transparency/crop metric ($100\% = 100{,}000$) | `thousandthsPercent(50000)` |
| `percent(val)` | `Percent` | Standard percentage value (0 to 100) | `percent(50)` |

---

## Unit Conversion Functions

Convert between user-facing dimension units and OpenXML AST metrics:

### Inches & EMU ($1\text{ in} = 914{,}400\text{ EMU}$)
```typescript
const widthEmu = inchesToEmu(inches(13.333)); // 12,192,000 EMU
const widthInches = emuToInches(widthEmu);    // 13.333 Inches
```

### Points & EMU ($1\text{ pt} = 12{,}700\text{ EMU}$)
```typescript
const fontEmu = pointsToEmu(points(24));  // 304,800 EMU
const fontPt = emuToPoints(fontEmu);     // 24 Points
```

### Centimeters & EMU ($1\text{ cm} = 360{,}000\text{ EMU}$)
```typescript
const cardWidthEmu = cmToEmu(centimeters(10)); // 3,600,000 EMU
const cardWidthCm = emuToCm(cardWidthEmu);     // 10 Centimeters
```

### Pixels & EMU ($96\text{ DPI}$)
```typescript
const imgWidthEmu = pxToEmu(pixels(800), 96); // 7,620,000 EMU
const imgWidthPx = emuToPx(imgWidthEmu, 96);  // 800 Pixels
```

### Degrees & EmuDegrees ($1^\circ = 60{,}000\text{ EmuDegrees}$)
```typescript
const rotEmu = degreesToEmuDegree(degrees(90)); // 5,400,000 EmuDegrees
const rotDeg = rotationToDegrees(rotEmu);      // 90 Degrees
```

### Percentages & Thousandths of a Percent
```typescript
const tpct = percentToThousandthsPercent(percent(75)); // 75,000
const pct = thousandthsPercentToPercent(tpct);        // 75
const dec = thousandthsPercentToDecimal(tpct);        // 0.75
const tpctFromDec = decimalToThousandthsPercent(0.75); // 75,000
```
