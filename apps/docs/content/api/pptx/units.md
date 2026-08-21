---
title: "Units & Math"
package: "@hokkyss/pptx"
description: "Branded unit constructors and converter utilities."
---

# Units & Math

`@hokkyss/pptx` exports branded unit constructors and conversion utilities to prevent coordinate errors:

```typescript
import { inches, points, degrees, emu, toInches, toPoints, toDegrees, toEmu } from '@hokkyss/pptx';

const w = inches(4.5);
const pt = points(14);
const deg = degrees(90);
const emuVal = emu(914400);

const numInches = toInches(w); // 4.5
```
