---
title: "Groups & Transformations"
description: "Group elements together into composite OpenXML shapes with local coordinate transformations."
order: 11
section: "authoring"
---

# Groups & Transformations

Group multiple shapes, text frames, and lines into a single OpenXML group container (`<p:grpSp>`):

```typescript
import { inches } from '@hokkyss/pptx';

slide.addGroup({
  x: inches(1),
  y: inches(2),
  w: inches(8),
  h: inches(4),
}, (group) => {
  group.addShape('rect', {
    x: inches(0),
    y: inches(0),
    w: inches(8),
    h: inches(4),
    fill: 'F1F5F9',
  });
  group.addText('Grouped Card Heading', {
    x: inches(0.5),
    y: inches(0.5),
    w: inches(7),
    h: inches(1),
    bold: true,
  });
});
```
