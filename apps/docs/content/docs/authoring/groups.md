---
title: "Groups & Transformations"
description: "Group elements together into composite OpenXML shapes with local coordinate transformations."
order: 11
section: "authoring"
---

# Groups & Transformations

Group multiple shapes, text frames, and lines into a single OpenXML group (`<p:grpSp>`):

```typescript
slide.addGroup({
  id: 'pipeline-group',
  x: inches(1),
  y: inches(2),
  w: inches(8),
  h: inches(4),
  elements: [],
});
```
