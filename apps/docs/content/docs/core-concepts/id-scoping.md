---
title: "Unique ID Scoping & Lifecycle"
description: "Slide-scoped shape IDs, element indexing, conflict prevention, and attachment resolution."
order: 4
section: "core-concepts"
---

# Unique ID Scoping & Lifecycle

In OpenXML DrawingML, every shape on a slide has a unique numeric non-visual ID (`<p:cNvPr id="2"/>`).

In `@hokkyss/pptx`, you can assign human-readable string IDs (such as `'service-card-1'` or `'auth-box'`) to any element.

```typescript
// Assign an ID to a shape
slide.addShape('roundRect', {
  id: 'service-card-1',
  x: inches(2),
  y: inches(2),
  w: inches(3),
  h: inches(2),
});

// Attach a connector referencing that shape ID
slide.addConnector({
  from: { shapeId: 'service-card-1', position: 'right' },
  to: { shapeId: 'service-card-2', position: 'left' },
});
```

> [!IMPORTANT]
> Shape IDs are **slide-scoped**. Defining duplicate IDs on the same slide throws an explicit error at insertion time to prevent corrupted connector bindings.
