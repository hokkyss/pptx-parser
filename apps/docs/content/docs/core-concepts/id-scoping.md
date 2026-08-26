---
title: "Unique ID Scoping & Lifecycle"
description: "Slide-scoped shape IDs, element indexing, conflict prevention, and element removal."
order: 4
section: "core-concepts"
---

# Unique ID Scoping & Lifecycle

In OpenXML DrawingML, every shape on a slide has a unique non-visual numeric ID (`<p:cNvPr id="2"/>`).

In `@hokkyss/pptx`, you can assign human-readable string IDs (such as `'service-card-1'` or `'auth-box'`) to any element. However, because string IDs are not natively supported by OpenXML, the library will discard it at the very end, and replace it with a unique numeric ID.

```typescript
// Assign an explicit ID to a shape
slide.addShape('roundRect', {
  id: 'service-card-1',
  x: inches(2),
  y: inches(2),
  w: inches(3),
  h: inches(2),
});

// Remove an element by its ID
slide.removeElement('service-card-1');
```


> [!IMPORTANT]
> Shape IDs are **slide-scoped**. Defining duplicate IDs on the same slide throws an explicit error at insertion time to prevent corrupted connector bindings.
