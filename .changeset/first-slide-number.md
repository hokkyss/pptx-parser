---
"@hokkyss/pptx-core": minor
"@hokkyss/pptx-reader": minor
"@hokkyss/pptx-writer": minor
"@hokkyss/pptx": minor
---

feat(presentation): add configurable `firstSlideNumber` starting slide number

Adds `firstSlideNumber` support across all packages:

- **`pptx-core`**: Added `firstSlideNumber?: number` to `PptxMetadata` interface (OpenXML `firstSlideNum` attribute on `<p:presentation>`).
- **`pptx-reader`**: Parses `@_firstSlideNum` from `ppt/presentation.xml` and populates `metadata.firstSlideNumber`.
- **`pptx-writer`**: Serializes `firstSlideNum="..."` attribute on `<p:presentation>` when set.
- **`pptx`**: Exposes `firstSlideNumber` in `CreatePresentationOptions`, as a readable/writable getter/setter, and as a fluent `setFirstSlideNumber(n)` method.

**Usage:**
```ts
// Option 1: at create time
const pres = Presentation.create({ firstSlideNumber: 0 });

// Option 2: fluent setter
pres.setFirstSlideNumber(0);

// Option 3: property setter
pres.firstSlideNumber = 0;
```
