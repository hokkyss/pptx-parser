# @hokkyss/pptx

## 0.4.0

### Minor Changes

- 4ffe97b: feat(presentation): add configurable `firstSlideNumber` starting slide number
  
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

### Patch Changes

- Updated dependencies [4ffe97b]
  - @hokkyss/pptx-core@0.4.0
  - @hokkyss/pptx-reader@0.4.0
  - @hokkyss/pptx-writer@0.4.0

## 0.3.0

### Minor Changes

- 4106476: Release 0.3.0:
  
  - **Native OpenXML DrawingML Gradient Engine**: Linear, radial, multi-stop vector meshes, alpha transparency channels, and slide background gradients (`<a:gradFill>`, `<p:bgPr>`).
  - **Interactive Hyperlinks & Slide Navigation**: Web hyperlinks, internal slide index jumping, and native slide show action controls (`firstSlide`, `lastSlide`, `nextSlide`, `previousSlide`, `endShow`) with bidirectional security sanitizer.
  - **Layers & Z-Index Visual Composition**: Multi-tier visual layering with Master -> Layout -> Slide inheritance and 0-based z-index reordering.
  - **Native OpenXML Chart Topologies**: 13+ native chart topologies including clustered bar, stacked column, area, line, pie, and radar charts.
  - **CI Benchmarks & Bundlephobia Telemetry**: Automated sub-millisecond latency benchmarks and bundle budget monitoring.

### Patch Changes

- Updated dependencies [4106476]
  - @hokkyss/pptx-core@0.3.0
  - @hokkyss/pptx-reader@0.3.0
  - @hokkyss/pptx-writer@0.3.0
