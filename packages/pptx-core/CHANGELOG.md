# @hokkyss/pptx-core

## 0.5.0

### Minor Changes

- 7876d29: Add customizable connector arrowheads, shape attachments, O(1) element lookup, and duplicate ID validation:
  - Support `from` and `to` connector endpoints anchored directly to shapes via `position: 'top' | 'bottom' | 'left' | 'right'`.
  - Support custom arrowhead styling (`endArrow`, `startArrow`, `headEnd`, `tailEnd`) with configurable type (`triangle`, `stealth`, `oval`, `diamond`, `open`, `none`), width (`sm`, `med`, `lg`), and length (`sm`, `med`, `lg`).
  - Add O(1) element lookup via `slide.getElementById(id)` backed by an internal per-slide Map index.
  - Add early validation bailing with a descriptive error when duplicate element IDs are added on the same slide.
  - Map custom developer string IDs to valid OpenXML unsigned integers (`<p:cNvPr id="...">`, `<a:stCxn id="...">`, `<a:endCxn id="...">`) for seamless PowerPoint compatibility and zero repair prompts.
- 4c93296: Add dynamic `./version` subpath export across all published packages (`@hokkyss/pptx`, `@hokkyss/pptx-core`, `@hokkyss/pptx-reader`, `@hokkyss/pptx-writer`) to provide type-safe runtime access to library version information.

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

## 0.3.0

### Minor Changes

- 4106476: Release 0.3.0:
  
  - **Native OpenXML DrawingML Gradient Engine**: Linear, radial, multi-stop vector meshes, alpha transparency channels, and slide background gradients (`<a:gradFill>`, `<p:bgPr>`).
  - **Interactive Hyperlinks & Slide Navigation**: Web hyperlinks, internal slide index jumping, and native slide show action controls (`firstSlide`, `lastSlide`, `nextSlide`, `previousSlide`, `endShow`) with bidirectional security sanitizer.
  - **Layers & Z-Index Visual Composition**: Multi-tier visual layering with Master -> Layout -> Slide inheritance and 0-based z-index reordering.
  - **Native OpenXML Chart Topologies**: 13+ native chart topologies including clustered bar, stacked column, area, line, pie, and radar charts.
  - **CI Benchmarks & Bundlephobia Telemetry**: Automated sub-millisecond latency benchmarks and bundle budget monitoring.
