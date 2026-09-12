# @hokkyss/pptx-core

## 0.6.0

### Minor Changes

- 880c85d: fix(bullets): support multilevel bullets, block indentation, soft line breaks, and default text styles
  
  - **`@hokkyss/pptx`**:
    - Support `{ break: true }` in text runs to serialize native OpenXML `<a:br/>` tags within a paragraph without creating a new bullet point (equivalent to `Shift + Enter` in desktop PowerPoint).
    - Support non-bulleted hierarchical indentation (`{ level: N, bullet: false }`) by serializing `<a:buNone/>` at indent levels $>0$.
    - Added comprehensive unit tests for multilevel bullets, soft line breaks, and layout placeholder inheritance.
  - **`@hokkyss/pptx-writer`**:
    - Emits complete 9-level `<p:defaultTextStyle>` (`a:lvl1pPr` through `a:lvl9pPr`) in `ppt/presentation.xml` with standard default tab sizes (`defTabSz="914400"`) and incremental margins for complete Desktop PowerPoint Tab and Shift+Tab parity.
    - Added 9-level indent hierarchy across default slide master `txStyles` and default slide layout placeholders.
    - Standardized default bullet margins and hanging indents (`bulletGap = isNumbering ? 203200 : 152400` and `levelIndent = 228600`).
    - Removed hardcoded arbitrary serializer ID fallbacks in favor of centralized slide-level ID normalization.
    - Added dedicated test suite (`line-break.test.ts`) covering soft breaks and sequential run interleaving.
  - **`@hokkyss/pptx-reader`**:
    - Updated paragraph parser to recognize `<a:buNone/>` tags (`buNone !== undefined`), accurately reading non-bulleted hierarchical indentation.
    - Added unit test coverage for bullet-suppressed indented paragraphs.
- 5875aed: refactor: migrate write and read engines to 1st-party array-based `XmlElement` AST
  
  - **Zero External XML Dependencies**: Completely removed `fast-xml-parser` from the monorepo, replacing it with a custom zero-dependency isomorphic XML serializer and streaming recursive-descent parser.
  - **Ordered `XmlElement` AST in `@hokkyss/pptx-core`**: Introduced canonical, strongly-typed `XmlElement`, `XmlChild`, `XmlRaw`, and `XmlParser` interfaces shared across reader and writer.
  - **Sequential Document & Z-Order Preservation**:
    - Maintained exact visual z-order for shape trees (`<p:spTree>`, `<p:grpSp>`) across shapes, pictures, connectors, and graphic frames without tag-clustering.
    - Maintained exact interleaving order for paragraph children (`<a:p>`), supporting text runs (`<a:r>`), fields (`<a:fld>`), and soft line breaks (`<a:br>`).
  - **Full Type-Safety in Serializers**: All writer serializer modules now return strictly-typed `XmlElement` AST nodes instead of loosely-typed object trees.

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
