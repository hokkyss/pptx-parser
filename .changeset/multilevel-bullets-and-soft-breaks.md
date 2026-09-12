---
"@hokkyss/pptx": minor
"@hokkyss/pptx-core": minor
"@hokkyss/pptx-reader": minor
"@hokkyss/pptx-writer": minor
---

fix(bullets): support multilevel bullets, block indentation, soft line breaks, and default text styles

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
