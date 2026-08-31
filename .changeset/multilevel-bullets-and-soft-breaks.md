---
"@hokkyss/pptx": minor
"@hokkyss/pptx-core": minor
"@hokkyss/pptx-reader": minor
"@hokkyss/pptx-writer": minor
---

Add native support for soft line breaks, hierarchical block indentation, 9-level default text styles, and clean Master Layout inheritance:
- Support `{ break: true }` in text runs to serialize native OpenXML `<a:br/>` tags within paragraphs without creating a new bullet point (equivalent to `Shift + Enter` in desktop PowerPoint).
- Support non-bulleted hierarchical indentation (`{ level: N, bullet: false }`) by serializing `<a:buNone/>` at indent levels $>0$, with clean parser recognition in `@hokkyss/pptx-reader`.
- Emit complete 9-level `<p:defaultTextStyle>` in `ppt/presentation.xml` with standard default tab sizes (`defTabSz="914400"`) and incremental margins for full Desktop PowerPoint Tab and Shift+Tab parity.
- Omit redundant `@_lvl="0"` and `@_algn="l"` on default level 0 paragraphs for clean Slide Master and Layout property inheritance.
- Remove hardcoded arbitrary serializer ID fallbacks in favor of centralized slide-level ID normalization.
