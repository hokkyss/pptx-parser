---
"@hokkyss/pptx": minor
"@hokkyss/pptx-core": minor
"@hokkyss/pptx-reader": minor
"@hokkyss/pptx-writer": minor
---

Add customizable connector arrowheads, shape attachments, O(1) element lookup, and duplicate ID validation:
- Support `from` and `to` connector endpoints anchored directly to shapes via `position: 'top' | 'bottom' | 'left' | 'right'`.
- Support custom arrowhead styling (`endArrow`, `startArrow`, `headEnd`, `tailEnd`) with configurable type (`triangle`, `stealth`, `oval`, `diamond`, `open`, `none`), width (`sm`, `med`, `lg`), and length (`sm`, `med`, `lg`).
- Add O(1) element lookup via `slide.getElementById(id)` backed by an internal per-slide Map index.
- Add early validation bailing with a descriptive error when duplicate element IDs are added on the same slide.
- Map custom developer string IDs to valid OpenXML unsigned integers (`<p:cNvPr id="...">`, `<a:stCxn id="...">`, `<a:endCxn id="...">`) for seamless PowerPoint compatibility and zero repair prompts.
