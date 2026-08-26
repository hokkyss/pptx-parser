---
title: "3-Tier Layer Inheritance"
description: "How the parser resolves master layouts, theme palettes, and placeholder styles."
order: 2
section: "parsing-and-mutation"
---

# 3-Tier Layer Inheritance

`@hokkyss/pptx-reader` and `@hokkyss/pptx` support the 3-tier OpenXML template inheritance chain:

1. **Slide Master (`ppt/slideMasters/slideMaster*.xml`)**: Contains the global DrawingML theme palette, default title/body text styles, and master canvas backgrounds.
2. **Slide Layout (`ppt/slideLayouts/slideLayout*.xml`)**: Inherits from a Slide Master and defines positioned placeholder geometries (`title`, `body`, `subTitle`, `pic`, `tbl`, `chart`).
3. **Slide Canvas (`ppt/slides/slide*.xml`)**: Inherits from a Slide Layout, referencing layout placeholders or overlaying custom visual shapes.
