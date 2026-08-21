---
title: "OpenXML Model & Hierarchy"
description: "Understand the PresentationML and DrawingML structural hierarchy."
order: 1
section: "core-concepts"
---

# OpenXML Model & Hierarchy

PowerPoint presentations (`.pptx`) are structured under the ECMA-376 Office Open XML (OOXML) standard as a zipped package containing XML parts, media relationships, and content type manifests.

```
Presentation (.pptx archive)
├── ppt/presentation.xml (Deck metadata, slide size, slide master references)
├── ppt/slideMasters/slideMaster1.xml (Global theme styling, default placeholders)
├── ppt/slideLayouts/slideLayout1.xml (Layout templates: Title, Title+Body, Blank)
├── ppt/slides/slide1.xml (Slide shape tree, text, tables, connectors, charts)
├── ppt/theme/theme1.xml (Color scheme dk1..accent6, font scheme major/minor)
└── ppt/media/ (Images, SVGs, embedded audio/video)
```

### The 4-Tier Layer Inheritance Model
1. **Theme (`ppt/theme/theme1.xml`)**: Defines the color palette and typography fonts.
2. **Slide Master (`ppt/slideMasters/`)**: Inherits theme colors and defines default slide backgrounds and master placeholder geometries.
3. **Slide Layout (`ppt/slideLayouts/`)**: Inherits from Slide Master and defines layout types (e.g. `master:title-with-body`).
4. **Slide (`ppt/slides/`)**: Concrete slide instantiated by the user, inheriting layout defaults or overriding specific shape elements.
