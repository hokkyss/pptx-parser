---
title: "OpenXML Model & Hierarchy"
description: "Understand the underlying ECMA-376 OpenXML file layout and relationship graphs."
order: 1
section: "core-concepts"
---

# OpenXML Model & Hierarchy

PowerPoint presentations (`.pptx`) are ZIP packages conforming to the **ECMA-376 Office Open XML (OOXML)** standard.

### Package Directory Structure

```
presentation.pptx (ZIP Archive)
├── [Content_Types].xml            # MIME types for all package parts
├── _rels/.rels                    # Root package relationships
├── docProps/
│   ├── app.xml                    # Application properties (slides count)
│   └── core.xml                   # Dublin Core metadata (title, author, dates)
└── ppt/
    ├── presentation.xml           # Slide IDs, slide dimensions, firstSlideNumber
    ├── _rels/
    │   └── presentation.xml.rels  # Relationships to slide masters, themes, and slides
    ├── slideMasters/
    │   └── slideMaster1.xml       # Master templates and color schemes
    ├── slideLayouts/
    │   └── slideLayout1.xml       # Content placeholders and geometry
    ├── slides/
    │   ├── slide1.xml             # Visual shapes, text, tables, connectors
    │   └── _rels/
    │       └── slide1.xml.rels    # Relationships to charts, media, notes
    ├── charts/                    # Embedded DrawingML charts
    ├── media/                     # Embedded PNG/JPEG image assets
    └── theme/
        └── theme1.xml             # DrawingML color scheme & font schemes
```
