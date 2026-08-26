---
title: "Bundle Sizing & Tree-Shaking"
description: "Modular import scenarios and minimal byte footprint breakdown."
order: 1
section: "performance"
---

# Bundle Sizing & Tree-Shaking

| Import Scenario | Bundle Import Path | Minified | Gzip Size | Brotli Size |
| :--- | :--- | :--- | :--- | :--- |
| **Branded Units & Contracts Only** | `import { inches } from '@hokkyss/pptx-core'` | 1.59 KB | **437 B** | 382 B |
| **Pure OpenXML Parser** | `import { parsePptx } from '@hokkyss/pptx-reader'` | 95.4 KB | **25.8 KB** | 22.8 KB |
| **Pure Archive Serializer** | `import { writePptx } from '@hokkyss/pptx-writer'` | 125.9 KB | **31.6 KB** | 27.7 KB |
| **Complete Presentation SDK** | `import { Presentation } from '@hokkyss/pptx'` | 190.9 KB | **48.9 KB** | 43.3 KB |
