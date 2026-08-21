---
title: "Architecture"
description: "Isomorphic design, memory efficiency, and layered AST architecture."
order: 4
section: "getting-started"
---

# Architecture

The architecture of `@hokkyss/pptx` is built around a layered, modular pipeline designed for high performance and zero native dependencies.

```
┌─────────────────────────────────────────────────────────────┐
│                      @hokkyss/pptx                          │
│     (Fluent Presentation, Slide, Layout, Builder APIs)      │
└──────────────┬───────────────────────────────┬──────────────┘
               │                               │
               ▼                               ▼
┌──────────────────────────────┐ ┌────────────────────────────┐
│    @hokkyss/pptx-reader     │ │    @hokkyss/pptx-writer   │
│   (Isomorphic OpenXML Parser)│ │   (Archive XML Serializer) │
└──────────────┬───────────────┘ └─────────────┬──────────────┘
               │                               │
               └───────────────┬───────────────┘
                               ▼
┌─────────────────────────────────────────────────────────────┐
│                     @hokkyss/pptx-core                     │
│    (Universal AST Schemas, Branded Units, Types & Contracts)│
└─────────────────────────────────────────────────────────────┘
```

### 100% Round-Trip Fidelity Guarantee

Any presentation parsed by `@hokkyss/pptx-reader` can be passed directly to `@hokkyss/pptx-writer` to regenerate a bit-for-bit valid presentation without losing layout formatting.
