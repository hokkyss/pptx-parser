---
title: "Architecture"
description: "Understand the 3-tier architecture, AST schemas, and round-trip fidelity guarantee."
order: 4
section: "getting-started"
---

# Architecture

`@hokkyss/pptx` is engineered around a **decoupled, three-tier compiler architecture** that separates the user-facing fluent authoring API from the underlying OpenXML DrawingML AST and zip archive codecs.

```
┌─────────────────────────────────────────────────────────────┐
│                       @hokkyss/pptx                        │
│         (High-Level Fluent SDK & Authoring API)             │
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
