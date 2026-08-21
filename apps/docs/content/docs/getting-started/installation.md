---
title: "Installation"
description: "Install @hokkyss/pptx and configure TypeScript for your project."
order: 2
section: "getting-started"
---

# Installation

Install `@hokkyss/pptx` and its optional companion packages using your package manager of choice:

```bash
# Using pnpm (Recommended)
pnpm add @hokkyss/pptx

# Using npm
npm install @hokkyss/pptx

# Using yarn
yarn add @hokkyss/pptx

# Using bun
bun add @hokkyss/pptx
```

---

## Modular Tree-Shaking Packages

For edge runtimes or specialized workloads, you can install and import specific packages directly to minimize bundle footprint:

| Workload | Recommended Package | Minified / Gzip |
| :--- | :--- | :--- |
| **Full Presentation Authoring & Mutation** | `@hokkyss/pptx` | 190.9 KB / **48.9 KB** |
| **Pure OpenXML Parser (AST Ingestion / RAG)** | `@hokkyss/pptx-reader` | 95.4 KB / **25.8 KB** |
| **Pure Archive Serializer** | `@hokkyss/pptx-writer` | 125.9 KB / **31.6 KB** |
| **Branded Units & AST Schemas Only** | `@hokkyss/pptx-core` | 1.59 KB / **437 B** |
