---
title: "Installation"
description: "Install @hokkyss/pptx and configure TypeScript for your project."
order: 2
section: "getting-started"
---

# Installation

Install `@hokkyss/pptx` and its optional companion packages using your package manager of choice:

:::::tabs-root{sync="package-manager" defaultValue="pnpm"}

::::tabs-list

:tab-item[pnpm]{value="pnpm"}
:tab-item[yarn]{value="yarn"}
:tab-item[npm]{value="npm"}
:tab-item[bun]{value="bun"}

::::

:::tab-content{value="pnpm"}
```bash
pnpm add @hokkyss/pptx
```
:::

:::tab-content{value="npm"}
```bash
npm install @hokkyss/pptx
```
:::

:::tab-content{value="yarn"}
```bash
yarn add @hokkyss/pptx
```
:::

:::tab-content{value="bun"}
```bash
bun add @hokkyss/pptx
```
:::

:::::

---

## Modular Tree-Shaking Packages

For edge runtimes or specialized workloads, you can install and import specific packages directly to minimize bundle footprint:

| Workload | Recommended Package | Minified / Gzip |
| :--- | :--- | :--- |
| **Full Presentation Authoring & Mutation** | `@hokkyss/pptx` | 190.9 KB / **48.9 KB** |
| **Pure OpenXML Parser (AST Ingestion / RAG)** | `@hokkyss/pptx-reader` | 95.4 KB / **25.8 KB** |
| **Pure Archive Serializer** | `@hokkyss/pptx-writer` | 125.9 KB / **31.6 KB** |
| **Branded Units & AST Schemas Only** | `@hokkyss/pptx-core` | 1.59 KB / **437 B** |
