---
title: "Parsing Presentations"
description: "Decompress and parse .pptx binaries into structured AST schemas."
order: 1
section: "parsing-and-mutation"
---

# Parsing Presentations

Use `@hokkyss/pptx-reader` to parse any PowerPoint `.pptx` binary into a strongly-typed OpenXML AST.

```typescript
import { parsePptx } from '@hokkyss/pptx-reader';

const doc = await parsePptx(arrayBuffer);

console.log('Title:', doc.metadata.title);
console.log('Slides:', doc.slides.length);
```
