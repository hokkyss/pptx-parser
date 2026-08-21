---
title: "LLM & RAG Ingestion"
description: "Extract clean structured text, tables, and presenter notes for AI embedding and search pipelines."
order: 4
section: "parsing-and-mutation"
---

# LLM & RAG Ingestion

Extract structured text, tables, and notes for vector databases, embedding models, and LLM context windows:

```typescript
import { parsePptx } from '@hokkyss/pptx-reader';

const doc = await parsePptx(fileBuffer);
const chunks = doc.slides.map(s => ({ slideNumber: s.slideNumber, elements: s.elements }));
```
