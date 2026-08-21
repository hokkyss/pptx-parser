---
title: "LLM & RAG Ingestion"
description: "Extract clean structured text, tables, and presenter notes for AI embedding and search pipelines."
order: 4
section: "parsing-and-mutation"
---

# LLM & RAG Ingestion

Extract structured text, tables, and speaker notes for vector databases, embedding models, and LLM context windows:

```typescript
import { parsePptx } from '@hokkyss/pptx-reader';

const doc = await parsePptx(fileBuffer, { includeMedia: false });

const chunks = doc.slides.map((s) => {
  const textRuns = s.elements
    .filter((el) => el.elementType === 'shape' && el.textBody)
    .flatMap((el) => el.textBody!.paragraphs.flatMap((p) => p.runs.map((r) => r.text)))
    .join(' ');

  return {
    slideNumber: s.slideNumber,
    notes: s.notes || '',
    text: textRuns,
  };
});
```
