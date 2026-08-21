---
title: "Automated Template Populator"
description: "Batch-populate corporate PowerPoint templates with dynamic data."
category: "Automation"
difficulty: "Advanced"
---

# Automated Template Populator

```typescript
import { Presentation } from '@hokkyss/pptx';

export async function generateClientReport(templateBytes: Uint8Array, client: { name: string; score: number }) {
  const pres = await Presentation.load(templateBytes);
  const slide = pres.addSlide({ layout: 'master:client-summary' });

  slide.addText(client.name, { placeholder: 'title' });
  slide.addText(`Compliance Score: ${client.score}%`, { placeholder: 'body' });

  return await pres.toArrayBuffer();
}
```
