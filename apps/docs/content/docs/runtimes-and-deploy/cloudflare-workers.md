---
title: "Cloudflare Workers & V8 Isolates"
description: "Deploy serverless PPTX generation and parsing APIs on Cloudflare Workers with zero native dependencies."
order: 1
section: "runtimes-and-deploy"
---

# Cloudflare Workers & V8 Isolates

`@hokkyss/pptx` has zero Node.js native dependencies and runs directly inside Cloudflare Workers:

```typescript
import { Presentation, inches, points } from '@hokkyss/pptx';

export default {
  async fetch(request: Request): Promise<Response> {
    const pres = Presentation.create({ title: 'Edge Generated Deck' });
    const slide = pres.addSlide();
    slide.addText('Generated at Edge in 2ms', { fontSize: points(24) });

    const buffer = await pres.toArrayBuffer();

    return new Response(buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="edge_deck.pptx"',
      },
    });
  },
};
```
