---
title: "Node.js & Serverless"
description: "Run in Node.js >= 18, AWS Lambda, Netlify Functions, and Docker containers."
order: 3
section: "runtimes-and-deploy"
---

# Node.js & Serverless

Full support for Node.js 18, 20, 22+, AWS Lambda, and Netlify Functions:

```typescript
import fs from 'node:fs/promises';
import { Presentation, inches, points } from '@hokkyss/pptx';

const pres = Presentation.create({ title: 'Serverless Export' });
const slide = pres.addSlide();
slide.addText('Generated on Node.js Server', {
  x: inches(1),
  y: inches(1),
  fontSize: points(20),
});

const buffer = await pres.toBuffer();
await fs.writeFile('output.pptx', buffer);
```
