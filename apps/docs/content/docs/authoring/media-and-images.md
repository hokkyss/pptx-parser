---
title: "Media & Images"
description: "Embed PNG and JPEG images from Uint8Array or ArrayBuffer with position and dimensions."
order: 7
section: "authoring"
---

# Media & Images

Embed images seamlessly across Node.js, Web Browsers, and Cloudflare Workers:

```typescript
import { inches } from '@hokkyss/pptx';

// imageBytes is a Uint8Array or ArrayBuffer
slide.addImage(imageBytes, {
  fileName: 'hero.png',
  x: inches(1),
  y: inches(1),
  w: inches(4),
  h: inches(2.5),
});
```
