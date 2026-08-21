---
title: "Media & Images"
description: "Embed PNG, JPEG, SVG, WebP images from Buffer, Uint8Array, or Base64 with aspect ratio preservation."
order: 7
section: "authoring"
---

# Media & Images

Embed images seamlessly across Node.js, Web Browsers, and Cloudflare Workers:

```typescript
import { inches } from '@hokkyss/pptx';

slide.addImage({
  data: imageBytes,
  mimeType: 'image/png',
  x: inches(1),
  y: inches(1),
  w: inches(4),
  h: inches(2.5),
});
```
