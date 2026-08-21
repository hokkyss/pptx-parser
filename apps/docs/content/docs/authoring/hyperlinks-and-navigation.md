---
title: "Hyperlinks & Slide Navigation"
description: "External web URLs, internal slide jumps ({ slideIndex }), and presentation action controls (firstSlide, nextSlide, endShow)."
order: 9
section: "authoring"
---

# Hyperlinks & Slide Navigation

Add interactive clickable links to shapes and text runs using OpenXML `<a:hlinkClick>`.

```typescript
// External web link
slide.addText('Visit Documentation', {
  hyperlink: { url: 'https://github.com/hokkyss/pptx-parser', tooltip: 'Open GitHub Repo' },
});

// Internal slide jump
slide.addShape('roundRect', {
  text: 'Go to Benchmark Slide (Slide 14)',
  hyperlink: { slideIndex: 14, tooltip: 'Jump to Slide 14' },
});

// Built-in slide show controls
slide.addShape('roundRect', {
  text: 'Next Slide ▶',
  hyperlink: { action: 'nextSlide' },
});
```
