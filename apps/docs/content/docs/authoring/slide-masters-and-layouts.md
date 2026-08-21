---
title: "Slide Masters & Layouts"
description: "Create reusable slide layouts with semantic placeholders (title, body, subtitle, picture)."
order: 2
section: "authoring"
---

# Slide Masters & Layouts

Slide Masters and Slide Layouts allow defining reusable templates with semantic placeholders that slides can instantiate and populate.

```typescript
// Add custom slide layout to presentation master
const master = pres.addSlideMaster({
  name: 'Corporate Master',
});

master.addLayout({
  name: 'master:title-with-body',
  placeholders: [
    { type: 'title', x: inches(1), y: inches(1), w: inches(11.33), h: inches(1) },
    { type: 'body', x: inches(1), y: inches(2.2), w: inches(11.33), h: inches(4.5) },
  ],
});

// Instantiate slide with that layout
const slide = pres.addSlide({
  layout: 'master:title-with-body',
});

// Populate placeholders
slide.addText('Annual Strategy', { placeholder: 'title' });
slide.addText('• Expansion into APAC region\n• Launching v2 API', { placeholder: 'body' });
```
