---
title: "Transitions & Speaker Notes"
description: "Configure slide transition effects (fade, wipe, push) and rich formatted presenter speaker notes."
order: 10
section: "authoring"
---

# Transitions & Speaker Notes

```typescript
// Slide transition effect
slide.setTransition('fade', {
  durationMs: 500,
  throughBlack: true,
});

// Rich formatted presenter notes
slide.setNotes([
  {
    runs: [
      { text: 'Key takeaway for investors: ', bold: true },
      { text: 'Sub-millisecond execution latency at scale.' },
    ],
  },
  {
    bullet: true,
    level: 0,
    text: 'Emphasize 100% isomorphic browser & edge support.',
  },
]);
```
