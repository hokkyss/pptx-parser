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

// Speaker notes (supports plain text or structured runs)
slide.setNotes('Emphasize 100% isomorphic browser & edge support.');
```
