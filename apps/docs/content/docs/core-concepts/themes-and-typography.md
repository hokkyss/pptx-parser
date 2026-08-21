---
title: "Themes & Typography"
description: "Programmatic theme color palettes (dk1..lt2, accent1..6) and major/minor font scheme mappings."
order: 3
section: "core-concepts"
---

# Themes & Typography

DrawingML uses a 12-slot color palette scheme (`<a:clrScheme>`) and a Major/Minor font scheme (`<a:fontScheme>`) to maintain visual consistency across presentation decks.

```typescript
pres
  .setThemeName('Corporate Enterprise')
  .setThemeColors({
    accent1: '#0284C7', // Primary Brand Color
    accent2: '#6366F1', // Secondary Accent
    accent3: '#10B981', // Success / Positive
    accent4: '#F59E0B', // Warning / Metric
    accent5: '#EF4444', // Destructive / Alert
    accent6: '#8B5CF6', // Purple Accent
    dk1: '#0F172A',     // Primary Dark (Text on light background)
    dk2: '#1E293B',     // Secondary Dark
    lt1: '#FFFFFF',     // Primary Light (Text on dark background)
    lt2: '#F8FAFC',     // Secondary Light
  })
  .setThemeFonts({
    major: 'Inter',     // Headings & Slide Titles (+mj-lt)
    minor: 'Roboto',    // Body Copy & Tables (+mn-lt)
    name: 'Modern Sans',
  });
```
