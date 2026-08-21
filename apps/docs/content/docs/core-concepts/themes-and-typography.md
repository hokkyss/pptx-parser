---
title: "Themes & Typography"
description: "DrawingML theme schemes (accent1-6, dk1/2, lt1/2, hlink) and Major/Minor font schemes."
order: 3
section: "core-concepts"
---

# Themes & Typography

OpenXML presentations define visual branding through DrawingML Themes (`ppt/theme/theme1.xml`).

## The 12 Theme Color Slots

```typescript
pres.setThemeColors({
  accent1: '#0284C7', // Primary Brand / Charts Series 1
  accent2: '#10B981', // Secondary Brand / Charts Series 2
  accent3: '#F59E0B', // Warning / Highlights
  accent4: '#EF4444', // Danger
  accent5: '#8B5CF6', // Purple Accent
  accent6: '#EC4899', // Pink Accent
  dk1: '#0F172A',     // Primary Dark Text (Slate 900)
  dk2: '#334155',     // Secondary Dark Text (Slate 700)
  lt1: '#FFFFFF',     // Canvas White Background
  lt2: '#F8FAFC',     // Card Surface Light (Slate 50)
  hlink: '#0284C7',   // Hyperlink Color
  folHlink: '#6366F1' // Visited Hyperlink Color
});
```

## Major & Minor Font Schemes

```typescript
pres.setThemeFonts({
  major: 'Inter',   // Heading Title typography
  minor: 'Roboto',  // Body content typography
});
```
