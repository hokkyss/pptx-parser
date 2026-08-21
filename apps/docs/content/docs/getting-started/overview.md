---
title: "Overview"
description: "A high-performance, isomorphic, and type-safe PowerPoint (.pptx) toolkit for TypeScript and modern JavaScript runtimes."
order: 1
section: "getting-started"
---

# Overview

**`@hokkyss/pptx`** is a high-performance, isomorphic, and type-safe PowerPoint (`.pptx`) toolkit for TypeScript and modern JavaScript runtimes.

It provides a unified, zero-native-dependency ecosystem to **parse**, **construct**, **mutate**, **template**, and **serialize** OpenXML presentations with 100% round-trip fidelity across Node.js, Web Browsers, Cloudflare Workers, Deno, and Bun.

---

## Why @hokkyss/pptx?

Traditional JavaScript office document manipulation libraries suffer from critical limitations:
- **Monolithic Node-Only Bindings**: Relying on native C++ binaries (`libreoffice`, `unoconv`, or Node.js `fs` streams) that crash in serverless functions, V8 isolates, and browser environments.
- **Write-Only Generators**: Libraries like `PptxGenJS` can only create new presentations from scratch; they cannot parse, inspect, or mutate existing `.pptx` decks.
- **Loose Stringly-Typed APIs**: Passing raw numbers without units leads to hard-to-debug layout distortions across different presentation viewers.

`@hokkyss/pptx` solves these challenges from the ground up:

- ⚡ **100% Isomorphic & Zero Native Binaries**: Runs identically in Node.js (>= 18), modern browsers, Cloudflare Workers (V8 Isolates), AWS Lambda, Deno, and Bun. Powered by pure TypeScript codecs and `fflate`.
- 📐 **Branded Units & Type Safety**: Physical dimensions and coordinates are strictly typed via branded units (`Inches`, `Points`, `Degrees`, `Emu`), eliminating coordinate math errors at compile time.
- 🔄 **Perfect Round-Trip Symmetry**: Parse existing corporate decks into a structured AST, mutate text, swap chart series, or replace layout placeholders, and serialize back to standard ECMA-376 archives without loss of formatting.
- 🎨 **Enterprise Theming & Typography**: Native OpenXML DrawingML color schemes (`setThemeColors`), Major/Minor font pairs (`setThemeFonts`), and custom master layouts.
- 📊 **Native OpenXML Charts & Tables**: 8+ native chart topologies (Bar, Column, Line, Area, Doughnut, Pie, Radar) with per-datapoint customization (`<c:dPt>`) and multi-column styled tables.
- 🪢 **Shape Attachment & Glue Connectors**: Attach connectors directly to shape bounding box connection sites (`top`, `bottom`, `left`, `right`) so lines automatically follow shapes when moved in PowerPoint.

---

## Ecosystem Comparison

| Presentation Library | Minified Size | Gzip Size | Isomorphic (Browser + Edge + Node) | Tree-Shakeable | Round-Trip (Read + Write) | Branded Units |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: |
| **`@hokkyss/pptx`** | **190 KB** | **48.9 KB** | 🟢 **100% Yes** | 🟢 **Yes (Modular AST)** | 🟢 **Full AST Roundtrip** | 🟢 **Yes (`Inches`, `Points`)** |
| `pptxgenjs` | 358 KB | 108.4 KB | 🟡 Browser / Node only | 🔴 No (Monolithic) | 🔴 Write Only (No Parser) | 🔴 Untyped numbers |
| `officegen` | 1,240 KB | 385.0 KB | 🔴 Node.js Only (`fs`) | 🔴 No | 🔴 Write Only (No Parser) | 🔴 Untyped numbers |
| `python-pptx` | Python package | N/A | 🔴 Python runtime only | 🔴 No | 🟢 Full Roundtrip | 🟢 Python Length units |
