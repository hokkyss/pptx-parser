---
title: "parsePptx"
description: "Binary parsing and OpenXML AST extraction options in @hokkyss/pptx-reader."
order: 1
package: "@hokkyss/pptx-reader"
section: "pptx-reader"
---

# parsePptx

The `parsePptx` function decompresses an OpenXML `.pptx` ZIP archive, parsing slide XMLs, DrawingML shapes, text bodies, tables, charts, embedded media assets, themes, and master layouts into a strongly-typed `PptxDocument` AST.

```typescript
import { parsePptx, type PptxParseOptions } from '@hokkyss/pptx-reader';

const doc = await parsePptx(fileBuffer, {
  includeMedia: true,
  lazyMedia: false,
  parseAnimations: true,
  parseTransitions: true,
  customXml: false
});
```

---

## Function Signature

```typescript
export async function parsePptx(
  input: ArrayBuffer | Uint8Array,
  options?: PptxParseOptions
): Promise<PptxDocument>
```

---

## Options (`PptxParseOptions`)

```typescript
export interface PptxParseOptions {
  /** Whether to extract media binary data. @default true */
  includeMedia?: boolean;
  /** If true, media data is loaded lazily via getter functions. @default false */
  lazyMedia?: boolean;
  /** Whether to parse animation timelines. @default true */
  parseAnimations?: boolean;
  /** Whether to parse slide transitions. @default true */
  parseTransitions?: boolean;
  /** Whether to preserve custom XML data parts. @default false */
  customXml?: boolean;
}
```

### Option Details

| Option | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `includeMedia` | `boolean` | `true` | When true, extracts binary images, video, and audio files from `ppt/media/` into `PptxMediaAsset[]`. |
| `lazyMedia` | `boolean` | `false` | When true, defers reading binary media bytes into memory until `asset.lazyGetter()` is called. Drastically reduces initial parse latency and memory pressure for large decks with high-resolution photos. |
| `parseAnimations` | `boolean` | `true` | Parses shape entrance, exit, and emphasis animation timelines from `<p:timing>`. |
| `parseTransitions` | `boolean` | `true` | Parses slide transition effects (`<p:transition>`) such as `fade`, `push`, `wipe`, and `dissolve`. |
| `customXml` | `boolean` | `false` | When true, collects auxiliary package parts (such as embedded Excel worksheets `ppt/embeddings/*.xlsx`, comments, and custom XML streams) to guarantee 100% round-trip fidelity. |
