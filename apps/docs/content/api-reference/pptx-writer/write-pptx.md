---
title: "writePptx"
description: "Specification and options for the writePptx OpenXML serializer in @hokkyss/pptx-writer."
order: 1
package: "@hokkyss/pptx-writer"
section: "pptx-writer"
---

# writePptx

The `writePptx` function serializes a `PptxDocument` AST into a strictly compliant ECMA-376 OpenXML `.pptx` ZIP archive represented as a `Uint8Array`.

```typescript
import { writePptx, type WritePptxOptions } from '@hokkyss/pptx-writer';

const uint8Array = await writePptx(ast, {
  mode: 'lenient'
});
```

---

## Function Signature

```typescript
export async function writePptx(
  document: PptxDocument,
  options?: WritePptxOptions
): Promise<Uint8Array>
```

---

## Options (`WritePptxOptions`)

```typescript
export interface WritePptxOptions {
  /**
   * Validation and fallback mode:
   * - 'lenient' (default): Automatically injects missing layouts, themes, IDs, and dimensions.
   * - 'strict': Fails fast with descriptive error messages on invalid AST structures.
   */
  mode?: 'lenient' | 'strict';
}
```

### Mode Behavior

| Option Value | Description |
| :--- | :--- |
| `'lenient'` (default) | Best for programmatic generation and dynamic templating. If a slide is missing a layout ID, master relationship, or explicit dimensions, the serializer automatically injects fallback defaults (`slideLayout1.xml`, `theme1.xml`, `13.333in × 7.5in`). |
| `'strict'` | Best for automated testing, continuous integration, and schema validation. Throws a runtime Error if the AST is missing mandatory nodes (e.g. empty slide array: `Strict mode error: Document must contain at least one slide.`). |
