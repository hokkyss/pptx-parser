# @hokkyss/pptx-writer

A high-performance, tree-shakeable, and isomorphic OpenXML PowerPoint (`.pptx`) archive serializer for TypeScript and JavaScript.

Converts structured Presentation ASTs (`PptxDocument`) into standard `.pptx` binary archives across Node.js, browsers, Cloudflare Workers, Deno, and Bun.

---

## Features

- ⚡ **Pure Isomorphic Architecture**: Pure serializers and functional closures without native dependencies. Fast compression powered by `fflate`.
- 🌲 **Tree-Shakeable**: Marked `"sideEffects": false` with pure functional serializers.
- 🎨 **Full OpenXML Serialization Coverage**:
  - **Themes**: Serializes `<a:theme>`, `<a:clrScheme>`, `<a:fontScheme>`, `<a:fmtScheme>` while updating theme names, colors, and typography.
  - **Slide Masters & Layouts**: Emits `<p:sldMaster>`, `<p:sldLayout>`, slide relationships, and `<[Content_Types].xml>`.
  - **Rich Text**: Paragraphs (`<a:p>`), list levels (`lvl="0..8"`), run properties (`<a:rPr>` for bold, italic, underline, strike, baseline / superscript / subscript, colors, and fonts).
  - **Shapes & Geometries**: Preset geometries (`rect`, `roundRect`, `ellipse`), solid/gradient fills, and stroke outlines.
  - **Tables**: Multi-column `<a:tbl>`, `<a:tr>`, `<a:tc>` with cell background fills, margins, and borders.
  - **Media & Pictures**: Embedded image assets (`<p:pic>`, `<a:blipFill>`).
- 🔄 **Perfect Round-Trip Symmetry**: Accurately serializes ASTs produced directly by `@hokkyss/pptx-reader`.
- 🛡️ **Lenient & Strict Modes**: Automatically generates fallback themes and layouts for minimal ASTs in `'lenient'` mode, or fails fast with validation messages in `'strict'` mode.

---

## Installation

```bash
pnpm add @hokkyss/pptx-writer
# or
npm install @hokkyss/pptx-writer
```

---

## Quick Start

### Serializing a Presentation AST to a Uint8Array

```typescript
import {
  writePptx,
  emu,
  hundredthsPoint,
  inchesToEmu,
  type PptxDocument,
} from '@hokkyss/pptx-writer';

const doc: PptxDocument = {
  customXml: [],
  media: [],
  metadata: {
    title: 'Quarterly Executive Review',
    creator: 'OpenXML Engine',
    slideCount: 1,
    slideWidth: inchesToEmu(13.333),
    slideHeight: inchesToEmu(7.5),
  },
  slideLayouts: [],
  slideMasters: [],
  slides: [
    {
      animations: [],
      slideId: 'rId2',
      slideNumber: 1,
      shapes: [],
      elements: [
        {
          elementType: 'shape',
          type: 'shape',
          id: '2',
          name: 'Title Box',
          isVisible: true,
          zIndex: 0,
          position: {
            x: emu(1000000),
            y: emu(1000000),
            cx: emu(9000000),
            cy: emu(1500000),
          },
          rotation: 0 as any,
          textBody: {
            bodyProperties: { verticalAlignment: 'middle' },
            paragraphs: [
              {
                properties: { alignment: 'center' },
                runs: [
                  {
                    text: 'High-Performance Edge Architecture',
                    properties: {
                      bold: true,
                      color: '0284C7',
                      fontSize: hundredthsPoint(3200),
                    },
                  },
                ],
              },
            ],
          },
        },
      ],
    },
  ],
  themes: [],
};

// Generates in-memory Uint8Array binary of .pptx archive
const buffer = await writePptx(doc);
```

---

## Options

`writePptx(document, options?)` accepts an optional configuration object:

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

---

## OpenXML Specifications & Schema Standards

Serializers in `@hokkyss/pptx-writer` conform directly to ECMA-376 and ISO/IEC 29500:

| OpenXML Concept | ECMA-376 / ISO Schema | Writer Implementation |
| :--- | :--- | :--- |
| **PresentationML Structure** | [PresentationML Document Schema](https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document) | `serializePresentation`, `serializeSlide`, `serializeSlideMaster`, `serializeSlideLayout` |
| **Themes & Custom Palettes** | [DrawingML `<a:theme>` & `<a:extraClrSchemeLst>`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.theme?view=openxml-3.0.1) | `serializeTheme` |
| **Theme Typography** | [DrawingML `<a:fontScheme>`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fontscheme?view=openxml-3.0.1) | `serializeTheme` (`majorFont`, `minorFont`, `typeface="+mj-lt"`, `typeface="+mn-lt"`) |
| **DrawingML Text & Bullets** | [DrawingML Text (`<a:txBody>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textbody?view=openxml-3.0.1) | `serializeTextBody`, `serializeParagraph`, `serializeRunProperties` |
| **Charts & Topologies** | [DrawingML Charts (`<c:chartSpace>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts?view=openxml-3.0.1) | `serializeChart` (Bar, Column, Line, Area, Pie, Doughnut, Radar, `<c:dPt>`) |
| **DrawingML Tables** | [DrawingML Tables (`<a:tbl>`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.table?view=openxml-3.0.1) | `serializeTable` |
| **Package Content Types** | [Open Packaging Conventions (OPC)](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk) | `serializeContentTypes`, `serializeRelationships` |

---

## License

MIT License. Copyright (c) 2026 hokkyss.
