---
title: "Presentation"
package: "@hokkyss/pptx"
description: "API Reference for the Presentation class in @hokkyss/pptx."
---

# Presentation

The `Presentation` class manages presentation-wide settings, masters, themes, and serialization.

## Static Methods

### `Presentation.create(options?): Presentation`
Initializes a new blank presentation deck.

| Parameter | Type | Description |
| :--- | :--- | :--- |
| `options.title` | `string` | Deck title metadata |
| `options.author` | `string` | Author metadata |
| `options.company` | `string` | Company metadata |
| `options.revision` | `number` | Document revision number |

### `Presentation.load(buffer: ArrayBuffer | Uint8Array): Promise<Presentation>`
Parses and loads an existing `.pptx` binary into a fluent Presentation instance.

## Instance Methods

### `.addSlide(options?): Slide`
Adds and returns a new slide in the presentation.

### `.setThemeColors(colors: Partial<PptxThemeColors>): this`
Configures presentation theme color scheme (`accent1`..`accent6`, `dk1`, `dk2`, `lt1`, `lt2`).

### `.setThemeFonts(fonts: PptxThemeFonts): this`
Configures presentation typography font scheme (`major` and `minor`).

### `.save(filePath: string): Promise<void>`
Serializes and writes the presentation to disk (Node.js runtimes).

### `.toArrayBuffer(): Promise<ArrayBuffer>`
Serializes the presentation to a binary `ArrayBuffer` (Node, Browser, Cloudflare Workers).

### `.toAst(): PptxPresentation`
Returns the internal strongly-typed OpenXML AST.
