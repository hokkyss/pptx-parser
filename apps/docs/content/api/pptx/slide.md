---
title: "Slide"
package: "@hokkyss/pptx"
description: "API Reference for Slide methods and element constructors."
---

# Slide

The `Slide` class represents a single slide canvas and provides methods to add shapes, text, tables, connectors, charts, media, and transitions.

## Methods

### `.addText(content: string | TextItem[], options?: TextOptions): this`
Adds a single text run or multilevel bullet hierarchy to the slide.

### `.addShape(type: ShapeType, options?: ShapeOptions): this`
Adds a vector shape geometry (e.g. `'roundRect'`, `'ellipse'`, `'triangle'`) with solid/gradient fill, borders, and shadows.

### `.addConnector(options: AddConnectorOptions): this`
Adds a connector line attached between shapes or absolute coordinates.

### `.addTable(options: TableOptions): this`
Adds a multi-column styled data table.

### `.addChart(type: ChartType, options: ChartOptions): this`
Adds an OpenXML native chart.

### `.addImage(options: ImageOptions): this`
Embeds an image from binary data or Base64.

### `.setBackground(fill: string | GradientOptions): this`
Sets the slide background color or gradient.

### `.setTransition(type: TransitionType, options?: TransitionOptions): this`
Sets slide show transition effect and duration.

### `.setNotes(notes: string | SpeakerNoteItem[]): this`
Sets rich presenter speaker notes.
