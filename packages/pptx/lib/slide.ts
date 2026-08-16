import type {
  PptxConnectorElement,
  PptxElement,
  PptxFill,
  PptxGroupElement,
  PptxPictureElement,
  PptxShapeElement,
  PptxSlide,
} from '@hokkyss/pptx-core';
import {
  degreesToEmuDegree,
  emu,
  emuDegree,
  emuToInches,
  type Degrees,
  type Inches,
  inchesToEmu,
  type ThousandthsPercent,
} from '@hokkyss/pptx-core';
import {
  AddChartOptions,
  buildChartElement,
} from './builders/chart-builder';
import {
  AddShapeOptions,
  buildShapeElement,
  normalizeFill,
} from './builders/shape-builder';
import {
  AddTableOptions,
  TableBuilder,
  type TableMatrix,
} from './builders/table-builder';
import {
  buildTextBody,
  type ParagraphConfig,
  type TextOptions,
  type TextRunConfig,
} from './builders/text-builder';
import type { Presentation } from './presentation';

export interface AddTextOptions extends TextOptions {
  fill?: PptxFill | string;
  h?: Inches;
  id?: string;
  name?: string;
  placeholder?: number | string;
  rotation?: Degrees;
  w?: Inches;
  x?: Inches;
  y?: Inches;
  zIndex?: number;
}

export interface AddImageOptions {
  alpha?: ThousandthsPercent;
  fileName?: string;
  h?: Inches;
  id?: string;
  mediaId?: string;
  name?: string;
  placeholder?: number | string;
  rotation?: Degrees;
  w?: Inches;
  x?: Inches;
  y?: Inches;
  zIndex?: number;
}

/**
 * High-level wrapper class for manipulating an individual slide.
 */
export class Slide {
  private _ast: PptxSlide;
  private _presentation: Presentation;
  private _elementCounter: number = 1;

  constructor(ast: PptxSlide, presentation: Presentation) {
    this._ast = ast;
    this._presentation = presentation;
    this._elementCounter = (ast.elements?.length || 0) + 1;
  }

  /** Slide number (1-based index) */
  get slideNumber(): number {
    return this._ast.slideNumber;
  }

  /** Slide ID (e.g. `rId2`) */
  get slideId(): string {
    return this._ast.slideId;
  }

  /** Direct reference to underlying `PptxSlide` AST node */
  get ast(): PptxSlide {
    return this._ast;
  }

  /** Slide Layout ID */
  get layoutId(): string | undefined {
    return this._ast.layoutId;
  }

  /**
   * Returns an array of all visual elements on this slide.
   */
  getElements(): PptxElement[] {
    return this._ast.elements || [];
  }

  /**
   * Sets the slide background fill.
   */
  setBackground(fill: PptxFill | string): this {
    const normalized = normalizeFill(fill);
    this._ast.background = normalized ? { fill: normalized } : undefined;
    delete this._ast.rawXml;
    return this;
  }

  /**
   * Speaker notes text for this slide.
   */
  get notes(): string | undefined {
    return this._ast.notes;
  }

  /**
   * Structured text body for speaker notes.
   */
  get notesBody(): import('@hokkyss/pptx-core').PptxTextBody | undefined {
    return this._ast.notesBody;
  }

  /**
   * Sets the speaker notes for this slide with support for rich text runs (underline, bold, italic, colors, etc.).
   */
  setNotes(notes: ParagraphConfig[] | string | TextRunConfig[], options?: TextOptions): this {
    if (typeof notes === 'string') {
      this._ast.notes = notes;
      this._ast.notesBody = buildTextBody(notes, options);
    } else {
      const textBody = buildTextBody(notes, options);
      this._ast.notesBody = textBody;
      const plainLines: string[] = [];
      for (const p of textBody.paragraphs) {
        plainLines.push(p.runs.map((r) => r.text).join(''));
      }
      this._ast.notes = plainLines.join('\n');
    }
    return this;
  }

  /**
   * Resolves a placeholder shape from the slide, its parent layout, or master.
   */
  resolvePlaceholder(query: number | string): PptxElement | undefined {
    const matchCriteria = (el: PptxElement): boolean => {
      if (typeof query === 'number') {
        return el.placeholder?.idx === query || String(el.placeholder?.idx) === String(query);
      }
      const q = query.toLowerCase();
      return (
        el.name?.toLowerCase() === q
        || el.name?.toLowerCase().includes(q)
        || el.placeholder?.type?.toLowerCase() === q
        || (el.placeholder?.idx !== undefined && String(el.placeholder.idx) === query)
      );
    };

    // 1. Check slide elements first
    const slideMatch = (this._ast.elements || []).find(matchCriteria);
    if (slideMatch) return slideMatch;

    // 2. Check layout elements
    const layout = this._presentation.ast.slideLayouts?.find((l) => l.id === this._ast.layoutId);
    const layoutMatch = layout?.elements?.find(matchCriteria);
    if (layoutMatch) return layoutMatch;

    // 3. Check master elements
    if (layout?.masterId) {
      const master = this._presentation.ast.slideMasters?.find((m) => m.id === layout.masterId);
      const masterMatch = master?.elements?.find(matchCriteria);
      if (masterMatch) return masterMatch;
    }

    return undefined;
  }

  /**
   * Returns all available placeholder elements (both on this slide and un-instantiated from its layout).
   */
  getPlaceholders(): PptxElement[] {
    const slidePhs = (this._ast.elements || []).filter(
      (el) => Boolean(el.placeholder) || el.name?.startsWith('placeholder:'),
    );

    const layout = this._presentation.ast.slideLayouts?.find((l) => l.id === this._ast.layoutId);
    if (!layout?.elements) {
      return slidePhs;
    }

    const layoutPhs = layout.elements.filter(
      (el) => Boolean(el.placeholder) || el.name?.startsWith('placeholder:'),
    );

    const result = [...slidePhs];
    for (const lPh of layoutPhs) {
      const existsOnSlide = slidePhs.some(
        (s) =>
          (lPh.name && s.name === lPh.name)
          || (lPh.placeholder?.type && s.placeholder?.type === lPh.placeholder.type)
          || (lPh.placeholder?.idx !== undefined && s.placeholder?.idx === lPh.placeholder.idx),
      );
      if (!existsOnSlide) {
        result.push(lPh);
      }
    }

    return result;
  }

  /**
   * Adds a formatted text box shape or populates a layout placeholder on the slide.
   */
  addText(
    content: ParagraphConfig[] | string | TextRunConfig[],
    options: AddTextOptions = {},
  ): this {
    if (options.placeholder !== undefined) {
      const placeholderEl = this.resolvePlaceholder(options.placeholder);
      const matchCriteria = (el: PptxElement): boolean => {
        if (placeholderEl && el.id === placeholderEl.id) return true;
        if (placeholderEl?.name && el.name === placeholderEl.name) return true;
        if (typeof options.placeholder === 'number') {
          return el.placeholder?.idx === options.placeholder || String(el.placeholder?.idx) === String(options.placeholder);
        }
        const q = String(options.placeholder).toLowerCase();
        return (
          el.name?.toLowerCase() === q
          || el.placeholder?.type?.toLowerCase() === q
          || (el.placeholder?.idx !== undefined && String(el.placeholder.idx) === q)
        );
      };

      const existingSlideShape = (this._ast.elements || []).find(matchCriteria);

      if (existingSlideShape && existingSlideShape.elementType === 'shape') {
        existingSlideShape.textBody = buildTextBody(content, options);
        if (options.fill) {
          existingSlideShape.fill = normalizeFill(options.fill);
        }
        if (options.x !== undefined && options.y !== undefined && options.w !== undefined && options.h !== undefined) {
          existingSlideShape.position = {
            cx: inchesToEmu(options.w),
            cy: inchesToEmu(options.h),
            x: inchesToEmu(options.x),
            y: inchesToEmu(options.y),
          };
        }
        delete this._ast.rawXml;
        return this;
      }

      // Create new shape element bound to this placeholder
      const id = options.id || String(this._elementCounter++);
      const name = options.name || placeholderEl?.name || `placeholder:${options.placeholder}`;
      const fill = options.fill ? normalizeFill(options.fill) : placeholderEl?.fill;

      const shape: PptxShapeElement = {
        elementType: 'shape',
        fill,
        id,
        isVisible: true,
        name,
        placeholder: placeholderEl?.placeholder || (typeof options.placeholder === 'string' && ['body', 'ctrTitle', 'subTitle', 'title'].includes(options.placeholder) ? { type: options.placeholder } : undefined),
        position: {
          cx: options.w !== undefined ? inchesToEmu(options.w) : (placeholderEl?.position?.cx ?? emu(9144000)),
          cy: options.h !== undefined ? inchesToEmu(options.h) : (placeholderEl?.position?.cy ?? emu(1828800)),
          x: options.x !== undefined ? inchesToEmu(options.x) : (placeholderEl?.position?.x ?? emu(914400)),
          y: options.y !== undefined ? inchesToEmu(options.y) : (placeholderEl?.position?.y ?? emu(914400)),
        },
        rotation: options.rotation !== undefined ? degreesToEmuDegree(options.rotation) : (placeholderEl?.rotation ?? emuDegree(0)),
        shapeType: placeholderEl?.shapeType || 'rect',
        textBody: buildTextBody(content, options),
        type: 'shape',
        zIndex: options.zIndex ?? this._ast.elements.length,
      };

      this._ast.elements.push(shape);
      delete this._ast.rawXml;
      return this;
    }

    // Default standalone text box
    const id = options.id || String(this._elementCounter++);
    const name = options.name || `Text Box ${id}`;
    const fill = normalizeFill(options.fill);

    const shape: PptxShapeElement = {
      elementType: 'shape',
      fill,
      id,
      isVisible: true,
      name,
      position: {
        cx: options.w !== undefined ? inchesToEmu(options.w) : emu(5486400),
        cy: options.h !== undefined ? inchesToEmu(options.h) : emu(914400),
        x: options.x !== undefined ? inchesToEmu(options.x) : emu(914400),
        y: options.y !== undefined ? inchesToEmu(options.y) : emu(914400),
      },
      rotation: options.rotation ? degreesToEmuDegree(options.rotation) : emuDegree(0),
      shapeType: 'rect',
      textBody: buildTextBody(content, options),
      type: 'shape',
      zIndex: options.zIndex ?? 0,
    };

    this._ast.elements.push(shape);
    delete this._ast.rawXml;
    return this;
  }

  /**
   * Adds a geometric shape (rect, roundRect, ellipse, arrow, etc.) to the slide.
   */
  addShape(shapeType: string, options: AddShapeOptions): this {
    const shape = buildShapeElement(shapeType, options, this._elementCounter++);
    this._ast.elements.push(shape);
    delete this._ast.rawXml;
    return this;
  }

  /**
   * Adds an embedded image to the slide (or inside a picture placeholder).
   */
  addImage(
    imageData: ArrayBuffer | Uint8Array,
    options: AddImageOptions = {},
  ): this {
    const bytes = imageData instanceof Uint8Array ? imageData : new Uint8Array(imageData);
    const id = options.id || String(this._elementCounter++);
    const name = options.name || `Picture ${id}`;

    let placeholderEl;
    if (options.placeholder !== undefined) {
      placeholderEl = this.resolvePlaceholder(options.placeholder);
    }

    const mediaId = options.mediaId || `img_${Date.now()}_${id}`;
    const fileName = options.fileName || `${mediaId}.png`;

    this._presentation.ast.media.push({
      data: bytes,
      fileName,
      filename: fileName,
      id: mediaId,
      mimeType: fileName.endsWith('.jpg') || fileName.endsWith('.jpeg') ? 'image/jpeg' : 'image/png',
      path: `ppt/media/${fileName}`,
    });

    const picElement: PptxPictureElement = {
      elementType: 'picture',
      id,
      isVisible: true,
      name: placeholderEl?.name || name,
      picture: {
        alpha: options.alpha,
        mediaId,
      },
      position: {
        cx: options.w !== undefined ? inchesToEmu(options.w) : (placeholderEl?.position?.cx ?? emu(3657600)),
        cy: options.h !== undefined ? inchesToEmu(options.h) : (placeholderEl?.position?.cy ?? emu(2743200)),
        x: options.x !== undefined ? inchesToEmu(options.x) : (placeholderEl?.position?.x ?? emu(914400)),
        y: options.y !== undefined ? inchesToEmu(options.y) : (placeholderEl?.position?.y ?? emu(914400)),
      },
      rotation: options.rotation ? degreesToEmuDegree(options.rotation) : (placeholderEl?.rotation ?? emuDegree(0)),
      type: 'picture',
      zIndex: options.zIndex ?? 0,
    };

    this._ast.elements.push(picElement);
    delete this._ast.rawXml;
    return this;
  }

  /**
   * Adds a table to the slide using either a 2D matrix or a fluent builder.
   */
  addTable(
    dataOrBuilder: ((builder: TableBuilder) => void) | TableBuilder | TableMatrix,
    options: AddTableOptions = {},
  ): this {
    const resolvedOptions = { ...options };

    if (options.placeholder !== undefined) {
      const placeholderEl = this.resolvePlaceholder(options.placeholder);
      if (placeholderEl?.position) {
        if (resolvedOptions.x === undefined) resolvedOptions.x = emuToInches(placeholderEl.position.x);
        if (resolvedOptions.y === undefined) resolvedOptions.y = emuToInches(placeholderEl.position.y);
        if (resolvedOptions.w === undefined) resolvedOptions.w = emuToInches(placeholderEl.position.cx);
        if (resolvedOptions.h === undefined) resolvedOptions.h = emuToInches(placeholderEl.position.cy);
      }
    }

    let tableElement;

    if (dataOrBuilder instanceof TableBuilder) {
      tableElement = dataOrBuilder.build(this._elementCounter++);
    } else if (typeof dataOrBuilder === 'function') {
      const builder = new TableBuilder(resolvedOptions);
      dataOrBuilder(builder);
      tableElement = builder.build(this._elementCounter++);
    } else {
      tableElement = TableBuilder.fromMatrix(dataOrBuilder, resolvedOptions, this._elementCounter++);
    }

    this._ast.elements.push(tableElement);
    delete this._ast.rawXml;
    return this;
  }

  /**
   * Adds a connector or line between two points.
   */
  addConnector(options: AddConnectorOptions): this {
    const id = options.id || String(this._elementCounter++);
    const name = options.name || `Connector ${id}`;
    const x1 = inchesToEmu(options.from.x);
    const y1 = inchesToEmu(options.from.y);
    const x2 = inchesToEmu(options.to.x);
    const y2 = inchesToEmu(options.to.y);
    const minX = Math.min(Number(x1), Number(x2));
    const minY = Math.min(Number(y1), Number(y2));
    const cx = Math.abs(Number(x2) - Number(x1));
    const cy = Math.abs(Number(y2) - Number(y1));

    const connector: PptxConnectorElement = {
      elementType: 'connector',
      geometry: { presetGeometry: options.shapeType || 'line' },
      id,
      isVisible: true,
      line: {
        dashStyle: options.dashStyle,
        fill: options.color ? { solidColor: { type: 'srgb', value: options.color.replace(/^#/, '') }, type: 'solid' } : undefined,
        width: options.width ? inchesToEmu(options.width) : emu(19050),
      },
      name,
      position: {
        cx: emu(cx > 0 ? cx : 1),
        cy: emu(cy > 0 ? cy : 1),
        x: emu(minX),
        y: emu(minY),
      },
      rotation: emuDegree(0),
      shapeType: options.shapeType || 'line',
      type: 'connector',
      zIndex: options.zIndex ?? this._ast.elements.length,
    };

    this._ast.elements.push(connector);
    delete this._ast.rawXml;
    return this;
  }

  /**
   * Adds an interactive data chart (bar, line, pie, area) to the slide.
   */
  addChart(options: AddChartOptions): this {
    const chartElement = buildChartElement(options, this._elementCounter++);
    this._ast.elements.push(chartElement);
    delete this._ast.rawXml;
    return this;
  }

  /**
   * Adds a grouped composite shape container with multiple child elements.
   */
  addGroup(
    options: AddGroupOptions,
    builderCallback: (group: GroupBuilder) => void,
  ): this {
    const id = options.id || String(this._elementCounter++);
    const name = options.name || `Group ${id}`;
    const builder = new GroupBuilder();
    builderCallback(builder);
    const groupElement = builder.build(id, name, options);
    this._ast.elements.push(groupElement);
    delete this._ast.rawXml;
    return this;
  }

  /**
   * Deletes an element from the slide by its ID.
   */
  removeElement(elementId: string): boolean {
    const initialLen = this._ast.elements.length;
    this._ast.elements = this._ast.elements.filter((el) => el.id !== elementId);
    if (this._ast.elements.length !== initialLen) {
      delete this._ast.rawXml;
      return true;
    }
    return false;
  }
}

export interface AddConnectorOptions {
  color?: string;
  dashStyle?: string;
  from: { x: Inches; y: Inches };
  id?: string;
  name?: string;
  shapeType?: 'bentConnector2' | 'curvedConnector3' | 'line' | 'straightConnector1';
  to: { x: Inches; y: Inches };
  width?: Inches;
  zIndex?: number;
}

export interface AddGroupOptions {
  h: Inches;
  id?: string;
  name?: string;
  rotation?: Degrees;
  w: Inches;
  x: Inches;
  y: Inches;
  zIndex?: number;
}

export class GroupBuilder {
  private _elements: PptxElement[] = [];
  private _elementCounter = 1;

  addShape(shapeType: string, options: AddShapeOptions): this {
    const shape = buildShapeElement(shapeType, options, this._elementCounter++);
    this._elements.push(shape);
    return this;
  }

  addText(content: ParagraphConfig[] | string | TextRunConfig[], options: AddTextOptions): this {
    const id = options.id || String(this._elementCounter++);
    const name = options.name || `Text Box ${id}`;
    const fill = normalizeFill(options.fill);

    const shape: PptxShapeElement = {
      elementType: 'shape',
      fill,
      id,
      isVisible: true,
      name,
      position: {
        cx: options.w !== undefined ? inchesToEmu(options.w) : emu(5486400),
        cy: options.h !== undefined ? inchesToEmu(options.h) : emu(914400),
        x: options.x !== undefined ? inchesToEmu(options.x) : emu(914400),
        y: options.y !== undefined ? inchesToEmu(options.y) : emu(914400),
      },
      rotation: options.rotation ? degreesToEmuDegree(options.rotation) : emuDegree(0),
      shapeType: 'rect',
      textBody: buildTextBody(content, options),
      type: 'shape',
      zIndex: options.zIndex ?? 0,
    };
    this._elements.push(shape);
    return this;
  }

  build(id: string, name: string, options: AddGroupOptions): PptxGroupElement {
    return {
      children: this._elements,
      elementType: 'group',
      id,
      isVisible: true,
      name,
      position: {
        cx: inchesToEmu(options.w),
        cy: inchesToEmu(options.h),
        x: inchesToEmu(options.x),
        y: inchesToEmu(options.y),
      },
      rotation: options.rotation ? degreesToEmuDegree(options.rotation) : emuDegree(0),
      type: 'group',
      zIndex: options.zIndex ?? 0,
    };
  }
}
