import type {
  PptxConnectionPosition,
  PptxConnectorElement,
  PptxElement,
  PptxFill,
  PptxGroupElement,
  PptxHyperlink,
  PptxLineEnd,
  PptxLineEndLength,
  PptxLineEndType,
  PptxLineEndWidth,
  PptxPictureElement,
  PptxShapeAttachment,
  PptxShapeElement,
  PptxSlide,
  PptxTransition,
  PptxTransitionDirection,
  PptxTransitionSpeed,
  PptxTransitionType,
} from '@hokkyss/pptx-core';

export type ConnectionPosition = PptxConnectionPosition;
export type ShapeAttachment = PptxShapeAttachment;
export type ConnectorEndpoint = { x: Inches; y: Inches } | ShapeAttachment;

export type {
  PptxLineEnd,
  PptxLineEndLength,
  PptxLineEndType,
  PptxLineEndWidth,
  PptxTransition,
  PptxTransitionDirection,
  PptxTransitionSpeed,
  PptxTransitionType,
};
import {
  degreesToEmuDegree,
  emu,
  emuDegree,
  emuToInches,
  type Degrees,
  type Emu,
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
  type FillInput,
  type GradientFillInput,
  type GradientStopInput,
  normalizeFill,
} from './builders/shape-builder';

export type { FillInput, GradientFillInput, GradientStopInput };
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
  hyperlink?: PptxHyperlink | string;
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
  setBackground(fill: FillInput): this {
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
        hyperlink: options.hyperlink,
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
      hyperlink: options.hyperlink,
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
      hyperlink: options.hyperlink,
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
   * Adds a connector or line between two points or attached to shapes.
   * @example
   * ```ts
   * // Attached to shapes
   * slide.addConnector({
   *   from: { shapeId: 'card-1', position: 'right' },
   *   to: { shapeId: 'card-2', position: 'left' },
   *   endArrow: 'triangle',
   *   color: '0284C7',
   * });
   *
   * // Arbitrary coordinates
   * slide.addConnector({
   *   from: { x: inches(1), y: inches(1) },
   *   to: { x: inches(4), y: inches(1) },
   *   endArrow: 'triangle',
   *   startArrow: 'oval',
   * });
   * ```
   */
  addConnector(options: AddConnectorOptions): this {
    const connector = buildConnectorElement(options, this._elementCounter++, this._ast.elements);
    if (options.zIndex === undefined) {
      connector.zIndex = this._ast.elements.length;
    }

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

  /**
   * Configures or updates the slide transition effect.
   * @param transition Transition type string (e.g. 'fade', 'wipe', 'push', 'split') or full PptxTransition configuration.
   * @param options Optional transition settings (direction, speed, durationMs, advanceAfterMs, etc.)
   * @example
   * ```ts
   * slide.setTransition('fade', { durationMs: 500 });
   * slide.setTransition('wipe', { direction: 'right', speed: 'fast' });
   * slide.setTransition('push', { direction: 'up', advanceAfterMs: 3000 });
   * ```
   */
  setTransition(
    transition: PptxTransition | PptxTransitionType,
    options?: Omit<PptxTransition, 'type'>,
  ): this {
    if (typeof transition === 'string') {
      this._ast.transition = {
        type: transition,
        ...options,
      };
    } else {
      this._ast.transition = {
        ...transition,
        ...options,
      };
    }
    delete this._ast.rawXml;
    return this;
  }

  /**
   * Retrieves the current slide transition configuration if defined.
   */
  getTransition(): PptxTransition | undefined {
    return this._ast.transition;
  }

  /**
   * Removes any transition effect from the slide.
   */
  clearTransition(): this {
    delete this._ast.transition;
    delete this._ast.rawXml;
    return this;
  }
}

/**
 * Normalizes string or object line end configuration.
 */
function normalizeLineEnd(input?: PptxLineEnd | PptxLineEndType): PptxLineEnd | undefined {
  if (!input) return undefined;
  if (typeof input === 'string') {
    return { type: input };
  }
  return input;
}

/**
 * Resolves a connector endpoint (coordinate or shape attachment) to EMU coordinates and optional OpenXML attachment point.
 */
function resolveEndpoint(
  endpoint: ConnectorEndpoint,
  slideElements?: PptxElement[],
): { connection?: PptxShapeAttachment; emuX: Emu; emuY: Emu } {
  if ('shapeId' in endpoint) {
    const shape = slideElements?.find((el) => el.id === endpoint.shapeId);
    if (!shape) {
      throw new Error(`Shape with id "${endpoint.shapeId}" was not found on this slide. Ensure the shape is added before attaching a connector to it.`);
    }
    const x = Number(shape.position?.x ?? 0);
    const y = Number(shape.position?.y ?? 0);
    const w = Number(shape.position?.cx ?? 0);
    const h = Number(shape.position?.cy ?? 0);

    let emuX: Emu;
    let emuY: Emu;

    switch (endpoint.position) {
      case 'bottom':
        emuX = emu(Math.round(x + w / 2));
        emuY = emu(Math.round(y + h));
        break;
      case 'left':
        emuX = emu(Math.round(x));
        emuY = emu(Math.round(y + h / 2));
        break;
      case 'right':
        emuX = emu(Math.round(x + w));
        emuY = emu(Math.round(y + h / 2));
        break;
      case 'top':
        emuX = emu(Math.round(x + w / 2));
        emuY = emu(Math.round(y));
        break;
      default:
        emuX = emu(Math.round(x));
        emuY = emu(Math.round(y));
    }

    return {
      connection: { position: endpoint.position, shapeId: endpoint.shapeId },
      emuX,
      emuY,
    };
  }

  return {
    emuX: inchesToEmu(endpoint.x),
    emuY: inchesToEmu(endpoint.y),
  };
}

/**
 * Builds a connector element AST node.
 */
export function buildConnectorElement(
  options: AddConnectorOptions,
  fallbackId: number | string,
  slideElements?: PptxElement[],
): PptxConnectorElement {
  const id = options.id || String(fallbackId);
  const name = options.name || `Connector ${id}`;

  const fromResolved = resolveEndpoint(options.from, slideElements);
  const toResolved = resolveEndpoint(options.to, slideElements);

  const x1 = fromResolved.emuX;
  const y1 = fromResolved.emuY;
  const x2 = toResolved.emuX;
  const y2 = toResolved.emuY;
  const minX = Math.min(Number(x1), Number(x2));
  const minY = Math.min(Number(y1), Number(y2));
  const cx = Math.abs(Number(x2) - Number(x1));
  const cy = Math.abs(Number(y2) - Number(y1));

  const headEnd = normalizeLineEnd(options.headEnd ?? options.endArrow);
  const tailEnd = normalizeLineEnd(options.tailEnd ?? options.startArrow);

  return {
    elementType: 'connector',
    ...(toResolved.connection && { endConnection: toResolved.connection }),
    geometry: { presetGeometry: options.shapeType || 'line' },
    id,
    isVisible: true,
    line: {
      dashStyle: options.dashStyle,
      fill: options.color
        ? { solidColor: { type: 'srgb', value: options.color.replace(/^#/, '') }, type: 'solid' }
        : undefined,
      width: options.width ? inchesToEmu(options.width) : emu(19050),
      ...(headEnd && { headEnd }),
      ...(tailEnd && { tailEnd }),
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
    ...(fromResolved.connection && { startConnection: fromResolved.connection }),
    type: 'connector',
    zIndex: options.zIndex ?? 0,
  };
}

export interface AddConnectorOptions {
  /** Line stroke color (hex string, e.g. '#0284C7' or '0284C7') */
  color?: string;
  /** Stroke dash style. OpenXML: `<a:prstDash @_val>` */
  dashStyle?: string;
  /** End arrowhead / line marker (alias for headEnd). OpenXML: `<a:headEnd>` */
  endArrow?: PptxLineEnd | PptxLineEndType;
  /** Starting coordinate point `{ x: Inches, y: Inches }` or shape attachment `{ shapeId: string, position: 'top' | 'bottom' | 'left' | 'right' }` */
  from: ConnectorEndpoint;
  /** Head / end arrowhead marker. OpenXML: `<a:headEnd>` */
  headEnd?: PptxLineEnd | PptxLineEndType;
  /** Optional custom ID */
  id?: string;
  /** Optional element name */
  name?: string;
  /** Shape preset geometry for connector (defaults to 'line') */
  shapeType?: 'bentConnector2' | 'curvedConnector3' | 'line' | 'straightConnector1';
  /** Start arrowhead / line marker (alias for tailEnd). OpenXML: `<a:tailEnd>` */
  startArrow?: PptxLineEnd | PptxLineEndType;
  /** Tail / start arrowhead marker. OpenXML: `<a:tailEnd>` */
  tailEnd?: PptxLineEnd | PptxLineEndType;
  /** Ending coordinate point `{ x: Inches, y: Inches }` or shape attachment `{ shapeId: string, position: 'top' | 'bottom' | 'left' | 'right' }` */
  to: ConnectorEndpoint;
  /** Line thickness width in inches */
  width?: Inches;
  /** Visual stacking order */
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

  addConnector(options: AddConnectorOptions): this {
    const connector = buildConnectorElement(options, this._elementCounter++, this._elements);
    this._elements.push(connector);
    return this;
  }

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
