import type {
  PptxFill,
  PptxGradientStop,
  PptxHyperlink,
  PptxLine,
  PptxShadow,
  PptxShapeElement,
} from '@hokkyss/pptx-core';
import {
  degreesToEmuDegree,
  emu,
  emuDegree,
  type Degrees,
  type Inches,
  inchesToEmu,
} from '@hokkyss/pptx-core';
import {
  buildTextBody,
  type ParagraphConfig,
  type TextOptions,
  type TextRunConfig,
} from './text-builder';

export interface GradientStopInput {
  color: string;
  opacity?: number;
  position?: number;
}

export interface GradientFillInput {
  angle?: Degrees | number;
  flip?: 'none' | 'x' | 'xy' | 'y';
  pathBounds?: {
    bottom?: number;
    left?: number;
    right?: number;
    top?: number;
  };
  rotateWithShape?: boolean;
  stops: (GradientStopInput | string)[];
  type?: 'linear' | 'path' | 'radial';
}

export type FillInput = GradientFillInput | PptxFill | string;

export interface ShapeShadowOptions {
  alignment?: string;
  blur?: Inches;
  color?: string;
  direction?: Degrees;
  distance?: Inches;
  opacity?: number;
  rotateWithShape?: boolean;
}

export interface AddShapeOptions {
  fill?: FillInput;
  h: Inches;
  hyperlink?: PptxHyperlink | string;
  id?: string;
  line?: {
    color?: string;
    dashStyle?: string;
    width?: Inches;
  };
  name?: string;
  rotation?: Degrees;
  shadow?: ShapeShadowOptions;
  text?: ParagraphConfig[] | string | TextRunConfig[];
  textOptions?: TextOptions;
  w: Inches;
  x: Inches;
  y: Inches;
  zIndex?: number;
}

/**
 * Normalizes fill string, shorthand gradient object, or PptxFill into a typed PptxFill.
 */
export function normalizeFill(fillInput?: FillInput): PptxFill | undefined {
  if (!fillInput) return undefined;
  if (typeof fillInput === 'string') {
    return {
      solidColor: {
        type: 'srgb',
        value: fillInput.replace(/^#/, ''),
      },
      type: 'solid',
    };
  }

  // Handle shorthand GradientFillInput
  if ('stops' in fillInput && Array.isArray(fillInput.stops)) {
    const gradInput = fillInput;
    const rawStops = gradInput.stops;
    const total = rawStops.length;

    const stops: PptxGradientStop[] = rawStops.map((stop, idx) => {
      if (typeof stop === 'string') {
        const pos = total > 1 ? idx / (total - 1) : 0;
        return {
          color: stop,
          position: pos,
        };
      }
      return {
        color: stop.color,
        opacity: stop.opacity,
        position: stop.position !== undefined ? stop.position : (total > 1 ? idx / (total - 1) : 0),
      };
    });

    return {
      gradient: {
        angle: gradInput.angle,
        flip: gradInput.flip,
        pathBounds: gradInput.pathBounds,
        rotateWithShape: gradInput.rotateWithShape,
        stops,
        type: gradInput.type ?? 'linear',
      },
      type: 'gradient',
    };
  }

  return fillInput as PptxFill;
}

const PRESET_GEOMETRY_MAP: Record<string, string> = {
  box: 'rect',
  circle: 'ellipse',
  cylinder: 'can',
  oval: 'ellipse',
  square: 'rect',
  star: 'star5',
  wedgeRoundRect: 'wedgeRoundRectCallout',
};

/**
 * Builds a PptxShapeElement AST node for a geometric shape.
 */
export function buildShapeElement(
  shapeType: string,
  options: AddShapeOptions,
  counter: number = 1,
): PptxShapeElement {
  const resolvedShapeType = PRESET_GEOMETRY_MAP[shapeType] || shapeType;
  const id = options.id || String(counter);
  const name = options.name || `${shapeType} ${id}`;

  const fill = normalizeFill(options.fill);

  let line: PptxLine | undefined;
  if (options.line) {
    line = {
      dashStyle: options.line.dashStyle,
      fill: options.line.color ? { solidColor: { type: 'srgb', value: options.line.color.replace(/^#/, '') }, type: 'solid' } : undefined,
      width: options.line.width ? inchesToEmu(options.line.width) : emu(12700),
    };
  }

  let shadow: PptxShadow | undefined;
  if (options.shadow) {
    shadow = {
      alignment: options.shadow.alignment,
      blurRadius: options.shadow.blur ? inchesToEmu(options.shadow.blur) : emu(50800),
      color: options.shadow.color ? options.shadow.color.replace(/^#/, '') : '000000',
      direction: options.shadow.direction ? degreesToEmuDegree(options.shadow.direction) : emuDegree(5400000),
      distance: options.shadow.distance ? inchesToEmu(options.shadow.distance) : emu(38100),
      opacity: options.shadow.opacity ?? 0.25,
      rotateWithShape: options.shadow.rotateWithShape,
    };
  }

  let textBody = undefined;
  if (options.text) {
    textBody = buildTextBody(options.text, options.textOptions);
  }

  return {
    elementType: 'shape',
    fill,
    hyperlink: options.hyperlink,
    id,
    isVisible: true,
    line,
    name,
    position: {
      cx: inchesToEmu(options.w),
      cy: inchesToEmu(options.h),
      x: inchesToEmu(options.x),
      y: inchesToEmu(options.y),
    },
    rotation: options.rotation ? degreesToEmuDegree(options.rotation) : emuDegree(0),
    shadow,
    shapeType: resolvedShapeType,
    textBody,
    type: 'shape',
    zIndex: options.zIndex ?? 0,
  };
}
