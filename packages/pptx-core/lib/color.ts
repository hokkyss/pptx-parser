import type { Degrees, EmuDegree, ThousandthsPercent } from './units';
import { thousandthsPercent } from './units';

/** Represents a color value */
export interface PptxColor {
  /** Alpha transparency in Thousandths of a Percent. OpenXML: `<a:alpha @_val>` */
  alpha?: ThousandthsPercent;
  /** Color type. OpenXML: `<a:srgbClr>`, `<a:schemeClr>`, `<a:sysClr>` */
  type: 'scheme' | 'srgb' | 'system';
  /** Color value (e.g., hex string or scheme color name). OpenXML: `<a:*Clr @_val>` */
  value: string;
}

/** Gradient stop */
export interface PptxGradientStop {
  /** Color at stop. OpenXML: `<a:gs><a:srgbClr>` or `<a:gs><a:schemeClr>` */
  color: PptxColor | string;
  /** Stop position in Thousandths of a Percent (0..100000) or normalized float (0..1). OpenXML: `<a:gs @_pos>` */
  position: number | ThousandthsPercent;
  /** Optional alpha/opacity (0..1 or ThousandthsPercent 0..100000) */
  opacity?: number;
}

/** Gradient fill type classification */
export type PptxGradientType = 'linear' | 'path' | 'radial';

/** Gradient fill properties */
export interface PptxGradientFill {
  /** Angle in degrees (0..360) or EMU degrees (0..21600000). OpenXML: `<a:lin @_ang>` */
  angle?: Degrees | EmuDegree | number;
  /** Flip mode. OpenXML: `<a:gradFill @_flip>` */
  flip?: 'none' | 'x' | 'xy' | 'y';
  /** Radial / path fill rectangle bounds (percentages 0..1 or 0..100000). OpenXML: `<a:fillToRect>` */
  pathBounds?: {
    bottom?: number;
    left?: number;
    right?: number;
    top?: number;
  };
  /** Whether the gradient rotates with the shape. OpenXML: `<a:gradFill @_rotWithShape>` */
  rotateWithShape?: boolean;
  /** Gradient stops. OpenXML: `<a:gsLst><a:gs>` */
  stops: PptxGradientStop[];
  /** Gradient fill type: linear (default), radial, or path. */
  type?: PptxGradientType;
}

/** Represents fill properties */
export interface PptxFill {
  /** Gradient fill. OpenXML: `<a:gradFill>` */
  gradient?: PptxGradientFill;
  /** Solid color fill. OpenXML: `<a:solidFill>` */
  solidColor?: PptxColor;
  /** Fill type. OpenXML: `<a:solidFill>`, `<a:gradFill>`, `<a:pattFill>`, `<a:blipFill>`, `<a:noFill>` */
  type: 'gradient' | 'none' | 'pattern' | 'picture' | 'solid';
}

/**
 * Converts angle in degrees (0..360) to OpenXML DrawingML linear gradient angle (60,000ths of a degree).
 * @param deg Degrees (0..360).
 * @returns DrawingML angle integer (0..21600000).
 */
export function degreesToGradientAngle(deg: number): number {
  const normalized = ((deg % 360) + 360) % 360;
  return Math.round(normalized * 60000);
}

/**
 * Converts OpenXML DrawingML linear gradient angle (60,000ths of a degree) to degrees (0..360).
 * @param ang DrawingML angle integer.
 * @returns Degrees float (0..360).
 */
export function gradientAngleToDegrees(ang: number): number {
  return Math.round((ang / 60000) * 100) / 100;
}

/**
 * Normalizes a gradient stop position from either a 0..1 float or 0..100000 integer to ThousandthsPercent.
 * @param pos Stop position (0..1 or 0..100000).
 * @param defaultPos Optional fallback position.
 * @returns Normalized ThousandthsPercent.
 */
export function normalizeStopPosition(pos?: number, defaultPos: number = 0): ThousandthsPercent {
  if (pos === undefined || Number.isNaN(pos)) {
    return thousandthsPercent(Math.round(defaultPos <= 1 ? defaultPos * 100000 : defaultPos));
  }
  if (pos <= 1) {
    return thousandthsPercent(Math.round(Math.max(0, Math.min(1, pos)) * 100000));
  }
  return thousandthsPercent(Math.round(Math.max(0, Math.min(100000, pos))));
}
