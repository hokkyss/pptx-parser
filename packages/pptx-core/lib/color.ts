import type { EmuDegree, ThousandthsPercent } from './units';

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
  /** Color at stop. OpenXML: `<a:gs><a:srgbClr>` */
  color: PptxColor;
  /** Stop position in Thousandths of a Percent. OpenXML: `<a:gs @_pos>` */
  position: ThousandthsPercent;
}

/** Gradient fill properties */
export interface PptxGradientFill {
  /** Angle in EMU degrees. OpenXML: `<a:lin @_ang>` */
  angle?: EmuDegree;
  /** Gradient stops. OpenXML: `<a:gsLst><a:gs>` */
  stops: PptxGradientStop[];
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
