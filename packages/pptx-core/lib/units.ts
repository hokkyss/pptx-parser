/**
 * Branded type for English Metric Units (EMU).
 * 1 inch = 914400 EMU, 1 pt = 12700 EMU.
 */
export type Emu = { readonly __brand: 'Emu' } & number;

/**
 * Branded type for Inches.
 */
export type Inches = { readonly __brand: 'Inches' } & number;

/**
 * Branded type for Standard Points (pt).
 * 1 pt = 1/72 inch = 12700 EMU.
 */
export type Points = { readonly __brand: 'Points' } & number;

/**
 * Branded type for Centimeters (cm).
 * 1 cm = 360000 EMU.
 */
export type Centimeters = { readonly __brand: 'Centimeters' } & number;

/**
 * Branded type for Screen Pixels (px).
 */
export type Pixels = { readonly __brand: 'Pixels' } & number;

/**
 * Branded type for Hundredths of a Point (cpt).
 * 1 pt = 100 Hundredths of a Point.
 */
export type HundredthsPoint = { readonly __brand: 'HundredthsPoint' } & number;

/**
 * Branded type for EMU Degrees.
 * 1 degree = 60000 EMU.
 */
export type EmuDegree = { readonly __brand: 'EmuDegree' } & number;

/**
 * Branded type for Standard Degrees (°).
 */
export type Degrees = { readonly __brand: 'Degrees' } & number;

/**
 * Branded type for Thousandths of a Percent.
 * 100% = 100000 Thousandths of a Percent.
 */
export type ThousandthsPercent = { readonly __brand: 'ThousandthsPercent' } & number;

/**
 * Branded type for Percent (%).
 * e.g. 50% = 50.
 */
export type Percent = { readonly __brand: 'Percent' } & number;

/**
 * Identity constructor for English Metric Units (EMU).
 * @param value
 */
export function emu(value: number): Emu {
  return value as Emu;
}

/**
 * Identity constructor for Inches.
 * @param value
 */
export function inches(value: number): Inches {
  return value as Inches;
}

/**
 * Identity constructor for Points (pt).
 * @param value
 */
export function points(value: number): Points {
  return value as Points;
}

/**
 * Identity constructor for Centimeters (cm).
 * @param value
 */
export function centimeters(value: number): Centimeters {
  return value as Centimeters;
}

/**
 * Identity constructor for Pixels (px).
 * @param value
 */
export function pixels(value: number): Pixels {
  return value as Pixels;
}

/**
 * Identity constructor for Hundredths of a Point (cpt).
 * @param value
 */
export function hundredthsPoint(value: number): HundredthsPoint {
  return value as HundredthsPoint;
}

/**
 * Identity constructor for EMU Degrees.
 * @param value
 */
export function emuDegree(value: number): EmuDegree {
  return value as EmuDegree;
}

/**
 * Identity constructor for Standard Degrees (°).
 * @param value
 */
export function degrees(value: number): Degrees {
  return value as Degrees;
}

/**
 * Identity constructor for Thousandths of a Percent.
 * @param value
 */
export function thousandthsPercent(value: number): ThousandthsPercent {
  return value as ThousandthsPercent;
}

/**
 * Identity constructor for Percent (%).
 * @param value
 */
export function percent(value: number): Percent {
  return value as Percent;
}

// -------------------------------------------------------------
// Unit Conversions
// -------------------------------------------------------------

/**
 * Converts Centimeters (cm) to English Metric Units (EMU).
 * Formula: 1 cm = 360,000 EMU.
 * @param cm
 */
export function cmToEmu(cm: Centimeters | number): Emu {
  return emu(cm * 360000);
}

/**
 * Converts a unit Decimal (0.0 to 1.0) to OpenXML Thousandths of a Percent.
 * Formula: 1.0 (100%) = 100,000 Thousandths of a Percent.
 * @param decimal
 */
export function decimalToThousandthsPercent(decimal: number): ThousandthsPercent {
  return thousandthsPercent(decimal * 100000);
}

/**
 * Converts standard floating-point degrees (0° to 360°) to OpenXML EMU Degrees.
 * Formula: 1 degree = 60,000 EMU Degrees.
 * @param deg
 */
export function degreesToEmuDegree(deg: Degrees | number): EmuDegree {
  return emuDegree(deg * 60000);
}

/** Alias for degreesToEmuDegree */
export const degreesToRotation = degreesToEmuDegree;

/**
 * Converts OpenXML English Metric Units (EMU) to Centimeters (cm).
 * Formula: 1 cm = 360,000 EMU.
 * @param emuVal
 */
export function emuToCm(emuVal: Emu | number): Centimeters {
  return centimeters(emuVal / 360000);
}

/**
 * Converts OpenXML English Metric Units (EMU) to Inches.
 * Formula: 1 inch = 914,400 EMU.
 * @param emuVal
 */
export function emuToInches(emuVal: Emu | number): Inches {
  return inches(emuVal / 914400);
}

/**
 * Converts OpenXML English Metric Units (EMU) to Points (pt).
 * Formula: 1 point = 12,700 EMU.
 * @param emuVal
 */
export function emuToPoints(emuVal: Emu | number): Points {
  return points(emuVal / 12700);
}

/**
 * Converts OpenXML English Metric Units (EMU) to Screen Pixels (px).
 * Formula: (emu / 914,400) * DPI.
 * @param emuVal
 * @param dpi
 */
export function emuToPx(emuVal: Emu | number, dpi: number = 96): Pixels {
  return pixels((emuVal / 914400) * dpi);
}

/**
 * Converts OpenXML Hundredths of a Point (cpt) to standard Points (pt).
 * Formula: 1 pt = 100 Hundredths of a Point.
 * @param cpt
 */
export function hundredthsPointToPoints(cpt: HundredthsPoint | number): Points {
  return points(cpt / 100);
}

/** Alias for `hundredthsPointToPoints` */
export const hundredthsPointToPoint = hundredthsPointToPoints;

/**
 * Converts Inches to OpenXML English Metric Units (EMU).
 * Formula: 1 inch = 914,400 EMU.
 * @param inchVal
 */
export function inchesToEmu(inchVal: Inches | number): Emu {
  return emu(inchVal * 914400);
}

/**
 * Converts Points (pt) to OpenXML English Metric Units (EMU).
 * Formula: 1 pt = 12,700 EMU.
 * @param pt
 */
export function pointsToEmu(pt: number | Points): Emu {
  return emu(pt * 12700);
}

/**
 * Converts standard Points (pt) to OpenXML Hundredths of a Point (cpt).
 * Formula: 1 pt = 100 Hundredths of a Point.
 * @param pt
 */
export function pointsToHundredthsPoint(pt: number | Points): HundredthsPoint {
  return hundredthsPoint(pt * 100);
}

/** Alias for `pointsToHundredthsPoint` */
export const pointToHundredthsPoint = pointsToHundredthsPoint;

/**
 * Converts Screen Pixels (px) to OpenXML English Metric Units (EMU).
 * Formula: ((px / DPI) * 914,400).
 * @param px
 * @param dpi
 */
export function pxToEmu(px: number | Pixels, dpi: number = 96): Emu {
  return emu((px / dpi) * 914400);
}

/**
 * Converts OpenXML EMU Degrees to standard floating-point degrees (0° to 360°).
 * Formula: 1 degree = 60,000 EMU Degrees.
 * @param rot
 */
export function rotationToDegrees(rot: EmuDegree | number): Degrees {
  return degrees(rot / 60000);
}

/**
 * Converts OpenXML Thousandths of a Percent to a unit decimal (0.0 to 1.0).
 * Formula: 100,000 Thousandths of a Percent = 1.0 (100%).
 * @param tpct
 */
export function thousandthsPercentToDecimal(tpct: number | ThousandthsPercent): number {
  return tpct / 100000;
}

/**
 * Converts OpenXML Thousandths of a Percent to standard Percent (0 to 100).
 * Formula: 1,000 Thousandths of a Percent = 1%.
 * @param tpct
 */
export function thousandthsPercentToPercent(tpct: number | ThousandthsPercent): Percent {
  return percent(tpct / 1000);
}

/**
 * Converts standard Percent (0 to 100) to OpenXML Thousandths of a Percent.
 * Formula: 1% = 1,000 Thousandths of a Percent.
 * @param pct
 */
export function percentToThousandthsPercent(pct: number | Percent): ThousandthsPercent {
  return thousandthsPercent(pct * 1000);
}
