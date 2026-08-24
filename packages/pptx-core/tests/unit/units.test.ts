import { describe, expect, it } from 'vitest';
import * as Core from '../../lib/index';
import {
  centimeters,
  cmToEmu,
  decimalToThousandthsPercent,
  degrees,
  degreesToEmuDegree,
  degreesToRotation,
  emu,
  emuDegree,
  emuToCm,
  emuToInches,
  emuToPoints,
  emuToPx,
  hundredthsPoint,
  hundredthsPointToPoint,
  hundredthsPointToPoints,
  inches,
  inchesToEmu,
  percent,
  percentToThousandthsPercent,
  pixels,
  points,
  pointsToEmu,
  pointsToHundredthsPoint,
  pointToHundredthsPoint,
  pxToEmu,
  rotationToDegrees,
  thousandthsPercent,
  thousandthsPercentToDecimal,
  thousandthsPercentToPercent,
} from '../../lib/units';

describe('Units & Branded Types (@hokkyss/pptx-core)', () => {
  it('exports all types and functions from root index', () => {
    expect(Core.emu).toBeDefined();
    expect(Core.inches).toBeDefined();
  });

  describe('Constructors & Validations', () => {
    it('creates branded numbers for valid finite inputs', () => {
      expect(emu(914400)).toBe(914400);
      expect(inches(1.5)).toBe(1.5);
      expect(points(12)).toBe(12);
      expect(centimeters(2.54)).toBe(2.54);
      expect(pixels(96)).toBe(96);
      expect(hundredthsPoint(1200)).toBe(1200);
      expect(emuDegree(5400000)).toBe(5400000);
      expect(degrees(90)).toBe(90);
      expect(thousandthsPercent(50000)).toBe(50000);
      expect(percent(50)).toBe(50);
    });

    it('allows 0 and negative values', () => {
      expect(emu(0)).toBe(0);
      expect(emu(-100)).toBe(-100);
      expect(degrees(-45)).toBe(-45);
      expect(percent(0)).toBe(0);
    });

    it('throws TypeError for non-finite values (NaN, Infinity, non-number)', () => {
      expect(() => emu(NaN)).toThrow(TypeError);
      expect(() => emu(Infinity)).toThrow(TypeError);
      expect(() => emu(-Infinity)).toThrow(TypeError);
      expect(() => inches('1' as unknown as number)).toThrow(TypeError);
      expect(() => points(undefined as unknown as number)).toThrow(TypeError);
    });
  });

  describe('Conversions', () => {
    it('converts EMU to Inches and back', () => {
      expect(emuToInches(emu(914400))).toBe(1);
      expect(inchesToEmu(inches(1))).toBe(914400);
      expect(emuToInches(emu(0))).toBe(0);
      expect(inchesToEmu(inches(0))).toBe(0);
    });

    it('converts EMU to Points and back', () => {
      expect(emuToPoints(emu(12700))).toBe(1);
      expect(pointsToEmu(points(1))).toBe(12700);
      expect(pointsToEmu(points(72))).toBe(914400);
    });

    it('converts EMU to Centimeters and back', () => {
      expect(emuToCm(emu(360000))).toBe(1);
      expect(cmToEmu(centimeters(1))).toBe(360000);
    });

    it('converts EMU to Pixels and back', () => {
      expect(emuToPx(emu(914400), 96)).toBe(96);
      expect(pxToEmu(pixels(96), 96)).toBe(914400);
    });

    it('converts EmuDegree to Degrees and back', () => {
      expect(rotationToDegrees(emuDegree(5400000))).toBe(90);
      expect(degreesToRotation(degrees(90))).toBe(5400000);
      expect(degreesToEmuDegree(degrees(90))).toBe(5400000);
      expect(rotationToDegrees(emuDegree(0))).toBe(0);
      expect(degreesToRotation(degrees(0))).toBe(0);
    });

    it('converts Points to HundredthsPoint and back', () => {
      expect(pointsToHundredthsPoint(points(12))).toBe(1200);
      expect(pointToHundredthsPoint(points(12))).toBe(1200);
      expect(hundredthsPointToPoints(hundredthsPoint(1200))).toBe(12);
      expect(hundredthsPointToPoint(hundredthsPoint(1200))).toBe(12);
    });

    it('converts Percent to ThousandthsPercent and back', () => {
      expect(percentToThousandthsPercent(percent(50))).toBe(50000);
      expect(thousandthsPercentToPercent(thousandthsPercent(50000))).toBe(50);
      expect(percentToThousandthsPercent(percent(100))).toBe(100000);
      expect(thousandthsPercentToPercent(thousandthsPercent(100000))).toBe(100);
      expect(thousandthsPercentToDecimal(thousandthsPercent(100000))).toBe(1.0);
      expect(decimalToThousandthsPercent(1.0)).toBe(100000);
    });
  });
});
