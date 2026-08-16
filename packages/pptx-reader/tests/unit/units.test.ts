import { describe, it, expect } from 'vitest';
import {
  centimeters,
  cmToEmu,
  decimalToThousandthsPercent,
  degrees,
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
} from '@hokkyss/pptx-core';

describe('Unit Identity Functions', () => {
  it('should declare branded units with identity constructor functions', () => {
    expect(emu(100)).toBe(100);
    expect(inches(1.5)).toBe(1.5);
    expect(points(12)).toBe(12);
    expect(centimeters(2.54)).toBe(2.54);
    expect(pixels(96)).toBe(96);
    expect(hundredthsPoint(1800)).toBe(1800);
    expect(emuDegree(5400000)).toBe(5400000);
    expect(degrees(90)).toBe(90);
    expect(thousandthsPercent(50000)).toBe(50000);
    expect(percent(50)).toBe(50);
  });
});

describe('Unit Conversion Functions (Strongly-Typed Units)', () => {
  it('should convert Emu to Inches and back', () => {
    const inchVal = emuToInches(emu(914400));
    expect(inchVal).toBe(1);

    const emuVal = inchesToEmu(inches(1));
    expect(emuVal).toBe(914400);
  });

  it('should convert Emu to Centimeters and back', () => {
    const cmVal = emuToCm(emu(360000));
    expect(cmVal).toBe(1);

    const emuVal = cmToEmu(centimeters(1));
    expect(emuVal).toBe(360000);
  });

  it('should convert Emu to Points and back', () => {
    const ptVal = emuToPoints(emu(12700));
    expect(ptVal).toBe(1);

    const emuVal = pointsToEmu(points(1));
    expect(emuVal).toBe(12700);
  });

  it('should convert Emu to Pixels and back', () => {
    const pxVal = emuToPx(emu(914400), 96);
    expect(pxVal).toBe(96);

    const emuVal = pxToEmu(pixels(96), 96);
    expect(emuVal).toBe(914400);
  });

  it('should convert HundredthsPoint to Points and back', () => {
    const ptVal = hundredthsPointToPoints(hundredthsPoint(1800));
    expect(ptVal).toBe(18);

    const cptVal = pointsToHundredthsPoint(points(18));
    expect(cptVal).toBe(1800);

    // Aliases
    expect(hundredthsPointToPoint(hundredthsPoint(1800))).toBe(18);
    expect(pointToHundredthsPoint(points(18))).toBe(1800);
  });

  it('should convert ThousandthsPercent to decimal, percent and back', () => {
    expect(thousandthsPercentToDecimal(thousandthsPercent(100000))).toBe(1);
    expect(thousandthsPercentToDecimal(thousandthsPercent(50000))).toBe(0.5);
    expect(decimalToThousandthsPercent(0.5)).toBe(50000);

    const pctVal = thousandthsPercentToPercent(thousandthsPercent(50000));
    expect(pctVal).toBe(50);

    const tpctVal = percentToThousandthsPercent(percent(50));
    expect(tpctVal).toBe(50000);
  });

  it('should convert EmuDegree to Degrees and back', () => {
    expect(rotationToDegrees(emuDegree(5400000))).toBe(90);
    expect(rotationToDegrees(emuDegree(18000000))).toBe(300);
    expect(rotationToDegrees(emuDegree(0))).toBe(0);

    expect(degreesToRotation(degrees(90))).toBe(5400000);
    expect(degreesToRotation(degrees(300))).toBe(18000000);
    expect(degreesToRotation(degrees(0))).toBe(0);
  });
});
