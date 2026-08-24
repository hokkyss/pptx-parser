import { describe, expect, it } from 'vitest';
import {
  degreesToGradientAngle,
  gradientAngleToDegrees,
  normalizeStopPosition,
} from '../../lib/color';

describe('Color & Gradient Utilities (@hokkyss/pptx-core)', () => {
  describe('degreesToGradientAngle & gradientAngleToDegrees', () => {
    it('converts degrees (0..360) to DrawingML gradient angle (60,000ths of a deg) and back', () => {
      expect(degreesToGradientAngle(0)).toBe(0);
      expect(degreesToGradientAngle(90)).toBe(5400000);
      expect(degreesToGradientAngle(180)).toBe(10800000);
      expect(degreesToGradientAngle(270)).toBe(16200000);
      expect(degreesToGradientAngle(360)).toBe(0);

      expect(gradientAngleToDegrees(0)).toBe(0);
      expect(gradientAngleToDegrees(5400000)).toBe(90);
      expect(gradientAngleToDegrees(10800000)).toBe(180);
      expect(gradientAngleToDegrees(16200000)).toBe(270);
    });

    it('handles negative or oversized degrees', () => {
      expect(degreesToGradientAngle(-90)).toBe(16200000);
      expect(degreesToGradientAngle(450)).toBe(5400000);
    });
  });

  describe('normalizeStopPosition', () => {
    it('normalizes 0..1 float position to ThousandthsPercent (0..100000)', () => {
      expect(normalizeStopPosition(0)).toBe(0);
      expect(normalizeStopPosition(0.5)).toBe(50000);
      expect(normalizeStopPosition(1)).toBe(100000);
    });

    it('normalizes integer ThousandthsPercent position (0..100000)', () => {
      expect(normalizeStopPosition(25000)).toBe(25000);
      expect(normalizeStopPosition(100000)).toBe(100000);
    });

    it('handles undefined or NaN with defaultPos fallback', () => {
      expect(normalizeStopPosition(undefined, 0.5)).toBe(50000);
      expect(normalizeStopPosition(NaN, 0)).toBe(0);
      expect(normalizeStopPosition(undefined, 75000)).toBe(75000);
    });

    it('clamps out of bounds values', () => {
      expect(normalizeStopPosition(-0.5)).toBe(0);
      expect(normalizeStopPosition(150000)).toBe(100000);
    });
  });
});
