import { describe, it, expect } from 'vitest';
import {
  emuToPixels,
  pixelsToEmu,
  emuToPoints,
  pointsToEmu,
  pointsToPixels,
  pixelsToPoints,
  centipointsToPixels,
  ooxmlPercentToDecimal,
  ooxmlAngleToDegrees,
  EMU_PER_INCH,
  EMU_PER_POINT,
  DEFAULT_DPI,
} from '../src/utils/units';

describe('Unit Conversions', () => {
  describe('emuToPixels', () => {
    it('converts 1 inch of EMUs to 96 pixels at default DPI', () => {
      expect(emuToPixels(914400)).toBe(96);
    });

    it('converts 10 inches of EMUs to 960 pixels', () => {
      expect(emuToPixels(9144000)).toBe(960);
    });

    it('handles zero', () => {
      expect(emuToPixels(0)).toBe(0);
    });

    it('respects custom DPI', () => {
      // At 72 DPI, 1 inch = 72 pixels
      expect(emuToPixels(914400, 72)).toBe(72);
    });
  });

  describe('pixelsToEmu', () => {
    it('converts 96 pixels to 1 inch of EMUs at default DPI', () => {
      expect(pixelsToEmu(96)).toBe(914400);
    });

    it('round-trips with emuToPixels', () => {
      const emu = 500000;
      expect(pixelsToEmu(emuToPixels(emu))).toBeCloseTo(emu);
    });
  });

  describe('emuToPoints', () => {
    it('converts EMU_PER_POINT to 1 point', () => {
      expect(emuToPoints(12700)).toBe(1);
    });

    it('converts 127000 EMUs to 10 points', () => {
      expect(emuToPoints(127000)).toBe(10);
    });
  });

  describe('pointsToEmu', () => {
    it('converts 1 point to EMU_PER_POINT', () => {
      expect(pointsToEmu(1)).toBe(12700);
    });

    it('round-trips with emuToPoints', () => {
      const points = 24;
      expect(emuToPoints(pointsToEmu(points))).toBe(points);
    });
  });

  describe('pointsToPixels', () => {
    it('converts 72 points to 96 pixels at default DPI', () => {
      expect(pointsToPixels(72)).toBe(96);
    });

    it('converts 12 points to 16 pixels', () => {
      expect(pointsToPixels(12)).toBe(16);
    });

    it('respects custom DPI', () => {
      // At 72 DPI, points = pixels
      expect(pointsToPixels(72, 72)).toBe(72);
    });
  });

  describe('pixelsToPoints', () => {
    it('converts 96 pixels to 72 points at default DPI', () => {
      expect(pixelsToPoints(96)).toBe(72);
    });

    it('round-trips with pointsToPixels', () => {
      const pixels = 48;
      expect(pointsToPixels(pixelsToPoints(pixels))).toBeCloseTo(pixels);
    });
  });

  describe('centipointsToPixels', () => {
    it('converts 1200 centipoints (12pt) to 16 pixels', () => {
      expect(centipointsToPixels(1200)).toBe(16);
    });

    it('converts 7200 centipoints (72pt) to 96 pixels', () => {
      expect(centipointsToPixels(7200)).toBe(96);
    });
  });

  describe('ooxmlPercentToDecimal', () => {
    it('converts 100000 to 1.0 (100%)', () => {
      expect(ooxmlPercentToDecimal(100000)).toBe(1);
    });

    it('converts 50000 to 0.5 (50%)', () => {
      expect(ooxmlPercentToDecimal(50000)).toBe(0.5);
    });

    it('converts 0 to 0', () => {
      expect(ooxmlPercentToDecimal(0)).toBe(0);
    });
  });

  describe('ooxmlAngleToDegrees', () => {
    it('converts 5400000 to 90 degrees', () => {
      expect(ooxmlAngleToDegrees(5400000)).toBe(90);
    });

    it('converts 21600000 to 360 degrees', () => {
      expect(ooxmlAngleToDegrees(21600000)).toBe(360);
    });

    it('converts 0 to 0 degrees', () => {
      expect(ooxmlAngleToDegrees(0)).toBe(0);
    });
  });

  describe('constants', () => {
    it('has correct EMU_PER_INCH', () => {
      expect(EMU_PER_INCH).toBe(914400);
    });

    it('has correct EMU_PER_POINT', () => {
      expect(EMU_PER_POINT).toBe(12700);
    });

    it('has correct DEFAULT_DPI', () => {
      expect(DEFAULT_DPI).toBe(96);
    });
  });
});
