/**
 * Unit conversion utilities for PPTX.
 *
 * PowerPoint uses English Metric Units (EMUs) internally.
 * - 914400 EMU = 1 inch
 * - 12700 EMU = 1 point
 * - Standard slide size: 9144000 x 6858000 EMU (10" x 7.5")
 */

/** EMUs per inch */
export const EMU_PER_INCH = 914400;

/** EMUs per point */
export const EMU_PER_POINT = 12700;

/** Points per inch */
export const POINTS_PER_INCH = 72;

/** Default DPI for screen rendering */
export const DEFAULT_DPI = 96;

/**
 * Converts EMUs to pixels at a given DPI.
 *
 * @param emu - Value in English Metric Units
 * @param dpi - Dots per inch (default: 96)
 * @returns Value in pixels
 *
 * @example
 * emuToPixels(914400); // => 96 (1 inch at 96 DPI)
 * emuToPixels(9144000); // => 960 (10 inches at 96 DPI)
 */
export function emuToPixels(emu: number, dpi: number = DEFAULT_DPI): number {
  return (emu / EMU_PER_INCH) * dpi;
}

/**
 * Converts pixels to EMUs at a given DPI.
 *
 * @param pixels - Value in pixels
 * @param dpi - Dots per inch (default: 96)
 * @returns Value in EMUs
 */
export function pixelsToEmu(pixels: number, dpi: number = DEFAULT_DPI): number {
  return (pixels / dpi) * EMU_PER_INCH;
}

/**
 * Converts EMUs to points.
 *
 * @param emu - Value in English Metric Units
 * @returns Value in points
 *
 * @example
 * emuToPoints(12700); // => 1
 * emuToPoints(127000); // => 10
 */
export function emuToPoints(emu: number): number {
  return emu / EMU_PER_POINT;
}

/**
 * Converts points to EMUs.
 *
 * @param points - Value in points
 * @returns Value in EMUs
 */
export function pointsToEmu(points: number): number {
  return points * EMU_PER_POINT;
}

/**
 * Converts points to pixels.
 *
 * @param points - Value in points
 * @param dpi - Dots per inch (default: 96)
 * @returns Value in pixels
 *
 * @example
 * pointsToPixels(72); // => 96 (72pt = 1 inch = 96px at 96 DPI)
 */
export function pointsToPixels(points: number, dpi: number = DEFAULT_DPI): number {
  return (points / POINTS_PER_INCH) * dpi;
}

/**
 * Converts pixels to points.
 *
 * @param pixels - Value in pixels
 * @param dpi - Dots per inch (default: 96)
 * @returns Value in points
 */
export function pixelsToPoints(pixels: number, dpi: number = DEFAULT_DPI): number {
  return (pixels / dpi) * POINTS_PER_INCH;
}

/**
 * Converts hundredths of a point to pixels.
 * OOXML often uses "centipoints" (1/100th of a point).
 *
 * @param centipoints - Value in 1/100ths of a point
 * @param dpi - Dots per inch (default: 96)
 * @returns Value in pixels
 *
 * @example
 * centipointsToPixels(1200); // => 16 (12pt at 96 DPI)
 */
export function centipointsToPixels(centipoints: number, dpi: number = DEFAULT_DPI): number {
  return pointsToPixels(centipoints / 100, dpi);
}

/**
 * Converts a percentage string (e.g., "50000") to a decimal.
 * OOXML uses 1/1000ths of a percent (so 100% = 100000).
 *
 * @param value - Percentage value in OOXML format
 * @returns Decimal value (0-1 for 0-100%)
 *
 * @example
 * ooxmlPercentToDecimal(50000); // => 0.5
 * ooxmlPercentToDecimal(100000); // => 1.0
 */
export function ooxmlPercentToDecimal(value: number): number {
  return value / 100000;
}

/**
 * Converts an angle in 60,000ths of a degree to degrees.
 * OOXML uses this format for rotation angles.
 *
 * @param angle - Angle in 60,000ths of a degree
 * @returns Angle in degrees
 *
 * @example
 * ooxmlAngleToDegrees(5400000); // => 90
 */
export function ooxmlAngleToDegrees(angle: number): number {
  return angle / 60000;
}

/**
 * Standard slide dimensions in EMUs.
 */
export const STANDARD_SLIDE_SIZES = {
  /** Standard 4:3 (10" x 7.5") */
  STANDARD: { width: 9144000, height: 6858000 },
  /** Widescreen 16:9 (13.333" x 7.5") */
  WIDESCREEN: { width: 12192000, height: 6858000 },
  /** Widescreen 16:10 (10" x 6.25") */
  WIDESCREEN_16_10: { width: 9144000, height: 5715000 },
} as const;
