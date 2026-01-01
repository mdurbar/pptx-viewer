/**
 * Color parsing and manipulation utilities for PPTX.
 *
 * OOXML uses various color formats:
 * - srgbClr: Standard RGB hex (e.g., "FF0000" for red)
 * - schemeClr: Theme color reference (e.g., "accent1")
 * - prstClr: Preset color name (e.g., "red", "blue")
 */

import type { Color, ThemeColors } from '../core/types';

/**
 * Parses a 6-digit hex color string to a Color object.
 *
 * @param hex - Hex color without # (e.g., "FF0000")
 * @param alpha - Opacity from 0 to 1 (default: 1)
 * @returns Color object
 *
 * @example
 * parseHexColor("FF0000"); // => { hex: "#FF0000", alpha: 1 }
 * parseHexColor("00FF00", 0.5); // => { hex: "#00FF00", alpha: 0.5 }
 */
export function parseHexColor(hex: string, alpha: number = 1): Color {
  // Remove # if present
  const cleanHex = hex.replace(/^#/, '').toUpperCase();

  return {
    hex: `#${cleanHex}`,
    alpha: Math.max(0, Math.min(1, alpha)),
  };
}

/**
 * Converts OOXML alpha value to decimal.
 * OOXML uses 0-100000 scale (where 100000 = fully opaque).
 *
 * @param ooxmlAlpha - Alpha value in OOXML format
 * @returns Alpha value from 0 to 1
 */
export function parseOoxmlAlpha(ooxmlAlpha: number): number {
  return ooxmlAlpha / 100000;
}

/**
 * Resolves a theme color reference to a hex color.
 *
 * @param colorRef - Theme color reference (e.g., "accent1", "dk1")
 * @param theme - Theme colors object
 * @returns Hex color string
 */
export function resolveThemeColor(colorRef: string, theme: ThemeColors): string {
  const mapping: Record<string, keyof ThemeColors> = {
    dk1: 'dark1',
    dk2: 'dark2',
    lt1: 'light1',
    lt2: 'light2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hlink: 'hlink',
    folHlink: 'folHlink',
    // Common aliases
    tx1: 'dark1', // Text 1 typically maps to dark1
    tx2: 'dark2', // Text 2 typically maps to dark2
    bg1: 'light1', // Background 1 typically maps to light1
    bg2: 'light2', // Background 2 typically maps to light2
  };

  const themeKey = mapping[colorRef];
  if (themeKey && theme[themeKey]) {
    return theme[themeKey];
  }

  // Return black as fallback
  return '#000000';
}

/**
 * Preset color names and their hex values.
 * These are the standard colors defined in OOXML.
 */
export const PRESET_COLORS: Record<string, string> = {
  aliceBlue: '#F0F8FF',
  antiqueWhite: '#FAEBD7',
  aqua: '#00FFFF',
  aquamarine: '#7FFFD4',
  azure: '#F0FFFF',
  beige: '#F5F5DC',
  bisque: '#FFE4C4',
  black: '#000000',
  blanchedAlmond: '#FFEBCD',
  blue: '#0000FF',
  blueViolet: '#8A2BE2',
  brown: '#A52A2A',
  burlyWood: '#DEB887',
  cadetBlue: '#5F9EA0',
  chartreuse: '#7FFF00',
  chocolate: '#D2691E',
  coral: '#FF7F50',
  cornflowerBlue: '#6495ED',
  cornsilk: '#FFF8DC',
  crimson: '#DC143C',
  cyan: '#00FFFF',
  darkBlue: '#00008B',
  darkCyan: '#008B8B',
  darkGoldenrod: '#B8860B',
  darkGray: '#A9A9A9',
  darkGreen: '#006400',
  darkKhaki: '#BDB76B',
  darkMagenta: '#8B008B',
  darkOliveGreen: '#556B2F',
  darkOrange: '#FF8C00',
  darkOrchid: '#9932CC',
  darkRed: '#8B0000',
  darkSalmon: '#E9967A',
  darkSeaGreen: '#8FBC8F',
  darkSlateBlue: '#483D8B',
  darkSlateGray: '#2F4F4F',
  darkTurquoise: '#00CED1',
  darkViolet: '#9400D3',
  deepPink: '#FF1493',
  deepSkyBlue: '#00BFFF',
  dimGray: '#696969',
  dodgerBlue: '#1E90FF',
  firebrick: '#B22222',
  floralWhite: '#FFFAF0',
  forestGreen: '#228B22',
  fuchsia: '#FF00FF',
  gainsboro: '#DCDCDC',
  ghostWhite: '#F8F8FF',
  gold: '#FFD700',
  goldenrod: '#DAA520',
  gray: '#808080',
  green: '#008000',
  greenYellow: '#ADFF2F',
  honeydew: '#F0FFF0',
  hotPink: '#FF69B4',
  indianRed: '#CD5C5C',
  indigo: '#4B0082',
  ivory: '#FFFFF0',
  khaki: '#F0E68C',
  lavender: '#E6E6FA',
  lavenderBlush: '#FFF0F5',
  lawnGreen: '#7CFC00',
  lemonChiffon: '#FFFACD',
  lightBlue: '#ADD8E6',
  lightCoral: '#F08080',
  lightCyan: '#E0FFFF',
  lightGoldenrodYellow: '#FAFAD2',
  lightGray: '#D3D3D3',
  lightGreen: '#90EE90',
  lightPink: '#FFB6C1',
  lightSalmon: '#FFA07A',
  lightSeaGreen: '#20B2AA',
  lightSkyBlue: '#87CEFA',
  lightSlateGray: '#778899',
  lightSteelBlue: '#B0C4DE',
  lightYellow: '#FFFFE0',
  lime: '#00FF00',
  limeGreen: '#32CD32',
  linen: '#FAF0E6',
  magenta: '#FF00FF',
  maroon: '#800000',
  mediumAquamarine: '#66CDAA',
  mediumBlue: '#0000CD',
  mediumOrchid: '#BA55D3',
  mediumPurple: '#9370DB',
  mediumSeaGreen: '#3CB371',
  mediumSlateBlue: '#7B68EE',
  mediumSpringGreen: '#00FA9A',
  mediumTurquoise: '#48D1CC',
  mediumVioletRed: '#C71585',
  midnightBlue: '#191970',
  mintCream: '#F5FFFA',
  mistyRose: '#FFE4E1',
  moccasin: '#FFE4B5',
  navajoWhite: '#FFDEAD',
  navy: '#000080',
  oldLace: '#FDF5E6',
  olive: '#808000',
  oliveDrab: '#6B8E23',
  orange: '#FFA500',
  orangeRed: '#FF4500',
  orchid: '#DA70D6',
  paleGoldenrod: '#EEE8AA',
  paleGreen: '#98FB98',
  paleTurquoise: '#AFEEEE',
  paleVioletRed: '#DB7093',
  papayaWhip: '#FFEFD5',
  peachPuff: '#FFDAB9',
  peru: '#CD853F',
  pink: '#FFC0CB',
  plum: '#DDA0DD',
  powderBlue: '#B0E0E6',
  purple: '#800080',
  red: '#FF0000',
  rosyBrown: '#BC8F8F',
  royalBlue: '#4169E1',
  saddleBrown: '#8B4513',
  salmon: '#FA8072',
  sandyBrown: '#F4A460',
  seaGreen: '#2E8B57',
  seaShell: '#FFF5EE',
  sienna: '#A0522D',
  silver: '#C0C0C0',
  skyBlue: '#87CEEB',
  slateBlue: '#6A5ACD',
  slateGray: '#708090',
  snow: '#FFFAFA',
  springGreen: '#00FF7F',
  steelBlue: '#4682B4',
  tan: '#D2B48C',
  teal: '#008080',
  thistle: '#D8BFD8',
  tomato: '#FF6347',
  turquoise: '#40E0D0',
  violet: '#EE82EE',
  wheat: '#F5DEB3',
  white: '#FFFFFF',
  whiteSmoke: '#F5F5F5',
  yellow: '#FFFF00',
  yellowGreen: '#9ACD32',
};

/**
 * Resolves a preset color name to hex.
 *
 * @param name - Preset color name
 * @returns Hex color string or black as fallback
 */
export function resolvePresetColor(name: string): string {
  return PRESET_COLORS[name] || '#000000';
}

/**
 * Applies luminance modification to a color.
 * Used for tint/shade adjustments in OOXML.
 *
 * @param hex - Hex color string
 * @param lumMod - Luminance modifier (100000 = no change)
 * @param lumOff - Luminance offset (0 = no offset)
 * @returns Modified hex color
 */
export function applyLuminance(hex: string, lumMod: number = 100000, lumOff: number = 0): string {
  const cleanHex = hex.replace('#', '');
  const r = parseInt(cleanHex.slice(0, 2), 16);
  const g = parseInt(cleanHex.slice(2, 4), 16);
  const b = parseInt(cleanHex.slice(4, 6), 16);

  const mod = lumMod / 100000;
  const off = (lumOff / 100000) * 255;

  const newR = Math.round(Math.min(255, Math.max(0, r * mod + off)));
  const newG = Math.round(Math.min(255, Math.max(0, g * mod + off)));
  const newB = Math.round(Math.min(255, Math.max(0, b * mod + off)));

  return `#${newR.toString(16).padStart(2, '0')}${newG.toString(16).padStart(2, '0')}${newB.toString(16).padStart(2, '0')}`.toUpperCase();
}

/**
 * Converts a Color object to a CSS color string.
 *
 * @param color - Color object
 * @returns CSS color string (hex or rgba)
 */
export function colorToCss(color: Color): string {
  if (color.alpha === 1) {
    return color.hex;
  }

  const cleanHex = color.hex.replace('#', '');
  const r = parseInt(cleanHex.slice(0, 2), 16);
  const g = parseInt(cleanHex.slice(2, 4), 16);
  const b = parseInt(cleanHex.slice(4, 6), 16);

  return `rgba(${r}, ${g}, ${b}, ${color.alpha})`;
}

/**
 * Default theme colors (Office theme).
 */
export const DEFAULT_THEME_COLORS: ThemeColors = {
  dark1: '#000000',
  light1: '#FFFFFF',
  dark2: '#44546A',
  light2: '#E7E6E6',
  accent1: '#4472C4',
  accent2: '#ED7D31',
  accent3: '#A5A5A5',
  accent4: '#FFC000',
  accent5: '#5B9BD5',
  accent6: '#70AD47',
  hlink: '#0563C1',
  folHlink: '#954F72',
};
