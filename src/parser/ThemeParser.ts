/**
 * Parser for OOXML theme files.
 *
 * Themes define the color scheme and font scheme for a presentation.
 * Located in ppt/theme/theme1.xml.
 */

import type { Theme, ThemeColors, ThemeFonts } from '../core/types';
import { parseXml, findFirstByName, getAttribute } from '../utils/xml';
import { DEFAULT_THEME_COLORS } from '../utils/color';

/**
 * Parses a theme XML file.
 *
 * @param xmlContent - Raw XML content of the theme file
 * @returns Parsed theme object
 *
 * @example
 * const themeXml = archive.getText('ppt/theme/theme1.xml');
 * const theme = parseTheme(themeXml);
 */
export function parseTheme(xmlContent: string): Theme {
  const doc = parseXml(xmlContent);
  const root = doc.documentElement;

  // Get theme name
  const name = getAttribute(root, 'name') || undefined;

  // Parse color scheme
  const colors = parseColorScheme(root);

  // Parse font scheme
  const fonts = parseFontScheme(root);

  return {
    name,
    colors,
    fonts,
  };
}

/**
 * Parses the color scheme from a theme.
 */
function parseColorScheme(root: Element): ThemeColors {
  const colors: ThemeColors = { ...DEFAULT_THEME_COLORS };

  // Find clrScheme element
  const clrScheme = findFirstByName(root, 'clrScheme');
  if (!clrScheme) return colors;

  // Color element mappings
  const colorMappings: [keyof ThemeColors, string][] = [
    ['dark1', 'dk1'],
    ['light1', 'lt1'],
    ['dark2', 'dk2'],
    ['light2', 'lt2'],
    ['accent1', 'accent1'],
    ['accent2', 'accent2'],
    ['accent3', 'accent3'],
    ['accent4', 'accent4'],
    ['accent5', 'accent5'],
    ['accent6', 'accent6'],
    ['hlink', 'hlink'],
    ['folHlink', 'folHlink'],
  ];

  for (const [key, elementName] of colorMappings) {
    const colorElement = findFirstByName(clrScheme, elementName);
    if (colorElement) {
      const hexColor = extractColor(colorElement);
      if (hexColor) {
        colors[key] = hexColor;
      }
    }
  }

  return colors;
}

/**
 * Extracts a hex color from a color element.
 * Handles srgbClr, sysClr, and other color types.
 */
function extractColor(colorElement: Element): string | null {
  // Check for srgbClr (explicit RGB)
  const srgbClr = findFirstByName(colorElement, 'srgbClr');
  if (srgbClr) {
    const val = getAttribute(srgbClr, 'val');
    if (val) return `#${val}`;
  }

  // Check for sysClr (system color)
  const sysClr = findFirstByName(colorElement, 'sysClr');
  if (sysClr) {
    // Use lastClr as it represents the actual color value
    const lastClr = getAttribute(sysClr, 'lastClr');
    if (lastClr) return `#${lastClr}`;

    // Fallback to val which is the system color name
    const val = getAttribute(sysClr, 'val');
    return getSystemColor(val || '');
  }

  return null;
}

/**
 * Maps system color names to hex values.
 */
function getSystemColor(name: string): string {
  const systemColors: Record<string, string> = {
    windowText: '#000000',
    window: '#FFFFFF',
    buttonFace: '#F0F0F0',
    buttonHighlight: '#FFFFFF',
    buttonShadow: '#808080',
    buttonText: '#000000',
    captionText: '#000000',
    grayText: '#808080',
    highlight: '#0066CC',
    highlightText: '#FFFFFF',
    inactiveBorder: '#F4F7FC',
    inactiveCaption: '#BFCDDB',
    inactiveCaptionText: '#434E54',
    infoBackground: '#FFFFE1',
    infoText: '#000000',
    menu: '#F0F0F0',
    menuText: '#000000',
    scrollbar: '#C8C8C8',
    threeDDarkShadow: '#696969',
    threeDFace: '#F0F0F0',
    threeDHighlight: '#FFFFFF',
    threeDLightShadow: '#E3E3E3',
    threeDShadow: '#A0A0A0',
    windowFrame: '#646464',
  };

  return systemColors[name] || '#000000';
}

/**
 * Parses the font scheme from a theme.
 */
function parseFontScheme(root: Element): ThemeFonts {
  const fonts: ThemeFonts = {
    major: 'Calibri Light',
    minor: 'Calibri',
  };

  // Find fontScheme element
  const fontScheme = findFirstByName(root, 'fontScheme');
  if (!fontScheme) return fonts;

  // Parse major font (headings)
  const majorFont = findFirstByName(fontScheme, 'majorFont');
  if (majorFont) {
    const latin = findFirstByName(majorFont, 'latin');
    if (latin) {
      const typeface = getAttribute(latin, 'typeface');
      if (typeface) fonts.major = typeface;
    }
  }

  // Parse minor font (body)
  const minorFont = findFirstByName(fontScheme, 'minorFont');
  if (minorFont) {
    const latin = findFirstByName(minorFont, 'latin');
    if (latin) {
      const typeface = getAttribute(latin, 'typeface');
      if (typeface) fonts.minor = typeface;
    }
  }

  return fonts;
}

/**
 * Creates a default theme when no theme file is present.
 */
export function createDefaultTheme(): Theme {
  return {
    name: 'Default',
    colors: { ...DEFAULT_THEME_COLORS },
    fonts: {
      major: 'Calibri Light',
      minor: 'Calibri',
    },
  };
}
