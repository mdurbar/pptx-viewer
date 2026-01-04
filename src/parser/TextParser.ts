/**
 * Parser for text content in OOXML.
 *
 * Text in PPTX is structured as:
 * - txBody: Text body container
 *   - bodyPr: Body properties (margins, anchoring)
 *   - p: Paragraphs
 *     - pPr: Paragraph properties (alignment, spacing, bullets)
 *     - r: Runs (text with consistent formatting)
 *       - rPr: Run properties (font, size, color, bold, etc.)
 *       - t: Text content
 */

import type {
  TextBody,
  TextAutofit,
  Paragraph,
  TextRun,
  BulletStyle,
  Color,
  ThemeColors,
  TextGlow,
  TextReflection,
  TextOutline,
} from '../core/types';
import type { RelationshipMap } from './RelationshipParser';
import {
  findChildByName,
  findChildrenByName,
  getAttribute,
  getNumberAttribute,
  getBooleanAttribute,
  getTextContent,
} from '../utils/xml';
import { emuToPixels, centipointsToPixels } from '../utils/units';
import { parseHexColor, resolveThemeColor, resolvePresetColor, parseOoxmlAlpha } from '../utils/color';

/**
 * Parses a text body element (txBody).
 *
 * @param txBody - The txBody XML element
 * @param themeColors - Theme colors for resolving color references
 * @param relationships - Optional relationships for resolving hyperlinks
 * @returns Parsed text body
 */
export function parseTextBody(
  txBody: Element,
  themeColors: ThemeColors,
  relationships?: RelationshipMap
): TextBody {
  const paragraphs: Paragraph[] = [];

  // Parse body properties
  const bodyPr = findChildByName(txBody, 'bodyPr');
  const padding = parseBodyPadding(bodyPr);
  const verticalAlign = parseVerticalAlign(bodyPr);
  const autofit = parseAutofit(bodyPr);

  // Parse paragraphs
  const pElements = findChildrenByName(txBody, 'p');
  for (const pElement of pElements) {
    const paragraph = parseParagraph(pElement, themeColors, relationships);
    paragraphs.push(paragraph);
  }

  return {
    paragraphs,
    verticalAlign,
    padding,
    autofit,
  };
}

/**
 * Parses body padding/margins.
 */
function parseBodyPadding(bodyPr: Element | null): TextBody['padding'] {
  if (!bodyPr) {
    return { top: 4, right: 8, bottom: 4, left: 8 };
  }

  return {
    top: emuToPixels(getNumberAttribute(bodyPr, 'tIns', 45720)),
    right: emuToPixels(getNumberAttribute(bodyPr, 'rIns', 91440)),
    bottom: emuToPixels(getNumberAttribute(bodyPr, 'bIns', 45720)),
    left: emuToPixels(getNumberAttribute(bodyPr, 'lIns', 91440)),
  };
}

/**
 * Parses vertical text alignment.
 */
function parseVerticalAlign(bodyPr: Element | null): TextBody['verticalAlign'] {
  if (!bodyPr) return 'top';

  const anchor = getAttribute(bodyPr, 'anchor');

  switch (anchor) {
    case 'ctr':
      return 'middle';
    case 'b':
      return 'bottom';
    case 't':
    default:
      return 'top';
  }
}

/**
 * Parses text autofit settings.
 */
function parseAutofit(bodyPr: Element | null): TextAutofit | undefined {
  if (!bodyPr) return undefined;

  // Check for normal autofit (shrink text to fit)
  const normAutofit = findChildByName(bodyPr, 'normAutofit');
  if (normAutofit) {
    const fontScale = getNumberAttribute(normAutofit, 'fontScale', 100000) / 100000;
    const lnSpcReduction = getNumberAttribute(normAutofit, 'lnSpcReduction', 0) / 100000;

    return {
      type: 'normal',
      fontScale,
      lineSpacingReduction: lnSpcReduction,
    };
  }

  // Check for shape autofit (resize shape to fit text)
  const spAutoFit = findChildByName(bodyPr, 'spAutoFit');
  if (spAutoFit) {
    return {
      type: 'shape',
    };
  }

  // Check for no autofit
  const noAutofit = findChildByName(bodyPr, 'noAutofit');
  if (noAutofit) {
    return {
      type: 'none',
    };
  }

  return undefined;
}

/**
 * Parses a paragraph element.
 */
function parseParagraph(
  pElement: Element,
  themeColors: ThemeColors,
  relationships?: RelationshipMap
): Paragraph {
  const runs: TextRun[] = [];

  // Parse paragraph properties
  const pPr = findChildByName(pElement, 'pPr');
  const align = parseParagraphAlign(pPr);
  const lineSpacing = parseLineSpacing(pPr);
  const spaceBefore = parseSpacing(pPr, 'spcBef');
  const spaceAfter = parseSpacing(pPr, 'spcAft');
  const bullet = parseBullet(pPr, themeColors);
  const level = getNumberAttribute(pPr || pElement, 'lvl', 0);
  const { marginLeft, indent } = parseIndentation(pPr);

  // Parse runs
  const rElements = findChildrenByName(pElement, 'r');
  for (const rElement of rElements) {
    const run = parseTextRun(rElement, themeColors, relationships);
    runs.push(run);
  }

  // Handle line breaks (br elements)
  const brElements = findChildrenByName(pElement, 'br');
  if (brElements.length > 0 && runs.length === 0) {
    runs.push({ text: '' });
  }

  // Handle fields (fld elements) like slide numbers, dates
  const fldElements = findChildrenByName(pElement, 'fld');
  for (const fldElement of fldElements) {
    const tElement = findChildByName(fldElement, 't');
    if (tElement) {
      runs.push({ text: getTextContent(tElement) });
    }
  }

  return {
    runs,
    align,
    lineSpacing,
    spaceBefore,
    spaceAfter,
    bullet,
    level,
    marginLeft,
    indent,
  };
}

/**
 * Parses paragraph alignment.
 */
function parseParagraphAlign(pPr: Element | null): Paragraph['align'] {
  if (!pPr) return 'left';

  const algn = getAttribute(pPr, 'algn');

  switch (algn) {
    case 'ctr':
      return 'center';
    case 'r':
      return 'right';
    case 'just':
      return 'justify';
    case 'l':
    default:
      return 'left';
  }
}

/**
 * Parses line spacing.
 */
function parseLineSpacing(pPr: Element | null): number | undefined {
  if (!pPr) return undefined;

  const lnSpc = findChildByName(pPr, 'lnSpc');
  if (!lnSpc) return undefined;

  // Check for percentage-based spacing
  const spcPct = findChildByName(lnSpc, 'spcPct');
  if (spcPct) {
    const val = getNumberAttribute(spcPct, 'val', 100000);
    return val / 100000; // Convert to multiplier
  }

  // Check for point-based spacing
  const spcPts = findChildByName(lnSpc, 'spcPts');
  if (spcPts) {
    const val = getNumberAttribute(spcPts, 'val', 1200);
    return centipointsToPixels(val) / 16; // Approximate line height ratio
  }

  return undefined;
}

/**
 * Parses spacing before/after paragraph.
 */
function parseSpacing(pPr: Element | null, elementName: string): number | undefined {
  if (!pPr) return undefined;

  const spcElement = findChildByName(pPr, elementName);
  if (!spcElement) return undefined;

  // Check for point-based spacing
  const spcPts = findChildByName(spcElement, 'spcPts');
  if (spcPts) {
    const val = getNumberAttribute(spcPts, 'val', 0);
    return centipointsToPixels(val);
  }

  // Check for percentage-based spacing (relative to line height)
  const spcPct = findChildByName(spcElement, 'spcPct');
  if (spcPct) {
    const val = getNumberAttribute(spcPct, 'val', 0);
    return (val / 100000) * 16; // Approximate based on default line height
  }

  return undefined;
}

/**
 * Parses bullet point style.
 */
function parseBullet(pPr: Element | null, themeColors: ThemeColors): BulletStyle | undefined {
  if (!pPr) return undefined;

  // Check for no bullet
  const buNone = findChildByName(pPr, 'buNone');
  if (buNone) return undefined;

  let bullet: BulletStyle | undefined;

  // Check for character bullet
  const buChar = findChildByName(pPr, 'buChar');
  if (buChar) {
    bullet = {
      type: 'bullet',
      char: getAttribute(buChar, 'char') || '•',
    };
  }

  // Check for auto-numbered bullet
  const buAutoNum = findChildByName(pPr, 'buAutoNum');
  if (buAutoNum) {
    const startAt = getNumberAttribute(buAutoNum, 'startAt', 1);
    const numberType = getAttribute(buAutoNum, 'type') || 'arabicPeriod';
    bullet = {
      type: 'number',
      startAt,
      numberType,
    };
  }

  // If no bullet type found, return undefined
  if (!bullet) return undefined;

  // Parse bullet color
  const buClr = findChildByName(pPr, 'buClr');
  if (buClr) {
    const color = parseColorElement(buClr, themeColors);
    if (color) {
      bullet.color = color;
    }
  }

  // Parse bullet font
  const buFont = findChildByName(pPr, 'buFont');
  if (buFont) {
    const typeface = getAttribute(buFont, 'typeface');
    if (typeface) {
      bullet.font = typeface;
    }
  }

  // Parse bullet size
  const buSzPct = findChildByName(pPr, 'buSzPct');
  if (buSzPct) {
    const val = getNumberAttribute(buSzPct, 'val', 100000);
    bullet.sizePercent = val / 1000; // Convert from 1000ths to percentage
  }

  return bullet;
}

/**
 * Parses paragraph indentation (left margin and first line indent).
 */
function parseIndentation(pPr: Element | null): { marginLeft?: number; indent?: number } {
  if (!pPr) return {};

  const result: { marginLeft?: number; indent?: number } = {};

  // Left margin (marL) in EMUs
  const marL = getNumberAttribute(pPr, 'marL', 0);
  if (marL > 0) {
    result.marginLeft = emuToPixels(marL);
  }

  // First line indent (indent) in EMUs - negative for hanging indent
  const indentAttr = getAttribute(pPr, 'indent');
  if (indentAttr) {
    const indentVal = parseInt(indentAttr, 10);
    if (!isNaN(indentVal)) {
      result.indent = emuToPixels(indentVal);
    }
  }

  return result;
}

/**
 * Parses a text run element.
 */
function parseTextRun(
  rElement: Element,
  themeColors: ThemeColors,
  relationships?: RelationshipMap
): TextRun {
  // Get text content
  const tElement = findChildByName(rElement, 't');
  const text = tElement ? getTextContent(tElement) : '';

  // Parse run properties
  const rPr = findChildByName(rElement, 'rPr');
  const formatting = parseRunProperties(rPr, themeColors, relationships);

  return {
    text,
    ...formatting,
  };
}

/**
 * Parses run properties (font, size, color, etc.).
 */
function parseRunProperties(
  rPr: Element | null,
  themeColors: ThemeColors,
  relationships?: RelationshipMap
): Omit<TextRun, 'text'> {
  if (!rPr) {
    return {};
  }

  const result: Omit<TextRun, 'text'> = {};

  // Font size (in hundredths of a point)
  const sz = getNumberAttribute(rPr, 'sz', 0);
  if (sz > 0) {
    result.fontSize = centipointsToPixels(sz);
  }

  // Bold
  if (getBooleanAttribute(rPr, 'b')) {
    result.bold = true;
  }

  // Italic
  if (getBooleanAttribute(rPr, 'i')) {
    result.italic = true;
  }

  // Underline
  const u = getAttribute(rPr, 'u');
  if (u && u !== 'none') {
    result.underline = true;
  }

  // Strikethrough
  const strike = getAttribute(rPr, 'strike');
  if (strike && strike !== 'noStrike') {
    result.strikethrough = true;
  }

  // Baseline (subscript/superscript)
  const baselineAttr = getAttribute(rPr, 'baseline');
  if (baselineAttr) {
    const baselineVal = parseInt(baselineAttr, 10);
    if (!isNaN(baselineVal) && baselineVal !== 0) {
      // Convert from 1000ths of percent to percent
      result.baseline = baselineVal / 1000;
    }
  }

  // Character spacing (in hundredths of a point, like font size)
  const spc = getNumberAttribute(rPr, 'spc', 0);
  if (spc !== 0) {
    result.characterSpacing = centipointsToPixels(spc);
  }

  // Text capitalization
  const cap = getAttribute(rPr, 'cap');
  if (cap === 'all') {
    result.capitalization = 'allCaps';
  } else if (cap === 'small') {
    result.capitalization = 'smallCaps';
  }

  // Font family
  const latin = findChildByName(rPr, 'latin');
  if (latin) {
    const typeface = getAttribute(latin, 'typeface');
    if (typeface) {
      result.fontFamily = typeface;
    }
  }

  // Color
  const solidFill = findChildByName(rPr, 'solidFill');
  if (solidFill) {
    const color = parseColorElement(solidFill, themeColors);
    if (color) {
      result.color = color;
    }
  }

  // Hyperlink
  const hlinkClick = findChildByName(rPr, 'hlinkClick');
  if (hlinkClick) {
    const rId = getAttribute(hlinkClick, 'r:id');
    if (rId && relationships) {
      // Resolve the actual URL from relationships
      const rel = relationships.get(rId);
      if (rel && rel.target) {
        result.link = rel.target;
      }
    }
  }

  // Highlight/background color
  const highlight = findChildByName(rPr, 'highlight');
  if (highlight) {
    const highlightColor = parseColorElement(highlight, themeColors);
    if (highlightColor) {
      result.highlight = highlightColor;
    }
  }

  // Text effects (glow, reflection)
  const effectLst = findChildByName(rPr, 'effectLst');
  if (effectLst) {
    // Parse glow effect
    const glow = parseTextGlow(effectLst, themeColors);
    if (glow) {
      result.glow = glow;
    }

    // Parse reflection effect
    const reflection = parseTextReflection(effectLst);
    if (reflection) {
      result.reflection = reflection;
    }
  }

  // Text outline (ln element)
  const outline = parseTextOutline(rPr, themeColors);
  if (outline) {
    result.outline = outline;
  }

  return result;
}

/**
 * Parses a color from a fill element.
 */
export function parseColorElement(fillElement: Element, themeColors: ThemeColors): Color | null {
  // Check for srgbClr (explicit RGB)
  const srgbClr = findChildByName(fillElement, 'srgbClr');
  if (srgbClr) {
    const val = getAttribute(srgbClr, 'val');
    if (val) {
      const alpha = parseAlphaFromElement(srgbClr);
      return parseHexColor(val, alpha);
    }
  }

  // Check for schemeClr (theme color reference)
  const schemeClr = findChildByName(fillElement, 'schemeClr');
  if (schemeClr) {
    const val = getAttribute(schemeClr, 'val');
    if (val) {
      const hex = resolveThemeColor(val, themeColors);
      const alpha = parseAlphaFromElement(schemeClr);
      return parseHexColor(hex.replace('#', ''), alpha);
    }
  }

  // Check for prstClr (preset color)
  const prstClr = findChildByName(fillElement, 'prstClr');
  if (prstClr) {
    const val = getAttribute(prstClr, 'val');
    if (val) {
      const hex = resolvePresetColor(val);
      const alpha = parseAlphaFromElement(prstClr);
      return parseHexColor(hex.replace('#', ''), alpha);
    }
  }

  return null;
}

/**
 * Parses alpha value from a color element's children.
 */
function parseAlphaFromElement(colorElement: Element): number {
  const alpha = findChildByName(colorElement, 'alpha');
  if (alpha) {
    const val = getNumberAttribute(alpha, 'val', 100000);
    return parseOoxmlAlpha(val);
  }
  return 1;
}

/**
 * Parses text glow effect from effectLst.
 *
 * Glow is defined as:
 * <a:glow rad="...">
 *   <a:srgbClr val="..."/>
 * </a:glow>
 */
function parseTextGlow(effectLst: Element, themeColors: ThemeColors): TextGlow | null {
  const glow = findChildByName(effectLst, 'glow');
  if (!glow) return null;

  // Parse radius (in EMUs)
  const rad = getNumberAttribute(glow, 'rad', 0);
  if (rad === 0) return null;

  // Parse color
  const color = parseColorElement(glow, themeColors);
  if (!color) return null;

  return {
    radius: emuToPixels(rad),
    color,
  };
}

/**
 * Parses text reflection effect from effectLst.
 *
 * Reflection is defined as:
 * <a:reflection blurRad="..." stA="..." endA="..." dist="..." dir="..."
 *               fadeDir="..." sy="..." kx="..." algn="..."/>
 */
function parseTextReflection(effectLst: Element): TextReflection | null {
  const reflection = findChildByName(effectLst, 'reflection');
  if (!reflection) return null;

  // Parse blur radius (in EMUs)
  const blurRad = getNumberAttribute(reflection, 'blurRad', 0);

  // Parse start and end opacity (in 1000ths of percent)
  const stA = getNumberAttribute(reflection, 'stA', 100000);
  const endA = getNumberAttribute(reflection, 'endA', 0);

  // Parse distance (in EMUs)
  const dist = getNumberAttribute(reflection, 'dist', 0);

  // Parse direction (in 60000ths of a degree)
  const dir = getNumberAttribute(reflection, 'dir', 0);

  // Parse fade direction (in 60000ths of a degree)
  const fadeDir = getNumberAttribute(reflection, 'fadeDir', 5400000);

  // Parse vertical scale (in 1000ths of percent)
  const sy = getNumberAttribute(reflection, 'sy', 100000);

  // Parse horizontal skew (in 60000ths of a degree)
  const kx = getNumberAttribute(reflection, 'kx', 0);

  // Parse alignment
  const algn = getAttribute(reflection, 'algn') || 'b';

  return {
    blurRadius: emuToPixels(blurRad),
    startOpacity: stA / 100000,
    endOpacity: endA / 100000,
    distance: emuToPixels(dist),
    direction: dir / 60000,
    fadeDirection: fadeDir / 60000,
    scaleY: sy / 1000,
    skewX: kx / 60000,
    align: algn === 't' ? 'top' : 'bottom',
  };
}

/**
 * Parses text outline/stroke from run properties.
 *
 * Text outline is defined as:
 * <a:ln w="...">
 *   <a:solidFill>
 *     <a:srgbClr val="..."/>
 *   </a:solidFill>
 * </a:ln>
 */
function parseTextOutline(rPr: Element, themeColors: ThemeColors): TextOutline | null {
  const ln = findChildByName(rPr, 'ln');
  if (!ln) return null;

  // Parse width (in EMUs)
  const w = getNumberAttribute(ln, 'w', 0);
  if (w === 0) return null;

  // Parse color from solidFill
  const solidFill = findChildByName(ln, 'solidFill');
  if (!solidFill) return null;

  const color = parseColorElement(solidFill, themeColors);
  if (!color) return null;

  return {
    color,
    width: emuToPixels(w),
  };
}
