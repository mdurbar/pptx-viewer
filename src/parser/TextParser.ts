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
  Paragraph,
  TextRun,
  BulletStyle,
  Color,
  ThemeColors,
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
  const bullet = parseBullet(pPr);
  const level = getNumberAttribute(pPr || pElement, 'lvl', 0);

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
function parseBullet(pPr: Element | null): BulletStyle | undefined {
  if (!pPr) return undefined;

  // Check for no bullet
  const buNone = findChildByName(pPr, 'buNone');
  if (buNone) return undefined;

  // Check for character bullet
  const buChar = findChildByName(pPr, 'buChar');
  if (buChar) {
    return {
      type: 'bullet',
      char: getAttribute(buChar, 'char') || '•',
    };
  }

  // Check for auto-numbered bullet
  const buAutoNum = findChildByName(pPr, 'buAutoNum');
  if (buAutoNum) {
    const startAt = getNumberAttribute(buAutoNum, 'startAt', 1);
    return {
      type: 'number',
      startAt,
    };
  }

  return undefined;
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
