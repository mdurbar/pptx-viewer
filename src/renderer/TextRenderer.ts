/**
 * Renderer for text content.
 *
 * Converts parsed TextBody into HTML elements with proper styling.
 */

import type { TextBody, Paragraph, TextRun, Color, BulletStyle, TextAutofit } from '../core/types';
import { colorToCss } from '../utils/color';

/**
 * Tracks numbering state for lists across paragraphs.
 */
interface NumberingState {
  /** Current number for each level */
  numbers: Map<number, number>;
  /** Last bullet type seen at each level */
  lastBulletType: Map<number, string>;
}

/**
 * Autofit context passed to paragraph and run renderers.
 */
interface AutofitContext {
  /** Font scale multiplier (1 = 100%) */
  fontScale: number;
  /** Line spacing reduction multiplier (0 = no reduction) */
  lineSpacingReduction: number;
}

/**
 * Renders a text body to HTML.
 *
 * @param text - Parsed text body
 * @param container - Container element to render into
 */
export function renderTextBody(text: TextBody, container: HTMLElement): void {
  // Apply container styles
  container.style.display = 'flex';
  container.style.flexDirection = 'column';
  container.style.overflow = 'hidden';
  container.style.wordWrap = 'break-word';
  container.style.whiteSpace = 'pre-wrap';

  // Apply vertical alignment
  switch (text.verticalAlign) {
    case 'middle':
      container.style.justifyContent = 'center';
      break;
    case 'bottom':
      container.style.justifyContent = 'flex-end';
      break;
    default:
      container.style.justifyContent = 'flex-start';
  }

  // Apply padding
  if (text.padding) {
    container.style.padding = `${text.padding.top}px ${text.padding.right}px ${text.padding.bottom}px ${text.padding.left}px`;
  }

  // Track numbering state across paragraphs
  const numberingState: NumberingState = {
    numbers: new Map(),
    lastBulletType: new Map(),
  };

  // Get autofit context
  const autofitContext: AutofitContext = {
    fontScale: text.autofit?.fontScale ?? 1,
    lineSpacingReduction: text.autofit?.lineSpacingReduction ?? 0,
  };

  // Render paragraphs
  for (const paragraph of text.paragraphs) {
    const pElement = renderParagraph(paragraph, numberingState, autofitContext);
    container.appendChild(pElement);
  }
}

/**
 * Renders a paragraph to HTML.
 */
function renderParagraph(
  paragraph: Paragraph,
  numberingState: NumberingState,
  autofitContext: AutofitContext
): HTMLElement {
  const p = document.createElement('p');
  p.style.margin = '0';
  p.style.padding = '0';
  p.style.display = 'flex';
  p.style.alignItems = 'baseline';

  // Apply alignment
  switch (paragraph.align) {
    case 'center':
      p.style.justifyContent = 'center';
      break;
    case 'right':
      p.style.justifyContent = 'flex-end';
      break;
    default:
      p.style.justifyContent = 'flex-start';
  }

  // Apply line spacing with autofit reduction
  if (paragraph.lineSpacing) {
    const reducedSpacing = paragraph.lineSpacing * (1 - autofitContext.lineSpacingReduction);
    p.style.lineHeight = String(Math.max(reducedSpacing, 0.8)); // Don't go below 0.8
  }

  // Apply spacing
  if (paragraph.spaceBefore) {
    p.style.marginTop = `${paragraph.spaceBefore}px`;
  }
  if (paragraph.spaceAfter) {
    p.style.marginBottom = `${paragraph.spaceAfter}px`;
  }

  // Calculate indentation
  const level = paragraph.level || 0;
  let leftMargin = paragraph.marginLeft ?? (level * 36); // Default 36px per level
  const hangingIndent = paragraph.indent ?? (paragraph.bullet ? -18 : 0); // Default hanging indent for bullets

  // Apply left margin
  if (leftMargin > 0) {
    p.style.marginLeft = `${leftMargin}px`;
  }

  // Add bullet point
  if (paragraph.bullet) {
    const bulletSpan = document.createElement('span');
    bulletSpan.style.flexShrink = '0';
    bulletSpan.style.display = 'inline-block';

    // Apply hanging indent width to bullet
    const bulletWidth = Math.abs(hangingIndent) || 18;
    bulletSpan.style.width = `${bulletWidth}px`;
    bulletSpan.style.marginLeft = hangingIndent < 0 ? `${hangingIndent}px` : '0';
    bulletSpan.style.textAlign = 'left';

    // Apply bullet styling
    if (paragraph.bullet.font) {
      bulletSpan.style.fontFamily = `"${paragraph.bullet.font}", sans-serif`;
    }
    if (paragraph.bullet.color) {
      bulletSpan.style.color = colorToCss(paragraph.bullet.color);
    }
    if (paragraph.bullet.sizePercent) {
      bulletSpan.style.fontSize = `${paragraph.bullet.sizePercent}%`;
    }

    if (paragraph.bullet.type === 'bullet') {
      bulletSpan.textContent = paragraph.bullet.char || '•';
      // Reset numbering when we see a bullet
      numberingState.numbers.delete(level);
    } else {
      // Numbered list - track and increment
      const bulletKey = `${level}-${paragraph.bullet.numberType || 'arabicPeriod'}`;
      const lastType = numberingState.lastBulletType.get(level);

      // Reset if bullet type changed or starting new list
      if (lastType !== bulletKey) {
        numberingState.numbers.set(level, paragraph.bullet.startAt || 1);
        numberingState.lastBulletType.set(level, bulletKey);
      }

      const currentNumber = numberingState.numbers.get(level) || paragraph.bullet.startAt || 1;
      bulletSpan.textContent = formatBulletNumber(currentNumber, paragraph.bullet.numberType);

      // Increment for next paragraph
      numberingState.numbers.set(level, currentNumber + 1);

      // Reset deeper levels
      for (const [l] of numberingState.numbers) {
        if (l > level) {
          numberingState.numbers.delete(l);
          numberingState.lastBulletType.delete(l);
        }
      }
    }

    p.appendChild(bulletSpan);
  }

  // Create a wrapper for text content
  const textWrapper = document.createElement('span');
  textWrapper.style.flex = '1';
  textWrapper.style.minWidth = '0'; // Allow text to shrink

  // Render text runs
  for (const run of paragraph.runs) {
    const runElement = renderTextRun(run, autofitContext);
    textWrapper.appendChild(runElement);
  }

  // Empty paragraph - add a line break to maintain height
  if (paragraph.runs.length === 0 && !paragraph.bullet) {
    textWrapper.innerHTML = '&nbsp;';
  }

  p.appendChild(textWrapper);

  return p;
}

/**
 * Formats a number according to the bullet number type.
 */
function formatBulletNumber(num: number, numberType?: string): string {
  switch (numberType) {
    case 'alphaLcParenBoth':
      return `(${toAlpha(num, false)})`;
    case 'alphaLcParenR':
      return `${toAlpha(num, false)})`;
    case 'alphaLcPeriod':
      return `${toAlpha(num, false)}.`;
    case 'alphaUcParenBoth':
      return `(${toAlpha(num, true)})`;
    case 'alphaUcParenR':
      return `${toAlpha(num, true)})`;
    case 'alphaUcPeriod':
      return `${toAlpha(num, true)}.`;
    case 'arabicParenBoth':
      return `(${num})`;
    case 'arabicParenR':
      return `${num})`;
    case 'arabicPeriod':
    case 'arabic':
    default:
      return `${num}.`;
    case 'arabicPlain':
      return `${num}`;
    case 'romanLcParenBoth':
      return `(${toRoman(num, false)})`;
    case 'romanLcParenR':
      return `${toRoman(num, false)})`;
    case 'romanLcPeriod':
      return `${toRoman(num, false)}.`;
    case 'romanUcParenBoth':
      return `(${toRoman(num, true)})`;
    case 'romanUcParenR':
      return `${toRoman(num, true)})`;
    case 'romanUcPeriod':
      return `${toRoman(num, true)}.`;
  }
}

/**
 * Converts a number to alphabetic representation (a, b, c, ... z, aa, ab, ...).
 */
function toAlpha(num: number, uppercase: boolean): string {
  let result = '';
  while (num > 0) {
    num--;
    result = String.fromCharCode((num % 26) + (uppercase ? 65 : 97)) + result;
    num = Math.floor(num / 26);
  }
  return result || (uppercase ? 'A' : 'a');
}

/**
 * Converts a number to Roman numeral representation.
 */
function toRoman(num: number, uppercase: boolean): string {
  const romanNumerals = [
    ['M', 1000], ['CM', 900], ['D', 500], ['CD', 400],
    ['C', 100], ['XC', 90], ['L', 50], ['XL', 40],
    ['X', 10], ['IX', 9], ['V', 5], ['IV', 4], ['I', 1]
  ] as const;

  let result = '';
  for (const [letter, value] of romanNumerals) {
    while (num >= value) {
      result += letter;
      num -= value;
    }
  }
  return uppercase ? result : result.toLowerCase();
}

/**
 * Renders a text run to HTML.
 */
function renderTextRun(run: TextRun, autofitContext: AutofitContext): HTMLElement {
  const span = document.createElement('span');

  // Set text content
  span.textContent = run.text;

  // Apply font family
  if (run.fontFamily) {
    span.style.fontFamily = `"${run.fontFamily}", sans-serif`;
  }

  // Apply font size with autofit scaling
  if (run.fontSize) {
    const scaledSize = run.fontSize * autofitContext.fontScale;
    span.style.fontSize = `${scaledSize}px`;
  }

  // Apply color
  if (run.color) {
    span.style.color = colorToCss(run.color);
  }

  // Apply bold
  if (run.bold) {
    span.style.fontWeight = 'bold';
  }

  // Apply italic
  if (run.italic) {
    span.style.fontStyle = 'italic';
  }

  // Apply underline
  if (run.underline) {
    span.style.textDecoration = 'underline';
  }

  // Apply strikethrough
  if (run.strikethrough) {
    span.style.textDecoration = span.style.textDecoration
      ? `${span.style.textDecoration} line-through`
      : 'line-through';
  }

  // Apply baseline (subscript/superscript)
  if (run.baseline) {
    if (run.baseline > 0) {
      // Superscript: positive baseline
      span.style.verticalAlign = 'super';
      span.style.fontSize = '0.7em'; // Make it smaller
    } else {
      // Subscript: negative baseline
      span.style.verticalAlign = 'sub';
      span.style.fontSize = '0.7em'; // Make it smaller
    }
  }

  // Apply character spacing
  if (run.characterSpacing) {
    span.style.letterSpacing = `${run.characterSpacing}px`;
  }

  // Apply text capitalization
  if (run.capitalization === 'allCaps') {
    span.style.textTransform = 'uppercase';
  } else if (run.capitalization === 'smallCaps') {
    span.style.fontVariant = 'small-caps';
  }

  // Apply highlight/background color
  if (run.highlight) {
    span.style.backgroundColor = colorToCss(run.highlight);
  }

  // Apply hyperlink
  if (run.link) {
    const link = document.createElement('a');
    link.href = run.link;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    link.style.color = 'inherit';
    link.style.textDecoration = 'underline';
    link.appendChild(span);
    return link;
  }

  return span;
}

/**
 * Renders text body to an SVG foreignObject.
 * Useful when embedding text within SVG shapes.
 *
 * @param text - Parsed text body
 * @param width - Available width
 * @param height - Available height
 * @returns SVG foreignObject element
 */
export function renderTextBodyToSvg(
  text: TextBody,
  width: number,
  height: number
): SVGForeignObjectElement {
  const foreignObject = document.createElementNS('http://www.w3.org/2000/svg', 'foreignObject');
  foreignObject.setAttribute('width', String(width));
  foreignObject.setAttribute('height', String(height));

  const div = document.createElement('div');
  div.style.width = '100%';
  div.style.height = '100%';
  div.style.boxSizing = 'border-box';
  div.style.border = 'none';
  div.style.outline = 'none';
  div.style.background = 'transparent';

  renderTextBody(text, div);

  foreignObject.appendChild(div);
  return foreignObject;
}
