/**
 * Renderer for text content.
 *
 * Converts parsed TextBody into HTML elements with proper styling.
 */

import type { TextBody, Paragraph, TextRun, Color, BulletStyle, TextAutofit, TextGlow, TextReflection } from '../core/types';
import { colorToCss } from '../utils/color';
import { SVG_NS } from '../utils/svg';
import { sanitizeFontFamily, isSafeLinkUrl } from '../utils/css';

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

  // Get autofit context. Values come from untrusted content: PowerPoint
  // only ever shrinks (fontScale ≤ 1) and caps line-spacing reduction at
  // 20%, so clamp to a sane range — an unclamped reduction of 1 would
  // collapse text to zero line height (hidden-disclaimer spoofing).
  const autofitContext: AutofitContext = {
    fontScale: clamp(text.autofit?.fontScale ?? 1, 0.05, 1),
    lineSpacingReduction: clamp(text.autofit?.lineSpacingReduction ?? 0, 0, 0.4),
  };

  // Render paragraphs
  for (const paragraph of text.paragraphs) {
    const pElement = renderParagraph(paragraph, numberingState, autofitContext);
    container.appendChild(pElement);
  }
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/**
 * Single line spacing as a fraction of the font size. OOXML `spcPct`
 * values are multiples of single spacing, which from typical font metrics
 * is ~1.2x the em size — mapping 100% to CSS line-height 1.0 renders
 * every explicitly-spaced paragraph too tight.
 */
export const SINGLE_LINE_SPACING = 1.2;

/** Browser default font size, the fallback when no run declares one. */
const DEFAULT_FONT_SIZE_PX = 16;

/**
 * Floors for line height so untrusted spacing values can't collapse text
 * to (near) zero. Unitless for `multiple` spacing, pixels for `exact`.
 */
export const MIN_LINE_HEIGHT = 0.1;
export const MIN_LINE_HEIGHT_PX = 4;

/**
 * The font size percent-based paragraph spacing is measured against:
 * the largest run size in the paragraph (falling back to the browser
 * default when no run declares a size).
 */
function effectiveFontSize(paragraph: Paragraph, autofitContext: AutofitContext): number {
  let size = 0;
  for (const run of paragraph.runs) {
    if (run.fontSize && run.fontSize > size) {
      size = run.fontSize;
    }
  }
  return (size || DEFAULT_FONT_SIZE_PX) * autofitContext.fontScale;
}

/**
 * Resolves space-before/space-after to pixels. Values come from
 * untrusted content, so the result is clamped to non-negative — a
 * negative margin would pull paragraphs over each other.
 */
function resolveParagraphSpacing(
  spacing: NonNullable<Paragraph['spaceBefore']>,
  paragraph: Paragraph,
  autofitContext: AutofitContext
): number {
  const px =
    spacing.type === 'exact'
      ? spacing.px
      : spacing.value * effectiveFontSize(paragraph, autofitContext);
  return Math.max(px, 0);
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

  // Paragraphs are normal block flow: alignment via text-align (a flex
  // row can't center text whose wrapper fills the row), hanging indents
  // via text-indent.
  if (paragraph.align && paragraph.align !== 'left') {
    p.style.textAlign = paragraph.align;
  }

  // Apply line spacing with autofit reduction. Values come from untrusted
  // content, so floor the result above zero — an explicit spcPct/spcPts of
  // 0 would otherwise collapse text to zero line height (hidden-text
  // spoofing), the same vector the autofit clamp guards against.
  const reduction = 1 - autofitContext.lineSpacingReduction;
  const spacing = paragraph.lineSpacing;
  if (!spacing) {
    if (autofitContext.lineSpacingReduction > 0) {
      p.style.lineHeight = String(SINGLE_LINE_SPACING * reduction);
    }
  } else if (spacing.type === 'multiple') {
    p.style.lineHeight = String(Math.max(spacing.value * SINGLE_LINE_SPACING * reduction, MIN_LINE_HEIGHT));
  } else {
    p.style.lineHeight = `${Math.max(spacing.px * reduction, MIN_LINE_HEIGHT_PX)}px`;
  }

  // Apply space before/after (percent variants are relative to text size)
  if (paragraph.spaceBefore) {
    p.style.marginTop = `${resolveParagraphSpacing(paragraph.spaceBefore, paragraph, autofitContext)}px`;
  }
  if (paragraph.spaceAfter) {
    p.style.marginBottom = `${resolveParagraphSpacing(paragraph.spaceAfter, paragraph, autofitContext)}px`;
  }

  // Indentation: marginLeft is where wrapped text sits; indent shifts the
  // first line relative to it (negative = hanging, the bullet case).
  const level = paragraph.level || 0;
  const leftMargin = paragraph.marginLeft ?? (level * 36); // Default 36px per level
  const hangingIndent = paragraph.indent ?? (paragraph.bullet ? -18 : 0); // Default hanging indent for bullets

  if (leftMargin > 0) {
    p.style.marginLeft = `${leftMargin}px`;
  }
  if (hangingIndent !== 0) {
    p.style.textIndent = `${hangingIndent}px`;
  }

  // Add bullet point
  if (paragraph.bullet) {
    const bulletSpan = document.createElement('span');
    // inline-block: gives the bullet a fixed width (so following text
    // starts at the paragraph's left margin) and resets text-indent
    bulletSpan.style.display = 'inline-block';

    const bulletWidth = Math.abs(hangingIndent) || 18;
    bulletSpan.style.width = `${bulletWidth}px`;
    bulletSpan.style.textAlign = 'left';

    // Apply bullet styling
    if (paragraph.bullet.font) {
      bulletSpan.style.fontFamily = `"${sanitizeFontFamily(paragraph.bullet.font)}", sans-serif`;
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

  // Render text runs, honoring hard line breaks (<a:br/>)
  for (const run of paragraph.runs) {
    if (run.breakBefore) {
      p.appendChild(document.createElement('br'));
    }
    p.appendChild(renderTextRun(run, autofitContext));
  }

  // Empty paragraph - add a non-breaking space to maintain height
  if (paragraph.runs.length === 0 && !paragraph.bullet) {
    const filler = document.createElement('span');
    filler.textContent = ' ';
    p.appendChild(filler);
  }

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
    span.style.fontFamily = `"${sanitizeFontFamily(run.fontFamily)}", sans-serif`;
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

  // Apply glow effect
  if (run.glow) {
    const glowShadow = createGlowShadow(run.glow);
    span.style.textShadow = glowShadow;
  }

  // Apply text outline/stroke
  if (run.outline) {
    span.style.webkitTextStroke = `${run.outline.width}px ${colorToCss(run.outline.color)}`;
    // Paint order ensures fill is on top of stroke
    span.style.paintOrder = 'stroke fill';
  }

  // Apply hyperlink. Targets come from untrusted relationship entries —
  // only navigation schemes may reach href (a javascript: URL would
  // execute in the host page on click).
  if (run.link && isSafeLinkUrl(run.link)) {
    const link = document.createElement('a');
    link.href = run.link;
    link.target = '_blank';
    link.rel = 'noopener noreferrer';
    link.style.color = 'inherit';
    link.style.textDecoration = 'underline';
    link.appendChild(span);

    // Handle reflection with hyperlink
    if (run.reflection) {
      return wrapWithReflection(link, run.reflection);
    }
    return link;
  }

  // Apply reflection effect
  if (run.reflection) {
    return wrapWithReflection(span, run.reflection);
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
  const foreignObject = document.createElementNS(SVG_NS, 'foreignObject');
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

/**
 * Creates a CSS text-shadow value for a glow effect.
 * Uses multiple layered shadows to create a smooth glow.
 */
function createGlowShadow(glow: TextGlow): string {
  const color = colorToCss(glow.color);
  const radius = glow.radius;

  // Create multiple shadow layers for a smoother glow
  const shadows: string[] = [];
  const layers = 3;

  for (let i = 1; i <= layers; i++) {
    const layerRadius = (radius / layers) * i;
    shadows.push(`0 0 ${layerRadius}px ${color}`);
  }

  return shadows.join(', ');
}

/**
 * Wraps an element with a reflection effect.
 * Creates a container with the original element and a reflected copy.
 */
function wrapWithReflection(element: HTMLElement, reflection: TextReflection): HTMLElement {
  const container = document.createElement('span');
  container.style.display = 'inline-flex';
  container.style.flexDirection = 'column';
  container.style.alignItems = 'flex-start';

  // Clone the element for the reflection
  const reflectionEl = element.cloneNode(true) as HTMLElement;

  // Apply reflection styles
  reflectionEl.style.transform = `scaleY(-${reflection.scaleY / 100}) skewX(${reflection.skewX}deg)`;
  reflectionEl.style.transformOrigin = 'center top';
  reflectionEl.style.marginTop = `${reflection.distance}px`;

  // Create gradient mask for fade effect
  const startOpacity = reflection.startOpacity;
  const endOpacity = reflection.endOpacity;

  // Use a mask with linear gradient for the fade
  reflectionEl.style.maskImage = `linear-gradient(to bottom, rgba(0,0,0,${startOpacity}), rgba(0,0,0,${endOpacity}))`;
  reflectionEl.style.webkitMaskImage = `linear-gradient(to bottom, rgba(0,0,0,${startOpacity}), rgba(0,0,0,${endOpacity}))`;

  // Apply blur if specified
  if (reflection.blurRadius > 0) {
    reflectionEl.style.filter = `blur(${reflection.blurRadius}px)`;
  }

  // Prevent reflection from being interactive
  reflectionEl.style.pointerEvents = 'none';
  reflectionEl.setAttribute('aria-hidden', 'true');

  // Remove any links in the reflection to avoid duplicate navigation,
  // keeping their (already-safe) child nodes without re-parsing HTML.
  const links = reflectionEl.querySelectorAll('a');
  links.forEach(link => {
    link.replaceWith(...link.childNodes);
  });

  container.appendChild(element);
  container.appendChild(reflectionEl);

  return container;
}
