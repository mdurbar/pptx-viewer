/**
 * Renderer for text content.
 *
 * Converts parsed TextBody into HTML elements with proper styling.
 */

import type { TextBody, Paragraph, TextRun, Color } from '../core/types';
import { colorToCss } from '../utils/color';

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

  // Render paragraphs
  for (const paragraph of text.paragraphs) {
    const pElement = renderParagraph(paragraph);
    container.appendChild(pElement);
  }
}

/**
 * Renders a paragraph to HTML.
 */
function renderParagraph(paragraph: Paragraph): HTMLElement {
  const p = document.createElement('p');
  p.style.margin = '0';
  p.style.padding = '0';

  // Apply alignment
  switch (paragraph.align) {
    case 'center':
      p.style.textAlign = 'center';
      break;
    case 'right':
      p.style.textAlign = 'right';
      break;
    case 'justify':
      p.style.textAlign = 'justify';
      break;
    default:
      p.style.textAlign = 'left';
  }

  // Apply line spacing
  if (paragraph.lineSpacing) {
    p.style.lineHeight = String(paragraph.lineSpacing);
  }

  // Apply spacing
  if (paragraph.spaceBefore) {
    p.style.marginTop = `${paragraph.spaceBefore}px`;
  }
  if (paragraph.spaceAfter) {
    p.style.marginBottom = `${paragraph.spaceAfter}px`;
  }

  // Apply indentation for lists
  if (paragraph.level && paragraph.level > 0) {
    p.style.marginLeft = `${paragraph.level * 24}px`;
  }

  // Add bullet point
  if (paragraph.bullet) {
    const bulletSpan = document.createElement('span');
    bulletSpan.style.marginRight = '8px';

    if (paragraph.bullet.type === 'bullet') {
      bulletSpan.textContent = paragraph.bullet.char || '•';
    } else {
      // Numbered - would need to track across paragraphs for proper numbering
      bulletSpan.textContent = `${paragraph.bullet.startAt || 1}.`;
    }

    p.appendChild(bulletSpan);
  }

  // Render text runs
  for (const run of paragraph.runs) {
    const runElement = renderTextRun(run);
    p.appendChild(runElement);
  }

  // Empty paragraph - add a line break to maintain height
  if (paragraph.runs.length === 0) {
    p.innerHTML = '&nbsp;';
  }

  return p;
}

/**
 * Renders a text run to HTML.
 */
function renderTextRun(run: TextRun): HTMLElement {
  const span = document.createElement('span');

  // Set text content
  span.textContent = run.text;

  // Apply font family
  if (run.fontFamily) {
    span.style.fontFamily = `"${run.fontFamily}", sans-serif`;
  }

  // Apply font size
  if (run.fontSize) {
    span.style.fontSize = `${run.fontSize}px`;
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
