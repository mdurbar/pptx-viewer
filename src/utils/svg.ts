/**
 * SVG-related constants and utilities.
 */

/** SVG namespace URI */
export const SVG_NS = 'http://www.w3.org/2000/svg';

/** XLink namespace URI (for href attributes in older SVG) */
export const XLINK_NS = 'http://www.w3.org/1999/xlink';

/**
 * Creates an SVG element with the correct namespace.
 *
 * @param tagName - The SVG element tag name (e.g., 'rect', 'circle', 'path')
 * @param attrs - Optional attributes to set on the element
 * @returns The created SVG element
 *
 * @example
 * ```typescript
 * const rect = createSvgElement('rect', { width: '100', height: '50', fill: 'red' });
 * const circle = createSvgElement('circle', { cx: '50', cy: '50', r: '25' });
 * ```
 */
export function createSvgElement<K extends keyof SVGElementTagNameMap>(
  tagName: K,
  attrs?: Record<string, string>
): SVGElementTagNameMap[K] {
  const element = document.createElementNS(SVG_NS, tagName);
  if (attrs) {
    setAttributes(element, attrs);
  }
  return element;
}

/**
 * Sets multiple attributes on an SVG element.
 *
 * @param element - The SVG element to modify
 * @param attrs - Object mapping attribute names to values
 *
 * @example
 * ```typescript
 * const rect = createSvgElement('rect');
 * setAttributes(rect, { x: '10', y: '20', width: '100', height: '50' });
 * ```
 */
export function setAttributes(element: SVGElement, attrs: Record<string, string>): void {
  for (const [key, value] of Object.entries(attrs)) {
    element.setAttribute(key, value);
  }
}

/** Counter for generating unique IDs */
let idCounter = 0;

/**
 * Generates a unique ID for SVG elements (defs, gradients, patterns, filters, etc.).
 *
 * Uses an incrementing counter to ensure uniqueness within the session.
 *
 * @param prefix - Prefix for the ID (e.g., 'gradient', 'pattern', 'filter')
 * @returns Unique ID string
 *
 * @example
 * ```typescript
 * const gradientId = generateId('gradient'); // "gradient_1"
 * const patternId = generateId('pattern');   // "pattern_2"
 * ```
 */
export function generateId(prefix: string): string {
  return `${prefix}_${++idCounter}`;
}

/**
 * Resets the ID counter. Useful for testing.
 */
export function resetIdCounter(): void {
  idCounter = 0;
}
