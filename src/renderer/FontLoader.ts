/**
 * Font loading utilities for injecting embedded fonts into the document.
 *
 * Uses @font-face CSS rules to make embedded fonts available for rendering.
 */

import type { EmbeddedFont } from '../core/types';
import { sanitizeFontFamily } from '../utils/css';

/** ID of the style element used for font injection */
const FONT_STYLE_ID = 'pptx-embedded-fonts';

/**
 * Injects @font-face rules for embedded fonts into the document.
 *
 * Creates a <style> element with @font-face declarations for each font,
 * making them available for use in CSS font-family properties.
 *
 * @param fonts - Map of embedded fonts to inject
 *
 * @example
 * ```typescript
 * const presentation = await parsePPTX(archive);
 * if (presentation.fonts.size > 0) {
 *   injectFontStyles(presentation.fonts);
 * }
 * ```
 */
export function injectFontStyles(fonts: Map<string, EmbeddedFont>): void {
  if (fonts.size === 0) return;

  // Remove existing style element if present
  cleanupFontStyles();

  // Create new style element
  const styleEl = document.createElement('style');
  styleEl.id = FONT_STYLE_ID;

  // Generate @font-face rules. Font names are sanitized with the same
  // shared policy used at every font-family reference site, so the
  // @font-face family always matches what renderers emit.
  const rules = Array.from(fonts.values())
    .map(
      (font) => `
@font-face {
  font-family: "${sanitizeFontFamily(font.name)}";
  src: url("${font.url}") format("${font.format}");
  font-display: swap;
}`
    )
    .join('\n');

  styleEl.textContent = rules;
  document.head.appendChild(styleEl);
}

/**
 * Removes the injected font styles from the document.
 *
 * Should be called when the presentation is unloaded to clean up resources.
 */
export function cleanupFontStyles(): void {
  const styleEl = document.getElementById(FONT_STYLE_ID);
  if (styleEl) {
    styleEl.remove();
  }
}

