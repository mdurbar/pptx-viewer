/**
 * Shared background parsing utilities.
 *
 * Used by SlideParser, LayoutParser, and MasterParser for parsing
 * background fills from bgPr elements.
 */

import type { Fill, Background, ThemeColors } from '../core/types';
import type { ShapeParseContext } from './ShapeParser';
import { parseColorElement } from './TextParser';
import { findChildByName, findFirstByName } from '../utils/xml';
import { getMimeType } from '../core/unzip';

/**
 * Parses background fill from a bgPr element.
 *
 * @param bgPr - The bgPr (background properties) XML element
 * @param context - Shape parsing context with colors, relationships, and archive
 * @returns Parsed Fill or undefined if no fill could be parsed
 */
export function parseBackgroundFill(bgPr: Element, context: ShapeParseContext): Fill | undefined {
  // Check for solid fill
  const solidFill = findChildByName(bgPr, 'solidFill');
  if (solidFill) {
    const color = parseColorElement(solidFill, context.themeColors);
    if (color) {
      return { type: 'solid', color };
    }
  }

  // Check for gradient fill
  const gradFill = findChildByName(bgPr, 'gradFill');
  if (gradFill) {
    // Simplified gradient parsing for backgrounds
    const color = parseColorElement(gradFill, context.themeColors);
    if (color) {
      return { type: 'solid', color };
    }
  }

  // Check for image fill
  const blipFill = findChildByName(bgPr, 'blipFill');
  if (blipFill) {
    const blip = findChildByName(blipFill, 'blip');
    if (blip) {
      const rEmbed = blip.getAttributeNS(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'embed'
      ) || blip.getAttribute('r:embed');

      if (rEmbed) {
        const imagePath = context.relationships.resolvePath(rEmbed, context.basePath);
        if (imagePath) {
          const mimeType = getMimeType(imagePath);
          const src = context.archive.getBlobUrl(imagePath, mimeType);
          if (src) {
            return {
              type: 'image',
              src,
              mode: 'cover',
            };
          }
        }
      }
    }
  }

  return undefined;
}

/**
 * Parses a background element from a cSld parent.
 *
 * @param root - The root element containing cSld
 * @param context - Shape parsing context
 * @returns Parsed Background or undefined
 */
export function parseBackground(root: Element, context: ShapeParseContext): Background | undefined {
  const cSld = findFirstByName(root, 'cSld');
  if (!cSld) return undefined;

  const bg = findChildByName(cSld, 'bg');
  if (!bg) return undefined;

  // Try bgPr (background properties)
  const bgPr = findChildByName(bg, 'bgPr');
  if (bgPr) {
    const fill = parseBackgroundFill(bgPr, context);
    if (fill) {
      return { fill };
    }
  }

  // Try bgRef (background reference to theme)
  const bgRef = findChildByName(bg, 'bgRef');
  if (bgRef) {
    const color = parseColorElement(bgRef, context.themeColors);
    if (color) {
      return {
        fill: { type: 'solid', color },
      };
    }
  }

  return undefined;
}
