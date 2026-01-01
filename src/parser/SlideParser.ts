/**
 * Parser for individual slide files.
 *
 * Each slide is stored in ppt/slides/slideN.xml.
 * Slides contain a shape tree (spTree) with all visual elements.
 */

import type { Slide, Background, ThemeColors, Fill } from '../core/types';
import type { PPTXArchive } from '../core/unzip';
import type { RelationshipMap } from './RelationshipParser';
import { parseRelationships } from './RelationshipParser';
import { parseShapeTree, type ShapeParseContext } from './ShapeParser';
import { parseColorElement } from './TextParser';
import { parseXml, findFirstByName, findChildByName } from '../utils/xml';
import { getSlideRelsPath, getMimeType } from '../core/unzip';

/**
 * Parses a slide XML file.
 *
 * @param xmlContent - Raw XML content of the slide file
 * @param slideIndex - 0-based slide index
 * @param archive - PPTX archive for accessing images
 * @param themeColors - Theme colors for color resolution
 * @param slidePath - Path to the slide file (for relationship resolution)
 * @returns Parsed slide object
 */
export function parseSlide(
  xmlContent: string,
  slideIndex: number,
  archive: PPTXArchive,
  themeColors: ThemeColors,
  slidePath: string
): Slide {
  const doc = parseXml(xmlContent);
  const root = doc.documentElement;

  // Load slide relationships
  const slideNumber = slideIndex + 1;
  const relsPath = getSlideRelsPath(slideNumber);
  const relsXml = archive.getText(relsPath);
  const relationships = relsXml
    ? parseRelationships(relsXml)
    : createEmptyRelationshipMap();

  // Create parsing context
  const context: ShapeParseContext = {
    themeColors,
    relationships,
    archive,
    basePath: slidePath,
  };

  // Parse background
  const background = parseSlideBackground(root, context);

  // Find the shape tree
  const cSld = findFirstByName(root, 'cSld');
  const spTree = cSld ? findFirstByName(cSld, 'spTree') : null;

  // Parse elements
  const elements = spTree ? parseShapeTree(spTree, context) : [];

  return {
    index: slideIndex,
    background,
    elements,
  };
}

/**
 * Parses the slide background.
 */
function parseSlideBackground(root: Element, context: ShapeParseContext): Background | undefined {
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
    // For now, parse any embedded color
    const color = parseColorElement(bgRef, context.themeColors);
    if (color) {
      return {
        fill: { type: 'solid', color },
      };
    }
  }

  return undefined;
}

/**
 * Parses background fill from bgPr element.
 */
function parseBackgroundFill(bgPr: Element, context: ShapeParseContext): Fill | undefined {
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
 * Creates an empty relationship map when no .rels file exists.
 */
function createEmptyRelationshipMap(): RelationshipMap {
  return {
    byId: new Map(),
    byType: new Map(),
    get() {
      return undefined;
    },
    getByType() {
      return [];
    },
    resolvePath() {
      return null;
    },
  };
}
