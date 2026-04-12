/**
 * Parser for individual slide files.
 *
 * Each slide is stored in ppt/slides/slideN.xml.
 * Slides contain a shape tree (spTree) with all visual elements.
 */

import type { Slide, Background, ThemeColors, SlideElement, SlideLayout, SlideMaster } from '../core/types';
import type { PPTXArchive } from '../core/unzip';
import type { RelationshipMap } from './RelationshipParser';
import { parseRelationships, createEmptyRelationshipMap, RELATIONSHIP_TYPES } from './RelationshipParser';
import { parseShapeTree, type ShapeParseContext } from './ShapeParser';
import { parseBackground } from './BackgroundParser';
import { parseXml, findFirstByName } from '../utils/xml';
import { getSlideRelsPath } from '../core/unzip';
import { XMLParseError } from '../core/errors';

/**
 * Parses a slide XML file.
 *
 * @param xmlContent - Raw XML content of the slide file
 * @param slideIndex - 0-based slide index
 * @param archive - PPTX archive for accessing images
 * @param themeColors - Theme colors for color resolution
 * @param slidePath - Path to the slide file (for relationship resolution)
 * @param layout - Pre-resolved slide layout (enables placeholder inheritance)
 * @param master - Pre-resolved slide master (second-tier placeholder fallback)
 * @returns Parsed slide object
 */
export function parseSlide(
  xmlContent: string,
  slideIndex: number,
  archive: PPTXArchive,
  themeColors: ThemeColors,
  slidePath: string,
  layout?: SlideLayout | null,
  master?: SlideMaster | null
): Slide {
  let doc;
  try {
    doc = parseXml(xmlContent);
  } catch (error) {
    throw new XMLParseError(
      error instanceof Error ? error.message : 'Unknown error',
      slidePath
    );
  }

  const root = doc.documentElement;

  // Load slide relationships (non-fatal if missing)
  const slideNumber = slideIndex + 1;
  const relsPath = getSlideRelsPath(slideNumber);
  const relsXml = archive.getText(relsPath);
  let relationships: RelationshipMap;

  try {
    relationships = relsXml
      ? parseRelationships(relsXml)
      : createEmptyRelationshipMap();
  } catch (error) {
    console.warn(`Failed to parse relationships for slide ${slideNumber}:`, error);
    relationships = createEmptyRelationshipMap();
  }

  // Create parsing context
  const context: ShapeParseContext = {
    themeColors,
    relationships,
    archive,
    basePath: slidePath,
    layout,
    master,
  };

  // Parse background (non-fatal if it fails)
  let background: Background | undefined;
  try {
    background = parseBackground(root, context);
  } catch (error) {
    console.warn(`Failed to parse background for slide ${slideNumber}:`, error);
  }

  // Find the shape tree
  const cSld = findFirstByName(root, 'cSld');
  const spTree = cSld ? findFirstByName(cSld, 'spTree') : null;

  // Parse elements (with error recovery for individual shapes)
  let elements: SlideElement[] = [];
  try {
    elements = spTree ? parseShapeTree(spTree, context) : [];
  } catch (error) {
    console.warn(`Failed to parse shapes for slide ${slideNumber}:`, error);
  }

  // Get the layout relationship ID for this slide
  const layoutId = getSlideLayoutId(relationships);

  return {
    index: slideIndex,
    background,
    elements,
    layoutId,
  };
}

/**
 * Gets the layout relationship ID for a slide.
 */
function getSlideLayoutId(relationships: RelationshipMap): string | undefined {
  const layoutRels = relationships.getByType(RELATIONSHIP_TYPES.SLIDE_LAYOUT);
  if (layoutRels.length > 0) {
    return layoutRels[0].id;
  }
  return undefined;
}

