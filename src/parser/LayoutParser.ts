/**
 * Parser for slide layout files.
 *
 * Slide layouts define placeholder positions and can override master properties.
 * Layouts inherit from slide masters and define the structure for specific slide types.
 *
 * Each layout is stored in ppt/slideLayouts/slideLayoutN.xml.
 */

import type { SlideLayout, Background, ThemeColors, ColorMap, SlideElement, Fill } from '../core/types';
import type { PPTXArchive } from '../core/unzip';
import type { RelationshipMap } from './RelationshipParser';
import { parseRelationships, RELATIONSHIP_TYPES } from './RelationshipParser';
import { parseShapeTree, type ShapeParseContext } from './ShapeParser';
import { parseColorElement } from './TextParser';
import { parseXml, findFirstByName, findChildByName, getAttribute } from '../utils/xml';
import { getMimeType } from '../core/unzip';
import { XMLParseError } from '../core/errors';

/**
 * Gets the path to a slide layout's relationships file.
 */
export function getSlideLayoutRelsPath(layoutPath: string): string {
  // Convert ppt/slideLayouts/slideLayout1.xml to ppt/slideLayouts/_rels/slideLayout1.xml.rels
  const parts = layoutPath.split('/');
  const filename = parts.pop()!;
  return [...parts, '_rels', `${filename}.rels`].join('/');
}

/**
 * Parses a slide layout XML file.
 *
 * @param xmlContent - Raw XML content of the layout file
 * @param layoutId - Relationship ID of this layout
 * @param archive - PPTX archive for accessing images
 * @param themeColors - Theme colors for color resolution
 * @param layoutPath - Path to the layout file (for relationship resolution)
 * @returns Parsed slide layout object
 */
export function parseSlideLayout(
  xmlContent: string,
  layoutId: string,
  archive: PPTXArchive,
  themeColors: ThemeColors,
  layoutPath: string
): SlideLayout {
  let doc;
  try {
    doc = parseXml(xmlContent);
  } catch (error) {
    throw new XMLParseError(
      error instanceof Error ? error.message : 'Unknown error',
      layoutPath
    );
  }

  const root = doc.documentElement;

  // Load layout relationships
  const relsPath = getSlideLayoutRelsPath(layoutPath);
  const relsXml = archive.getText(relsPath);
  let relationships: RelationshipMap;

  try {
    relationships = relsXml
      ? parseRelationships(relsXml)
      : createEmptyRelationshipMap();
  } catch (error) {
    console.warn(`Failed to parse relationships for layout ${layoutId}:`, error);
    relationships = createEmptyRelationshipMap();
  }

  // Create parsing context
  const context: ShapeParseContext = {
    themeColors,
    relationships,
    archive,
    basePath: layoutPath,
  };

  // Get the master ID from relationships
  const masterId = getMasterIdFromRels(relationships);

  // Check if master shapes should be shown
  const showMasterShapes = getAttribute(root, 'showMasterSp') !== '0';

  // Get the layout type from the type attribute
  const layoutType = getAttribute(root, 'type') || undefined;

  // Parse color map (if present, overrides master)
  const colorMap = parseColorMapOverride(root);

  // Parse background (optional override of master)
  let background: Background | undefined;
  try {
    background = parseLayoutBackground(root, context);
  } catch (error) {
    console.warn(`Failed to parse background for layout ${layoutId}:`, error);
  }

  // Find the shape tree
  const cSld = findFirstByName(root, 'cSld');
  const spTree = cSld ? findFirstByName(cSld, 'spTree') : null;

  // Parse elements
  let elements: SlideElement[] = [];
  try {
    elements = spTree ? parseShapeTree(spTree, context) : [];
  } catch (error) {
    console.warn(`Failed to parse shapes for layout ${layoutId}:`, error);
  }

  // Get the layout name
  const name = cSld ? getAttribute(cSld, 'name') : undefined;

  return {
    id: layoutId,
    name: name || undefined,
    type: layoutType,
    masterId: masterId || '',
    background,
    elements,
    showMasterShapes,
    colorMap,
  };
}

/**
 * Gets the master relationship ID from layout relationships.
 */
function getMasterIdFromRels(relationships: RelationshipMap): string | null {
  const masterRels = relationships.getByType(RELATIONSHIP_TYPES.SLIDE_MASTER);
  if (masterRels.length > 0) {
    return masterRels[0].id;
  }
  return null;
}

/**
 * Parses color map override from a layout.
 * Returns undefined if no override is specified.
 */
function parseColorMapOverride(root: Element): ColorMap | undefined {
  const clrMapOvr = findFirstByName(root, 'clrMapOvr');
  if (!clrMapOvr) return undefined;

  // Check if it's a master color map reference (no override)
  const masterClrMapping = findChildByName(clrMapOvr, 'masterClrMapping');
  if (masterClrMapping) return undefined;

  // Check for override
  const overrideClrMapping = findChildByName(clrMapOvr, 'overrideClrMapping');
  if (!overrideClrMapping) return undefined;

  return {
    bg1: getAttribute(overrideClrMapping, 'bg1') || undefined,
    bg2: getAttribute(overrideClrMapping, 'bg2') || undefined,
    tx1: getAttribute(overrideClrMapping, 'tx1') || undefined,
    tx2: getAttribute(overrideClrMapping, 'tx2') || undefined,
    accent1: getAttribute(overrideClrMapping, 'accent1') || undefined,
    accent2: getAttribute(overrideClrMapping, 'accent2') || undefined,
    accent3: getAttribute(overrideClrMapping, 'accent3') || undefined,
    accent4: getAttribute(overrideClrMapping, 'accent4') || undefined,
    accent5: getAttribute(overrideClrMapping, 'accent5') || undefined,
    accent6: getAttribute(overrideClrMapping, 'accent6') || undefined,
    hlink: getAttribute(overrideClrMapping, 'hlink') || undefined,
    folHlink: getAttribute(overrideClrMapping, 'folHlink') || undefined,
  };
}

/**
 * Parses the layout background.
 */
function parseLayoutBackground(root: Element, context: ShapeParseContext): Background | undefined {
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
