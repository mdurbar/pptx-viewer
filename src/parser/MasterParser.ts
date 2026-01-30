/**
 * Parser for slide master files.
 *
 * Slide masters define the base styling, background, and placeholder
 * positions for all slides that use them.
 *
 * Each master is stored in ppt/slideMasters/slideMasterN.xml.
 */

import type { SlideMaster, Background, ThemeColors, ColorMap, SlideElement } from '../core/types';
import type { PPTXArchive } from '../core/unzip';
import type { RelationshipMap } from './RelationshipParser';
import { parseRelationships, createEmptyRelationshipMap, getRelationshipsPath, RELATIONSHIP_TYPES } from './RelationshipParser';
import { parseShapeTree, type ShapeParseContext } from './ShapeParser';
import { parseBackground } from './BackgroundParser';
import { parseXml, findFirstByName, findChildByName, findChildrenByName, getAttribute } from '../utils/xml';
import { XMLParseError } from '../core/errors';


/**
 * Parses a slide master XML file.
 *
 * @param xmlContent - Raw XML content of the master file
 * @param masterId - Relationship ID of this master
 * @param archive - PPTX archive for accessing images
 * @param themeColors - Theme colors for color resolution
 * @param masterPath - Path to the master file (for relationship resolution)
 * @returns Parsed slide master object
 */
export function parseSlideMaster(
  xmlContent: string,
  masterId: string,
  archive: PPTXArchive,
  themeColors: ThemeColors,
  masterPath: string
): SlideMaster {
  let doc;
  try {
    doc = parseXml(xmlContent);
  } catch (error) {
    throw new XMLParseError(
      error instanceof Error ? error.message : 'Unknown error',
      masterPath
    );
  }

  const root = doc.documentElement;

  // Load master relationships
  const relsPath = getRelationshipsPath(masterPath);
  const relsXml = archive.getText(relsPath);
  let relationships: RelationshipMap;

  try {
    relationships = relsXml
      ? parseRelationships(relsXml)
      : createEmptyRelationshipMap();
  } catch (error) {
    console.warn(`Failed to parse relationships for master ${masterId}:`, error);
    relationships = createEmptyRelationshipMap();
  }

  // Create parsing context
  const context: ShapeParseContext = {
    themeColors,
    relationships,
    archive,
    basePath: masterPath,
  };

  // Parse color map
  const colorMap = parseColorMap(root);

  // Parse background
  let background: Background | undefined;
  try {
    background = parseBackground(root, context);
  } catch (error) {
    console.warn(`Failed to parse background for master ${masterId}:`, error);
  }

  // Find the shape tree
  const cSld = findFirstByName(root, 'cSld');
  const spTree = cSld ? findFirstByName(cSld, 'spTree') : null;

  // Parse elements
  let elements: SlideElement[] = [];
  try {
    elements = spTree ? parseShapeTree(spTree, context) : [];
  } catch (error) {
    console.warn(`Failed to parse shapes for master ${masterId}:`, error);
  }

  // Get the master name
  const name = cSld ? getAttribute(cSld, 'name') : undefined;

  // Extract layout relationship IDs
  const layoutIds = extractLayoutIds(root, relationships);

  return {
    id: masterId,
    name: name || undefined,
    background,
    elements,
    colorMap,
    layoutIds,
  };
}

/**
 * Parses the color map from a slide master.
 * The color map maps scheme colors (like bg1, tx1) to theme colors (like lt1, dk1).
 */
function parseColorMap(root: Element): ColorMap {
  const clrMap = findFirstByName(root, 'clrMap');
  if (!clrMap) {
    // Default color mapping
    return {
      bg1: 'lt1',
      bg2: 'lt2',
      tx1: 'dk1',
      tx2: 'dk2',
      accent1: 'accent1',
      accent2: 'accent2',
      accent3: 'accent3',
      accent4: 'accent4',
      accent5: 'accent5',
      accent6: 'accent6',
      hlink: 'hlink',
      folHlink: 'folHlink',
    };
  }

  return {
    bg1: getAttribute(clrMap, 'bg1') || 'lt1',
    bg2: getAttribute(clrMap, 'bg2') || 'lt2',
    tx1: getAttribute(clrMap, 'tx1') || 'dk1',
    tx2: getAttribute(clrMap, 'tx2') || 'dk2',
    accent1: getAttribute(clrMap, 'accent1') || 'accent1',
    accent2: getAttribute(clrMap, 'accent2') || 'accent2',
    accent3: getAttribute(clrMap, 'accent3') || 'accent3',
    accent4: getAttribute(clrMap, 'accent4') || 'accent4',
    accent5: getAttribute(clrMap, 'accent5') || 'accent5',
    accent6: getAttribute(clrMap, 'accent6') || 'accent6',
    hlink: getAttribute(clrMap, 'hlink') || 'hlink',
    folHlink: getAttribute(clrMap, 'folHlink') || 'folHlink',
  };
}

/**
 * Extracts layout relationship IDs from the master.
 */
function extractLayoutIds(root: Element, relationships: RelationshipMap): string[] {
  const layoutIds: string[] = [];

  // Find sldLayoutIdLst element
  const sldLayoutIdLst = findFirstByName(root, 'sldLayoutIdLst');
  if (sldLayoutIdLst) {
    const sldLayoutIds = findChildrenByName(sldLayoutIdLst, 'sldLayoutId');
    for (const sldLayoutId of sldLayoutIds) {
      const rId = getAttribute(sldLayoutId, 'r:id');
      if (rId) {
        layoutIds.push(rId);
      }
    }
  }

  // If no explicit list, try to get layouts from relationships
  if (layoutIds.length === 0) {
    const layoutRels = relationships.getByType(RELATIONSHIP_TYPES.SLIDE_LAYOUT);
    for (const rel of layoutRels) {
      layoutIds.push(rel.id);
    }
  }

  return layoutIds;
}

