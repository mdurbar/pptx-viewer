/**
 * Parser for SmartArt diagrams.
 *
 * SmartArt diagrams in PPTX contain pre-computed DrawingML shapes in the
 * drawing*.xml file. This parser extracts those shapes using the existing
 * shape parser infrastructure.
 */

import type { SlideElement, ThemeColors } from '../core/types';
import type { PPTXArchive } from '../core/unzip';
import type { RelationshipMap } from './RelationshipParser';
import { parseRelationships } from './RelationshipParser';
import { parseShapeTree, type ShapeParseContext } from './ShapeParser';
import { parseXml, findFirstByName } from '../utils/xml';

/**
 * Result of parsing a diagram.
 */
export interface DiagramParseResult {
  /** Child elements from the diagram */
  children: SlideElement[];
  /** Diagram type/name for debugging */
  diagramType?: string;
}

/**
 * Gets the relationships path for a diagram drawing file.
 */
function getDiagramRelsPath(drawingPath: string): string {
  const parts = drawingPath.split('/');
  const filename = parts.pop() || '';
  const dir = parts.join('/');
  return `${dir}/_rels/${filename}.rels`;
}

/**
 * Parses a SmartArt diagram from its drawing XML file.
 *
 * The drawing XML contains pre-computed DrawingML shapes that PowerPoint
 * has already laid out. We can reuse our existing shape parser.
 *
 * @param drawingPath - Path to the diagram drawing XML file
 * @param archive - PPTX archive
 * @param themeColors - Theme colors for resolving color references
 * @param slideRelationships - Relationships from the slide (for resolving images)
 * @param basePath - Base path for resolving relationships
 * @returns Parsed diagram data or null if parsing fails
 */
export function parseDiagram(
  drawingPath: string,
  archive: PPTXArchive,
  themeColors: ThemeColors,
  slideRelationships: RelationshipMap,
  basePath: string
): DiagramParseResult | null {
  const drawingXml = archive.getText(drawingPath);
  if (!drawingXml) return null;

  try {
    const doc = parseXml(drawingXml);
    const root = doc.documentElement;

    // Find the shape tree in the diagram drawing
    // Structure: <dgm:drawing><dgm:spTree>...</dgm:spTree></dgm:drawing>
    // or: <dsp:drawing><dsp:spTree>...</dsp:spTree></dsp:drawing>
    const spTree = findFirstByName(root, 'spTree');
    if (!spTree) return null;

    // Try to load diagram-specific relationships
    const diagramRelsPath = getDiagramRelsPath(drawingPath);
    const diagramRelsXml = archive.getText(diagramRelsPath);
    const diagramRelationships = diagramRelsXml
      ? parseRelationships(diagramRelsXml)
      : slideRelationships;

    // Get the directory of the drawing file for resolving relative paths
    const drawingDir = drawingPath.substring(0, drawingPath.lastIndexOf('/'));

    // Create context for parsing shapes
    const context: ShapeParseContext = {
      themeColors,
      relationships: diagramRelationships,
      archive,
      basePath: drawingDir,
    };

    // Parse shapes using existing infrastructure
    const children = parseShapeTree(spTree, context);

    return {
      children,
      diagramType: getDiagramType(root),
    };
  } catch {
    return null;
  }
}

/**
 * Attempts to extract the diagram type from the root element.
 */
function getDiagramType(root: Element): string | undefined {
  // Try to get diagram name from attributes or nested elements
  const name = root.getAttribute('name');
  if (name) return name;

  return undefined;
}
