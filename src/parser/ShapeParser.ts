/**
 * Parser for shapes in OOXML.
 *
 * Shapes in PPTX are defined in DrawingML and include:
 * - sp: Basic shapes (rectangles, ellipses, etc.)
 * - pic: Pictures/images
 * - graphicFrame: Charts, tables, SmartArt (we use fallback images)
 * - grpSp: Groups of shapes
 * - cxnSp: Connector lines
 */

import type {
  SlideElement,
  ShapeElement,
  TextElement,
  ImageElement,
  GroupElement,
  TableElement,
  TableRow,
  TableCell,
  TableStyle,
  CellBorders,
  Bounds,
  ShapeType,
  Fill,
  Stroke,
  ArrowHead,
  Shadow,
  Color,
  GradientStop,
  ThemeColors,
  PlaceholderInfo,
  PlaceholderType,
} from '../core/types';
import type { RelationshipMap } from './RelationshipParser';
import type { PPTXArchive } from '../core/unzip';
import {
  findChildByName,
  findChildrenByName,
  findFirstByName,
  getAttribute,
  getNumberAttribute,
} from '../utils/xml';
import { emuToPixels, ooxmlAngleToDegrees } from '../utils/units';
import { parseTextBody, parseColorElement } from './TextParser';
import { getMimeType } from '../core/unzip';

/**
 * Context for parsing shapes.
 */
export interface ShapeParseContext {
  themeColors: ThemeColors;
  relationships: RelationshipMap;
  archive: PPTXArchive;
  basePath: string;
}

/**
 * Parses a shape tree (spTree) element.
 *
 * @param spTree - The shape tree XML element
 * @param context - Parsing context
 * @returns Array of parsed slide elements
 */
export function parseShapeTree(spTree: Element, context: ShapeParseContext): SlideElement[] {
  const elements: SlideElement[] = [];

  for (const child of Array.from(spTree.children)) {
    const localName = child.localName || child.nodeName.split(':').pop();

    try {
      switch (localName) {
        case 'sp': {
          const shape = parseShape(child, context);
          if (shape) elements.push(shape);
          break;
        }
        case 'pic': {
          const image = parsePicture(child, context);
          if (image) elements.push(image);
          break;
        }
        case 'grpSp': {
          const group = parseGroup(child, context);
          if (group) elements.push(group);
          break;
        }
        case 'graphicFrame': {
          // Charts, tables, SmartArt - try to use fallback image
          const fallback = parseGraphicFrame(child, context);
          if (fallback) elements.push(fallback);
          break;
        }
        case 'cxnSp': {
          // Connector shapes - parse as lines
          const connector = parseConnector(child, context);
          if (connector) elements.push(connector);
          break;
        }
      }
    } catch (error) {
      // Log but don't fail on individual shape parse errors
      console.warn(`Failed to parse ${localName} element:`, error);
    }
  }

  return elements;
}

/**
 * Parses a basic shape (sp) element.
 */
function parseShape(sp: Element, context: ShapeParseContext): ShapeElement | TextElement | null {
  // Get non-visual properties for ID
  const nvSpPr = findChildByName(sp, 'nvSpPr');
  const cNvPr = nvSpPr ? findChildByName(nvSpPr, 'cNvPr') : null;
  const id = (cNvPr ? getAttribute(cNvPr, 'id') : null) || generateId();

  // Extract placeholder info (for slide master/layout inheritance)
  const placeholder = parsePlaceholderInfo(nvSpPr);

  // Get shape properties
  const spPr = findChildByName(sp, 'spPr');
  if (!spPr) return null;

  // Parse bounds
  const bounds = parseBounds(spPr);
  if (!bounds) return null;

  // Parse rotation
  const rotation = parseRotation(spPr);

  // Parse shape type
  const shapeType = parseShapeType(spPr);

  // Parse fill
  const fill = parseFill(spPr, context.themeColors);

  // Parse stroke
  const stroke = parseStroke(spPr, context.themeColors);

  // Parse shadow
  const shadow = parseShadow(spPr, context.themeColors);

  // Parse text content
  const txBody = findChildByName(sp, 'txBody');
  const text = txBody ? parseTextBody(txBody, context.themeColors, context.relationships) : undefined;

  // Check if fill/stroke are visually empty
  const hasVisibleFill = fill && fill.type !== 'none';
  const hasVisibleStroke = stroke && stroke.width > 0;

  // If it's just a text box (no visible shape), return as TextElement
  const isTextBox = shapeType === 'rect' && !hasVisibleFill && !hasVisibleStroke && text;
  if (isTextBox) {
    return {
      id,
      type: 'text',
      bounds,
      rotation,
      text: text!,
      placeholder,
      shadow,
    };
  }

  return {
    id,
    type: 'shape',
    bounds,
    rotation,
    shapeType,
    fill,
    stroke,
    shadow,
    text,
    placeholder,
  };
}

/**
 * Parses a picture (pic) element.
 */
function parsePicture(pic: Element, context: ShapeParseContext): ImageElement | null {
  // Get non-visual properties for ID
  const nvPicPr = findChildByName(pic, 'nvPicPr');
  const cNvPr = nvPicPr ? findChildByName(nvPicPr, 'cNvPr') : null;
  const id = (cNvPr ? getAttribute(cNvPr, 'id') : null) || generateId();
  const altText = (cNvPr ? getAttribute(cNvPr, 'descr') : null) || undefined;

  // Get picture properties
  const spPr = findChildByName(pic, 'spPr');
  if (!spPr) return null;

  // Parse bounds
  const bounds = parseBounds(spPr);
  if (!bounds) return null;

  // Parse rotation
  const rotation = parseRotation(spPr);

  // Parse shadow
  const shadow = parseShadow(spPr, context.themeColors);

  // Get the image reference
  const blipFill = findChildByName(pic, 'blipFill');
  if (!blipFill) return null;

  const blip = findChildByName(blipFill, 'blip');
  if (!blip) return null;

  const rEmbed = getAttribute(blip, 'r:embed');
  if (!rEmbed) return null;

  // Resolve the image path
  const imagePath = context.relationships.resolvePath(rEmbed, context.basePath);
  if (!imagePath) return null;

  // Get the image as a blob URL
  const mimeType = getMimeType(imagePath);
  const src = context.archive.getBlobUrl(imagePath, mimeType);
  if (!src) return null;

  return {
    id,
    type: 'image',
    bounds,
    rotation,
    src,
    mimeType,
    altText,
    shadow,
  };
}

/**
 * Parses a group (grpSp) element.
 */
function parseGroup(grpSp: Element, context: ShapeParseContext): GroupElement | null {
  // Get non-visual properties for ID
  const nvGrpSpPr = findChildByName(grpSp, 'nvGrpSpPr');
  const cNvPr = nvGrpSpPr ? findChildByName(nvGrpSpPr, 'cNvPr') : null;
  const id = (cNvPr ? getAttribute(cNvPr, 'id') : null) || generateId();

  // Get group shape properties
  const grpSpPr = findChildByName(grpSp, 'grpSpPr');
  if (!grpSpPr) return null;

  // Parse bounds
  const bounds = parseBounds(grpSpPr);
  if (!bounds) return null;

  // Parse rotation
  const rotation = parseRotation(grpSpPr);

  // Parse children
  const children = parseShapeTree(grpSp, context);

  return {
    id,
    type: 'group',
    bounds,
    rotation,
    children,
  };
}

/**
 * Parses a graphic frame (charts, tables, SmartArt).
 * Attempts to parse tables directly, falls back to images for other content.
 */
function parseGraphicFrame(graphicFrame: Element, context: ShapeParseContext): TableElement | ImageElement | null {
  // Get non-visual properties for ID
  const nvGraphicFramePr = findChildByName(graphicFrame, 'nvGraphicFramePr');
  const cNvPr = nvGraphicFramePr ? findChildByName(nvGraphicFramePr, 'cNvPr') : null;
  const id = (cNvPr ? getAttribute(cNvPr, 'id') : null) || generateId();

  // Get transform
  const xfrm = findChildByName(graphicFrame, 'xfrm');
  if (!xfrm) return null;

  const bounds = parseBoundsFromXfrm(xfrm);
  if (!bounds) return null;

  const rotation = parseRotationFromXfrm(xfrm);

  // Look for graphic content
  const graphic = findChildByName(graphicFrame, 'graphic');
  if (!graphic) return null;

  const graphicData = findChildByName(graphic, 'graphicData');
  if (!graphicData) return null;

  // Check if this is a table
  const tbl = findChildByName(graphicData, 'tbl');
  if (tbl) {
    return parseTable(tbl, id, bounds, rotation, context);
  }

  // Try to find a fallback image for charts/SmartArt
  const blip = findFirstByName(graphic, 'blip');
  if (blip) {
    const rEmbed = getAttribute(blip, 'r:embed');
    if (rEmbed) {
      const imagePath = context.relationships.resolvePath(rEmbed, context.basePath);
      if (imagePath) {
        const mimeType = getMimeType(imagePath);
        const src = context.archive.getBlobUrl(imagePath, mimeType);
        if (src) {
          return {
            id,
            type: 'image',
            bounds,
            rotation,
            src,
            mimeType,
            altText: 'Chart or diagram',
          };
        }
      }
    }
  }

  // No parseable content found
  return null;
}

/**
 * Parses a table element.
 */
function parseTable(
  tbl: Element,
  id: string,
  bounds: Bounds,
  rotation: number | undefined,
  context: ShapeParseContext
): TableElement {
  // Parse table properties
  const tblPr = findChildByName(tbl, 'tblPr');
  const style = parseTableStyle(tblPr);

  // Parse column widths from tblGrid
  const tblGrid = findChildByName(tbl, 'tblGrid');
  const columnWidths: number[] = [];
  if (tblGrid) {
    const gridCols = findChildrenByName(tblGrid, 'gridCol');
    for (const gridCol of gridCols) {
      const w = getNumberAttribute(gridCol, 'w', 0);
      columnWidths.push(emuToPixels(w));
    }
  }

  // Parse rows
  const rows: TableRow[] = [];
  const trElements = findChildrenByName(tbl, 'tr');
  for (const tr of trElements) {
    const row = parseTableRow(tr, context);
    rows.push(row);
  }

  return {
    id,
    type: 'table',
    bounds,
    rotation,
    rows,
    columnWidths,
    style,
  };
}

/**
 * Parses table style properties.
 */
function parseTableStyle(tblPr: Element | null): TableStyle | undefined {
  if (!tblPr) return undefined;

  return {
    firstRow: getAttribute(tblPr, 'firstRow') === '1',
    lastRow: getAttribute(tblPr, 'lastRow') === '1',
    firstCol: getAttribute(tblPr, 'firstCol') === '1',
    lastCol: getAttribute(tblPr, 'lastCol') === '1',
    bandRow: getAttribute(tblPr, 'bandRow') === '1',
    bandCol: getAttribute(tblPr, 'bandCol') === '1',
  };
}

/**
 * Parses a table row.
 */
function parseTableRow(tr: Element, context: ShapeParseContext): TableRow {
  const h = getNumberAttribute(tr, 'h', 0);
  const height = emuToPixels(h);

  const cells: TableCell[] = [];
  const tcElements = findChildrenByName(tr, 'tc');
  for (const tc of tcElements) {
    const cell = parseTableCell(tc, context);
    cells.push(cell);
  }

  return { height, cells };
}

/**
 * Parses a table cell.
 */
function parseTableCell(tc: Element, context: ShapeParseContext): TableCell {
  const cell: TableCell = {};

  // Parse text content
  const txBody = findChildByName(tc, 'txBody');
  if (txBody) {
    cell.text = parseTextBody(txBody, context.themeColors, context.relationships);
  }

  // Parse cell properties
  const tcPr = findChildByName(tc, 'tcPr');
  if (tcPr) {
    // Vertical alignment
    const anchor = getAttribute(tcPr, 'anchor');
    if (anchor === 'ctr') {
      cell.verticalAlign = 'middle';
    } else if (anchor === 'b') {
      cell.verticalAlign = 'bottom';
    } else {
      cell.verticalAlign = 'top';
    }

    // Cell fill
    const solidFill = findChildByName(tcPr, 'solidFill');
    if (solidFill) {
      const color = parseColorElement(solidFill, context.themeColors);
      if (color) {
        cell.fill = { type: 'solid', color };
      }
    }

    // Cell borders
    cell.borders = parseCellBorders(tcPr, context.themeColors);
  }

  // Column span
  const gridSpan = getNumberAttribute(tc, 'gridSpan', 1);
  if (gridSpan > 1) {
    cell.colSpan = gridSpan;
  }

  // Row span
  const rowSpan = getNumberAttribute(tc, 'rowSpan', 1);
  if (rowSpan > 1) {
    cell.rowSpan = rowSpan;
  }

  return cell;
}

/**
 * Parses cell border properties.
 */
function parseCellBorders(tcPr: Element, themeColors: ThemeColors): CellBorders | undefined {
  const borders: CellBorders = {};
  let hasBorders = false;

  const borderNames = ['lnL', 'lnR', 'lnT', 'lnB'] as const;
  const borderKeys = ['left', 'right', 'top', 'bottom'] as const;

  for (let i = 0; i < borderNames.length; i++) {
    const ln = findChildByName(tcPr, borderNames[i]);
    if (ln) {
      const w = getNumberAttribute(ln, 'w', 12700); // Default 1pt
      const solidFill = findChildByName(ln, 'solidFill');
      if (solidFill) {
        const color = parseColorElement(solidFill, themeColors);
        if (color) {
          borders[borderKeys[i]] = {
            width: emuToPixels(w),
            color,
          };
          hasBorders = true;
        }
      }
    }
  }

  return hasBorders ? borders : undefined;
}

/**
 * Parses a connector shape.
 */
function parseConnector(cxnSp: Element, context: ShapeParseContext): ShapeElement | null {
  const nvCxnSpPr = findChildByName(cxnSp, 'nvCxnSpPr');
  const cNvPr = nvCxnSpPr ? findChildByName(nvCxnSpPr, 'cNvPr') : null;
  const id = (cNvPr ? getAttribute(cNvPr, 'id') : null) || generateId();

  const spPr = findChildByName(cxnSp, 'spPr');
  if (!spPr) return null;

  const bounds = parseBounds(spPr);
  if (!bounds) return null;

  const rotation = parseRotation(spPr);
  const stroke = parseStroke(spPr, context.themeColors);

  return {
    id,
    type: 'shape',
    bounds,
    rotation,
    shapeType: 'line',
    stroke,
  };
}

/**
 * Parses bounds from shape properties.
 */
function parseBounds(spPr: Element): Bounds | null {
  const xfrm = findChildByName(spPr, 'xfrm');
  if (!xfrm) return null;

  return parseBoundsFromXfrm(xfrm);
}

/**
 * Parses bounds from an xfrm element.
 */
function parseBoundsFromXfrm(xfrm: Element): Bounds | null {
  const off = findChildByName(xfrm, 'off');
  const ext = findChildByName(xfrm, 'ext');

  if (!off || !ext) return null;

  return {
    x: emuToPixels(getNumberAttribute(off, 'x', 0)),
    y: emuToPixels(getNumberAttribute(off, 'y', 0)),
    width: emuToPixels(getNumberAttribute(ext, 'cx', 0)),
    height: emuToPixels(getNumberAttribute(ext, 'cy', 0)),
  };
}

/**
 * Parses rotation from shape properties.
 */
function parseRotation(spPr: Element): number | undefined {
  const xfrm = findChildByName(spPr, 'xfrm');
  if (!xfrm) return undefined;

  return parseRotationFromXfrm(xfrm);
}

/**
 * Parses rotation from an xfrm element.
 */
function parseRotationFromXfrm(xfrm: Element): number | undefined {
  const rot = getNumberAttribute(xfrm, 'rot', 0);
  if (rot === 0) return undefined;

  return ooxmlAngleToDegrees(rot);
}

/**
 * Parses shape type from preset geometry.
 */
function parseShapeType(spPr: Element): ShapeType {
  const prstGeom = findChildByName(spPr, 'prstGeom');
  if (!prstGeom) {
    // Check for custom geometry
    const custGeom = findChildByName(spPr, 'custGeom');
    if (custGeom) return 'custom';
    return 'rect';
  }

  const prst = getAttribute(prstGeom, 'prst');

  // Map OOXML preset names to our shape types
  const shapeMap: Record<string, ShapeType> = {
    // Basic shapes
    rect: 'rect',
    roundRect: 'roundRect',
    snip1Rect: 'snip1Rect',
    snip2SameRect: 'snip2Rect',
    snip2DiagRect: 'snip2Rect',
    snipRoundRect: 'snip1Rect',
    round1Rect: 'roundRect',
    round2SameRect: 'roundRect',
    round2DiagRect: 'roundRect',
    ellipse: 'ellipse',
    triangle: 'triangle',
    rtTriangle: 'rtTriangle',
    diamond: 'diamond',
    parallelogram: 'parallelogram',
    trapezoid: 'trapezoid',
    pentagon: 'pentagon',
    hexagon: 'hexagon',
    heptagon: 'heptagon',
    octagon: 'octagon',
    decagon: 'decagon',
    dodecagon: 'dodecagon',

    // Stars
    star4: 'star4',
    star5: 'star5',
    star6: 'star6',
    star8: 'star8',
    star10: 'star10',
    star12: 'star12',
    star16: 'star8',
    star24: 'star12',
    star32: 'star12',

    // Arrows
    rightArrow: 'rightArrow',
    leftArrow: 'leftArrow',
    upArrow: 'upArrow',
    downArrow: 'downArrow',
    leftRightArrow: 'leftRightArrow',
    upDownArrow: 'upDownArrow',
    bentArrow: 'arrow',
    uturnArrow: 'arrow',
    curvedRightArrow: 'rightArrow',
    curvedLeftArrow: 'leftArrow',
    curvedUpArrow: 'upArrow',
    curvedDownArrow: 'downArrow',
    stripedRightArrow: 'rightArrow',
    notchedRightArrow: 'notchedRightArrow',
    chevron: 'chevron',
    homePlate: 'homePlate',

    // Callouts
    wedgeRectCallout: 'wedgeRectCallout',
    wedgeRoundRectCallout: 'wedgeRoundRectCallout',
    wedgeEllipseCallout: 'wedgeEllipseCallout',
    cloudCallout: 'cloudCallout',
    borderCallout1: 'wedgeRectCallout',
    borderCallout2: 'wedgeRectCallout',
    borderCallout3: 'wedgeRectCallout',
    accentCallout1: 'wedgeRectCallout',
    accentCallout2: 'wedgeRectCallout',
    accentCallout3: 'wedgeRectCallout',
    callout1: 'wedgeRectCallout',
    callout2: 'wedgeRectCallout',
    callout3: 'wedgeRectCallout',

    // Block shapes
    cube: 'cube',
    can: 'can',
    lightningBolt: 'lightningBolt',
    heart: 'heart',
    sun: 'sun',
    moon: 'moon',
    cloud: 'cloud',
    arc: 'arc',
    donut: 'donut',
    noSmoking: 'noSmoking',
    blockArc: 'blockArc',
    foldedCorner: 'foldedCorner',
    frame: 'frame',
    halfFrame: 'halfFrame',
    corner: 'corner',
    mathPlus: 'plus',
    plaque: 'rect',
    plus: 'plus',
    cross: 'cross',

    // Flowchart (map to basic shapes)
    flowChartProcess: 'rect',
    flowChartDecision: 'diamond',
    flowChartTerminator: 'roundRect',
    flowChartDocument: 'foldedCorner',
    flowChartPredefinedProcess: 'rect',
    flowChartConnector: 'ellipse',

    // Lines and connectors
    line: 'line',
    straightConnector1: 'line',
    bentConnector2: 'line',
    bentConnector3: 'bentConnector3',
    bentConnector4: 'bentConnector3',
    bentConnector5: 'bentConnector3',
    curvedConnector2: 'line',
    curvedConnector3: 'curvedConnector3',
    curvedConnector4: 'curvedConnector3',
    curvedConnector5: 'curvedConnector3',
  };

  return shapeMap[prst || ''] || 'rect';
}

/**
 * Parses fill style from shape properties.
 */
function parseFill(spPr: Element, themeColors: ThemeColors): Fill | undefined {
  // Check for no fill
  const noFill = findChildByName(spPr, 'noFill');
  if (noFill) return { type: 'none' };

  // Check for solid fill
  const solidFill = findChildByName(spPr, 'solidFill');
  if (solidFill) {
    const color = parseColorElement(solidFill, themeColors);
    if (color) {
      return { type: 'solid', color };
    }
  }

  // Check for gradient fill
  const gradFill = findChildByName(spPr, 'gradFill');
  if (gradFill) {
    return parseGradientFill(gradFill, themeColors);
  }

  // Check for pattern fill (treat as solid with foreground color)
  const pattFill = findChildByName(spPr, 'pattFill');
  if (pattFill) {
    const fgClr = findChildByName(pattFill, 'fgClr');
    if (fgClr) {
      const color = parseColorElement(fgClr, themeColors);
      if (color) {
        return { type: 'solid', color };
      }
    }
  }

  return undefined;
}

/**
 * Parses gradient fill.
 */
function parseGradientFill(gradFill: Element, themeColors: ThemeColors): Fill {
  const stops: GradientStop[] = [];

  // Parse gradient stops
  const gsLst = findChildByName(gradFill, 'gsLst');
  if (gsLst) {
    const gsElements = findChildrenByName(gsLst, 'gs');
    for (const gs of gsElements) {
      const pos = getNumberAttribute(gs, 'pos', 0) / 100000; // Convert to 0-1
      const color = parseColorElement(gs, themeColors);
      if (color) {
        stops.push({ position: pos, color });
      }
    }
  }

  // Parse gradient direction
  let angle = 0;
  const lin = findChildByName(gradFill, 'lin');
  if (lin) {
    angle = ooxmlAngleToDegrees(getNumberAttribute(lin, 'ang', 0));
  }

  // Sort stops by position
  stops.sort((a, b) => a.position - b.position);

  return {
    type: 'gradient',
    angle,
    stops,
  };
}

/**
 * Parses stroke/outline style from shape properties.
 */
function parseStroke(spPr: Element, themeColors: ThemeColors): Stroke | undefined {
  const ln = findChildByName(spPr, 'ln');
  if (!ln) return undefined;

  // Check for no line
  const noFill = findChildByName(ln, 'noFill');
  if (noFill) return undefined;

  // Parse color - only return stroke if there's an explicit fill
  const solidFill = findChildByName(ln, 'solidFill');
  if (!solidFill) {
    // No explicit stroke color defined - treat as no stroke
    return undefined;
  }

  const parsedColor = parseColorElement(solidFill, themeColors);
  if (!parsedColor) return undefined;

  // Parse width
  const width = emuToPixels(getNumberAttribute(ln, 'w', 12700)); // Default ~1px
  const color: Color = parsedColor;

  // Parse dash style
  const prstDash = findChildByName(ln, 'prstDash');
  let dashStyle: Stroke['dashStyle'] = 'solid';
  if (prstDash) {
    const val = getAttribute(prstDash, 'val');
    switch (val) {
      case 'dash':
        dashStyle = 'dash';
        break;
      case 'dot':
        dashStyle = 'dot';
        break;
      case 'dashDot':
        dashStyle = 'dashDot';
        break;
    }
  }

  // Parse arrow heads
  const headEnd = parseArrowHead(findChildByName(ln, 'headEnd'));
  const tailEnd = parseArrowHead(findChildByName(ln, 'tailEnd'));

  return {
    color,
    width,
    dashStyle,
    headEnd,
    tailEnd,
  };
}

/**
 * Parses an arrow head element.
 */
function parseArrowHead(element: Element | null): ArrowHead | undefined {
  if (!element) return undefined;

  const typeAttr = getAttribute(element, 'type');
  if (!typeAttr || typeAttr === 'none') return undefined;

  // Map OOXML arrow types to our types
  let type: ArrowHead['type'];
  switch (typeAttr) {
    case 'triangle':
      type = 'triangle';
      break;
    case 'stealth':
      type = 'stealth';
      break;
    case 'diamond':
      type = 'diamond';
      break;
    case 'oval':
      type = 'oval';
      break;
    case 'arrow':
      type = 'arrow';
      break;
    default:
      type = 'triangle'; // Default to triangle for unknown types
  }

  const widthAttr = getAttribute(element, 'w');
  const width = (widthAttr === 'sm' || widthAttr === 'med' || widthAttr === 'lg')
    ? widthAttr as ArrowHead['width']
    : 'med';

  const lengthAttr = getAttribute(element, 'len');
  const length = (lengthAttr === 'sm' || lengthAttr === 'med' || lengthAttr === 'lg')
    ? lengthAttr as ArrowHead['length']
    : 'med';

  return { type, width, length };
}

/**
 * Parses shadow effects from shape properties.
 * Shadows are defined in effectLst with outerShdw or innerShdw elements.
 */
function parseShadow(spPr: Element, themeColors: ThemeColors): Shadow | undefined {
  const effectLst = findChildByName(spPr, 'effectLst');
  if (!effectLst) return undefined;

  // Check for outer shadow (most common)
  const outerShdw = findChildByName(effectLst, 'outerShdw');
  if (outerShdw) {
    return parseShadowElement(outerShdw, 'outer', themeColors);
  }

  // Check for inner shadow
  const innerShdw = findChildByName(effectLst, 'innerShdw');
  if (innerShdw) {
    return parseShadowElement(innerShdw, 'inner', themeColors);
  }

  return undefined;
}

/**
 * Parses a shadow element (outerShdw or innerShdw).
 */
function parseShadowElement(
  shdw: Element,
  type: 'outer' | 'inner',
  themeColors: ThemeColors
): Shadow | undefined {
  // Parse blur radius (in EMUs)
  const blurRad = getNumberAttribute(shdw, 'blurRad', 0);
  const blurRadius = emuToPixels(blurRad);

  // Parse distance (in EMUs)
  const dist = getNumberAttribute(shdw, 'dist', 0);
  const distance = emuToPixels(dist);

  // Parse direction (in 1/60000 of a degree)
  const dir = getNumberAttribute(shdw, 'dir', 0);
  const angle = dir / 60000; // Convert to degrees

  // Parse color
  const color = parseColorElement(shdw, themeColors);
  if (!color) {
    // Default shadow color if not specified
    return {
      type,
      color: { hex: '#000000', alpha: 0.4 },
      blurRadius: blurRadius || 4,
      distance: distance || 3,
      angle,
    };
  }

  return {
    type,
    color,
    blurRadius,
    distance,
    angle,
  };
}

/**
 * Parses placeholder information from non-visual shape properties.
 * Placeholders are used for slide master/layout inheritance.
 */
function parsePlaceholderInfo(nvSpPr: Element | null): PlaceholderInfo | undefined {
  if (!nvSpPr) return undefined;

  const nvPr = findChildByName(nvSpPr, 'nvPr');
  if (!nvPr) return undefined;

  const ph = findChildByName(nvPr, 'ph');
  if (!ph) return undefined;

  // Get placeholder type (defaults to 'body' if not specified)
  const typeAttr = getAttribute(ph, 'type');
  const type = (typeAttr || 'body') as PlaceholderType;

  // Get placeholder index (for matching with layout/master)
  const idxAttr = getAttribute(ph, 'idx');
  const idx = idxAttr ? parseInt(idxAttr, 10) : undefined;

  return { type, idx };
}

/**
 * Generates a unique ID for elements without one.
 */
let idCounter = 0;
function generateId(): string {
  return `elem_${++idCounter}`;
}
