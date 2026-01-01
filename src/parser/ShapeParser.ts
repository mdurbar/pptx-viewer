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
  Bounds,
  ShapeType,
  Fill,
  Stroke,
  Color,
  GradientStop,
  ThemeColors,
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
    text,
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
 * Attempts to use fallback image.
 */
function parseGraphicFrame(graphicFrame: Element, context: ShapeParseContext): ImageElement | null {
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

  // Look for fallback image in the graphic element
  const graphic = findChildByName(graphicFrame, 'graphic');
  if (!graphic) return null;

  // Try to find an image reference
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

  // No fallback image found
  return null;
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

  return {
    color,
    width,
    dashStyle,
  };
}

/**
 * Generates a unique ID for elements without one.
 */
let idCounter = 0;
function generateId(): string {
  return `elem_${++idCounter}`;
}
