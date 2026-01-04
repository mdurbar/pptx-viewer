/**
 * Renderer for shapes.
 *
 * Converts parsed shape elements into SVG elements.
 */

import type {
  SlideElement,
  ShapeElement,
  TextElement,
  ImageElement,
  GroupElement,
  TableElement,
  Fill,
  Stroke,
  ArrowHead,
  Shadow,
  ShapeType,
  Bounds,
  PatternType,
  Color,
} from '../core/types';
import { colorToCss } from '../utils/color';
import { renderTextBodyToSvg } from './TextRenderer';
import { renderTable } from './TableRenderer';

/**
 * Unique ID counter for SVG gradients and patterns.
 */
let defIdCounter = 0;

/**
 * Renders a slide element to SVG.
 *
 * @param element - The element to render
 * @param defs - SVG defs element for gradients/patterns
 * @returns SVG element group
 */
export function renderElement(element: SlideElement, defs: SVGDefsElement): SVGGElement {
  const group = document.createElementNS('http://www.w3.org/2000/svg', 'g');

  // Apply transform (position and rotation)
  const transform = buildTransform(element.bounds, element.rotation);
  if (transform) {
    group.setAttribute('transform', transform);
  }

  // Apply element-level opacity
  if (element.opacity !== undefined && element.opacity < 1) {
    group.setAttribute('opacity', String(element.opacity));
  }

  switch (element.type) {
    case 'shape':
      renderShape(element, group, defs);
      break;
    case 'text':
      renderTextBox(element, group, defs);
      break;
    case 'image':
      renderImage(element, group, defs);
      break;
    case 'group':
      renderGroup(element, group, defs);
      break;
    case 'table':
      renderTableElement(element, group);
      break;
  }

  return group;
}

/**
 * Builds a transform string for positioning and rotation.
 */
function buildTransform(bounds: Bounds, rotation?: number): string {
  const transforms: string[] = [];

  // Translate to position
  transforms.push(`translate(${bounds.x}, ${bounds.y})`);

  // Rotate around center
  if (rotation) {
    const cx = bounds.width / 2;
    const cy = bounds.height / 2;
    transforms.push(`rotate(${rotation}, ${cx}, ${cy})`);
  }

  return transforms.join(' ');
}

/**
 * Renders a shape element.
 */
function renderShape(shape: ShapeElement, group: SVGGElement, defs: SVGDefsElement): void {
  const { bounds, shapeType, fill, stroke, shadow, text, adjustments, flipH, flipV } = shape;

  // Check if shape has any visible fill or stroke
  const hasVisibleFill = fill && fill.type !== 'none';
  const hasVisibleStroke = stroke && stroke.width > 0;

  // Only create shape element if it has visible fill or stroke
  if (hasVisibleFill || hasVisibleStroke) {
    const shapeEl = createShapeElement(shapeType, bounds.width, bounds.height, adjustments, flipH, flipV);

    // Apply fill
    if (hasVisibleFill) {
      applyFill(shapeEl, fill, defs);
    } else {
      shapeEl.setAttribute('fill', 'none');
    }

    // Apply stroke
    if (hasVisibleStroke) {
      applyStroke(shapeEl, stroke, defs);
    } else {
      shapeEl.setAttribute('stroke', 'none');
    }

    // Apply shadow filter
    if (shadow) {
      const filterId = applyShadowFilter(shadow, defs);
      shapeEl.setAttribute('filter', `url(#${filterId})`);
    }

    group.appendChild(shapeEl);
  }

  // Add text if present
  if (text && text.paragraphs.length > 0) {
    const textFo = renderTextBodyToSvg(text, bounds.width, bounds.height);
    group.appendChild(textFo);
  }
}

/**
 * Creates an SVG element for a shape type.
 */
function createShapeElement(
  shapeType: ShapeType,
  width: number,
  height: number,
  adjustments?: Map<string, number>,
  flipH?: boolean,
  flipV?: boolean
): SVGElement {
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  let d: string;

  switch (shapeType) {
    case 'ellipse': {
      const ellipse = document.createElementNS('http://www.w3.org/2000/svg', 'ellipse');
      ellipse.setAttribute('cx', String(width / 2));
      ellipse.setAttribute('cy', String(height / 2));
      ellipse.setAttribute('rx', String(width / 2));
      ellipse.setAttribute('ry', String(height / 2));
      return ellipse;
    }

    case 'roundRect': {
      const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect.setAttribute('width', String(width));
      rect.setAttribute('height', String(height));
      // Use adjustment value if available, otherwise default to 1/8
      // The 'adj' value is typically 0-0.5 (0-50%) representing corner radius as fraction of min dimension
      const adjValue = adjustments?.get('adj') ?? 0.125; // Default ~12.5%
      const radius = Math.min(width, height) * adjValue;
      rect.setAttribute('rx', String(radius));
      rect.setAttribute('ry', String(radius));
      return rect;
    }

    case 'snip1Rect': {
      // Use adjustment for snip size
      const adjValue = adjustments?.get('adj') ?? 0.15;
      const snip = Math.min(width, height) * adjValue;
      d = `M 0 0 L ${width - snip} 0 L ${width} ${snip} L ${width} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'snip2Rect': {
      // Use adjustment for snip size
      const adjValue = adjustments?.get('adj1') ?? adjustments?.get('adj') ?? 0.15;
      const snip = Math.min(width, height) * adjValue;
      d = `M ${snip} 0 L ${width - snip} 0 L ${width} ${snip} L ${width} ${height - snip} L ${width - snip} ${height} L ${snip} ${height} L 0 ${height - snip} L 0 ${snip} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'triangle':
      d = `M ${width / 2} 0 L ${width} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;

    case 'rtTriangle':
      d = `M 0 0 L ${width} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;

    case 'diamond':
      d = `M ${width / 2} 0 L ${width} ${height / 2} L ${width / 2} ${height} L 0 ${height / 2} Z`;
      path.setAttribute('d', d);
      return path;

    case 'parallelogram': {
      const offset = width * 0.2;
      d = `M ${offset} 0 L ${width} 0 L ${width - offset} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'trapezoid': {
      const offset = width * 0.15;
      d = `M ${offset} 0 L ${width - offset} 0 L ${width} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'pentagon':
    case 'hexagon':
    case 'heptagon':
    case 'octagon':
    case 'decagon':
    case 'dodecagon': {
      const sides = {
        pentagon: 5, hexagon: 6, heptagon: 7, octagon: 8, decagon: 10, dodecagon: 12
      }[shapeType] || 6;
      const cx = width / 2;
      const cy = height / 2;
      const r = Math.min(width, height) / 2;
      const points = generatePolygonPoints(cx, cy, r, sides, -90);
      path.setAttribute('d', `M ${points.join(' L ')} Z`);
      return path;
    }

    // Stars
    case 'star4':
    case 'star5':
    case 'star6':
    case 'star8':
    case 'star10':
    case 'star12': {
      const starPoints = { star4: 4, star5: 5, star6: 6, star8: 8, star10: 10, star12: 12 }[shapeType] || 5;
      const cx = width / 2;
      const cy = height / 2;
      const outerR = Math.min(width, height) / 2;
      // Use adjustment for inner radius ratio (adj typically 0.19-0.5)
      const innerRatio = adjustments?.get('adj') ?? 0.4;
      const innerR = outerR * innerRatio;
      const starPath = generateStarPoints(cx, cy, outerR, innerR, starPoints);
      path.setAttribute('d', starPath);
      return path;
    }

    // Arrows
    case 'rightArrow':
    case 'arrow': {
      const tailWidth = height * 0.4;
      const tailY = (height - tailWidth) / 2;
      const headStart = width * 0.6;
      d = `M 0 ${tailY} L ${headStart} ${tailY} L ${headStart} 0 L ${width} ${height / 2} L ${headStart} ${height} L ${headStart} ${tailY + tailWidth} L 0 ${tailY + tailWidth} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'leftArrow': {
      const tailWidth = height * 0.4;
      const tailY = (height - tailWidth) / 2;
      const headEnd = width * 0.4;
      d = `M ${width} ${tailY} L ${headEnd} ${tailY} L ${headEnd} 0 L 0 ${height / 2} L ${headEnd} ${height} L ${headEnd} ${tailY + tailWidth} L ${width} ${tailY + tailWidth} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'upArrow': {
      const tailWidth = width * 0.4;
      const tailX = (width - tailWidth) / 2;
      const headEnd = height * 0.4;
      d = `M ${tailX} ${height} L ${tailX} ${headEnd} L 0 ${headEnd} L ${width / 2} 0 L ${width} ${headEnd} L ${tailX + tailWidth} ${headEnd} L ${tailX + tailWidth} ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'downArrow': {
      const tailWidth = width * 0.4;
      const tailX = (width - tailWidth) / 2;
      const headStart = height * 0.6;
      d = `M ${tailX} 0 L ${tailX} ${headStart} L 0 ${headStart} L ${width / 2} ${height} L ${width} ${headStart} L ${tailX + tailWidth} ${headStart} L ${tailX + tailWidth} 0 Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'leftRightArrow': {
      const tailWidth = height * 0.4;
      const tailY = (height - tailWidth) / 2;
      const headWidth = width * 0.2;
      d = `M 0 ${height / 2} L ${headWidth} 0 L ${headWidth} ${tailY} L ${width - headWidth} ${tailY} L ${width - headWidth} 0 L ${width} ${height / 2} L ${width - headWidth} ${height} L ${width - headWidth} ${tailY + tailWidth} L ${headWidth} ${tailY + tailWidth} L ${headWidth} ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'upDownArrow': {
      const tailWidth = width * 0.4;
      const tailX = (width - tailWidth) / 2;
      const headHeight = height * 0.2;
      d = `M ${width / 2} 0 L ${width} ${headHeight} L ${tailX + tailWidth} ${headHeight} L ${tailX + tailWidth} ${height - headHeight} L ${width} ${height - headHeight} L ${width / 2} ${height} L 0 ${height - headHeight} L ${tailX} ${height - headHeight} L ${tailX} ${headHeight} L 0 ${headHeight} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'chevron': {
      const notch = width * 0.25;
      d = `M 0 0 L ${width - notch} 0 L ${width} ${height / 2} L ${width - notch} ${height} L 0 ${height} L ${notch} ${height / 2} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'homePlate': {
      const notch = width * 0.2;
      d = `M 0 0 L ${width - notch} 0 L ${width} ${height / 2} L ${width - notch} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'notchedRightArrow': {
      const tailWidth = height * 0.4;
      const tailY = (height - tailWidth) / 2;
      const headStart = width * 0.6;
      const notch = width * 0.1;
      d = `M ${notch} ${tailY} L ${headStart} ${tailY} L ${headStart} 0 L ${width} ${height / 2} L ${headStart} ${height} L ${headStart} ${tailY + tailWidth} L ${notch} ${tailY + tailWidth} L 0 ${height / 2} Z`;
      path.setAttribute('d', d);
      return path;
    }

    // Callouts
    case 'wedgeRectCallout': {
      const tipX = width * 0.1;
      const tipY = height + height * 0.2;
      d = `M 0 0 L ${width} 0 L ${width} ${height} L ${width * 0.4} ${height} L ${tipX} ${tipY} L ${width * 0.2} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'wedgeRoundRectCallout': {
      const r = Math.min(width, height) * 0.1;
      const tipX = width * 0.1;
      const tipY = height + height * 0.2;
      d = `M ${r} 0 L ${width - r} 0 Q ${width} 0 ${width} ${r} L ${width} ${height - r} Q ${width} ${height} ${width - r} ${height} L ${width * 0.4} ${height} L ${tipX} ${tipY} L ${width * 0.2} ${height} L ${r} ${height} Q 0 ${height} 0 ${height - r} L 0 ${r} Q 0 0 ${r} 0 Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'wedgeEllipseCallout': {
      const cx = width / 2;
      const cy = height / 2;
      const rx = width / 2;
      const ry = height / 2;
      const tipX = width * 0.1;
      const tipY = height + height * 0.3;
      // Simplified: ellipse with pointer
      d = `M ${cx + rx} ${cy} A ${rx} ${ry} 0 1 1 ${cx - rx} ${cy} A ${rx} ${ry} 0 0 1 ${cx * 0.6} ${cy + ry * 0.8} L ${tipX} ${tipY} L ${cx * 0.8} ${cy + ry * 0.9} A ${rx} ${ry} 0 0 1 ${cx + rx} ${cy} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'cloudCallout': {
      // Simplified cloud shape
      d = `M ${width * 0.15} ${height * 0.6}
           Q 0 ${height * 0.5} ${width * 0.1} ${height * 0.35}
           Q 0 ${height * 0.15} ${width * 0.25} ${height * 0.15}
           Q ${width * 0.3} 0 ${width * 0.5} ${height * 0.1}
           Q ${width * 0.7} 0 ${width * 0.8} ${height * 0.2}
           Q ${width} ${height * 0.15} ${width * 0.95} ${height * 0.4}
           Q ${width} ${height * 0.6} ${width * 0.85} ${height * 0.7}
           Q ${width * 0.9} ${height * 0.9} ${width * 0.7} ${height * 0.85}
           Q ${width * 0.5} ${height} ${width * 0.3} ${height * 0.85}
           Q ${width * 0.1} ${height * 0.9} ${width * 0.15} ${height * 0.6} Z`;
      path.setAttribute('d', d);
      return path;
    }

    // Block shapes
    case 'heart': {
      const w = width;
      const h = height;
      d = `M ${w / 2} ${h * 0.3}
           C ${w * 0.1} ${h * -0.1} ${w * -0.2} ${h * 0.4} ${w / 2} ${h}
           C ${w * 1.2} ${h * 0.4} ${w * 0.9} ${h * -0.1} ${w / 2} ${h * 0.3} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'lightningBolt': {
      d = `M ${width * 0.4} 0 L ${width * 0.7} 0 L ${width * 0.5} ${height * 0.4} L ${width} ${height * 0.4} L ${width * 0.35} ${height} L ${width * 0.45} ${height * 0.55} L 0 ${height * 0.55} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'sun': {
      const cx = width / 2;
      const cy = height / 2;
      const outerR = Math.min(width, height) / 2;
      const innerR = outerR * 0.6;
      const starPath = generateStarPoints(cx, cy, outerR, innerR, 12);
      path.setAttribute('d', starPath);
      return path;
    }

    case 'moon': {
      const cx = width / 2;
      const cy = height / 2;
      const r = Math.min(width, height) / 2;
      d = `M ${cx} ${cy - r} A ${r} ${r} 0 1 1 ${cx} ${cy + r} A ${r * 0.6} ${r} 0 1 0 ${cx} ${cy - r} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'cloud': {
      d = `M ${width * 0.15} ${height * 0.6}
           Q 0 ${height * 0.5} ${width * 0.1} ${height * 0.35}
           Q 0 ${height * 0.15} ${width * 0.25} ${height * 0.15}
           Q ${width * 0.3} 0 ${width * 0.5} ${height * 0.1}
           Q ${width * 0.7} 0 ${width * 0.8} ${height * 0.2}
           Q ${width} ${height * 0.15} ${width * 0.95} ${height * 0.4}
           Q ${width} ${height * 0.6} ${width * 0.85} ${height * 0.7}
           Q ${width * 0.9} ${height * 0.9} ${width * 0.7} ${height * 0.85}
           Q ${width * 0.5} ${height} ${width * 0.3} ${height * 0.85}
           Q ${width * 0.1} ${height * 0.9} ${width * 0.15} ${height * 0.6} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'donut': {
      const cx = width / 2;
      const cy = height / 2;
      const outerR = Math.min(width, height) / 2;
      const innerR = outerR * 0.5;
      d = `M ${cx + outerR} ${cy} A ${outerR} ${outerR} 0 1 0 ${cx - outerR} ${cy} A ${outerR} ${outerR} 0 1 0 ${cx + outerR} ${cy} Z M ${cx + innerR} ${cy} A ${innerR} ${innerR} 0 1 1 ${cx - innerR} ${cy} A ${innerR} ${innerR} 0 1 1 ${cx + innerR} ${cy} Z`;
      path.setAttribute('d', d);
      path.setAttribute('fill-rule', 'evenodd');
      return path;
    }

    case 'noSmoking': {
      const cx = width / 2;
      const cy = height / 2;
      const outerR = Math.min(width, height) / 2;
      const innerR = outerR * 0.8;
      const barWidth = outerR * 0.15;
      // Circle with diagonal bar
      d = `M ${cx + outerR} ${cy} A ${outerR} ${outerR} 0 1 0 ${cx - outerR} ${cy} A ${outerR} ${outerR} 0 1 0 ${cx + outerR} ${cy} Z M ${cx + innerR} ${cy} A ${innerR} ${innerR} 0 1 1 ${cx - innerR} ${cy} A ${innerR} ${innerR} 0 1 1 ${cx + innerR} ${cy} Z`;
      path.setAttribute('d', d);
      path.setAttribute('fill-rule', 'evenodd');
      return path;
    }

    case 'plus': {
      const armWidth = Math.min(width, height) * 0.33;
      const hGap = (width - armWidth) / 2;
      const vGap = (height - armWidth) / 2;
      d = `M ${hGap} 0 L ${hGap + armWidth} 0 L ${hGap + armWidth} ${vGap} L ${width} ${vGap} L ${width} ${vGap + armWidth} L ${hGap + armWidth} ${vGap + armWidth} L ${hGap + armWidth} ${height} L ${hGap} ${height} L ${hGap} ${vGap + armWidth} L 0 ${vGap + armWidth} L 0 ${vGap} L ${hGap} ${vGap} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'cross': {
      const armWidth = Math.min(width, height) * 0.25;
      const hGap = (width - armWidth) / 2;
      const vGap = (height - armWidth) / 2;
      d = `M ${hGap} 0 L ${hGap + armWidth} 0 L ${hGap + armWidth} ${vGap} L ${width} ${vGap} L ${width} ${vGap + armWidth} L ${hGap + armWidth} ${vGap + armWidth} L ${hGap + armWidth} ${height} L ${hGap} ${height} L ${hGap} ${vGap + armWidth} L 0 ${vGap + armWidth} L 0 ${vGap} L ${hGap} ${vGap} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'foldedCorner': {
      const fold = Math.min(width, height) * 0.2;
      d = `M 0 0 L ${width - fold} 0 L ${width} ${fold} L ${width} ${height} L 0 ${height} Z M ${width - fold} 0 L ${width - fold} ${fold} L ${width} ${fold}`;
      path.setAttribute('d', d);
      return path;
    }

    case 'frame': {
      const border = Math.min(width, height) * 0.1;
      d = `M 0 0 L ${width} 0 L ${width} ${height} L 0 ${height} Z M ${border} ${border} L ${border} ${height - border} L ${width - border} ${height - border} L ${width - border} ${border} Z`;
      path.setAttribute('d', d);
      path.setAttribute('fill-rule', 'evenodd');
      return path;
    }

    case 'cube': {
      const depth = Math.min(width, height) * 0.2;
      d = `M 0 ${depth} L ${depth} 0 L ${width} 0 L ${width} ${height - depth} L ${width - depth} ${height} L 0 ${height} Z M 0 ${depth} L ${width - depth} ${depth} L ${width - depth} ${height} M ${width - depth} ${depth} L ${width} 0`;
      path.setAttribute('d', d);
      return path;
    }

    case 'can': {
      const ellipseHeight = height * 0.15;
      d = `M 0 ${ellipseHeight} L 0 ${height - ellipseHeight} A ${width / 2} ${ellipseHeight} 0 0 0 ${width} ${height - ellipseHeight} L ${width} ${ellipseHeight} A ${width / 2} ${ellipseHeight} 0 0 0 0 ${ellipseHeight} Z M 0 ${ellipseHeight} A ${width / 2} ${ellipseHeight} 0 0 1 ${width} ${ellipseHeight}`;
      path.setAttribute('d', d);
      return path;
    }

    case 'line': {
      // Line with flip support
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      const x1 = flipH ? width : 0;
      const y1 = flipV ? height : 0;
      const x2 = flipH ? 0 : width;
      const y2 = flipV ? 0 : height;
      line.setAttribute('x1', String(x1));
      line.setAttribute('y1', String(y1));
      line.setAttribute('x2', String(x2));
      line.setAttribute('y2', String(y2));
      return line;
    }

    case 'bentConnector3': {
      // Bent connector (elbow) - 3 segments with one corner
      const adj = adjustments?.get('adj1') ?? 50000; // Default to 50%
      const midX = width * (adj / 100000);

      // Determine start and end points based on flip
      const startX = flipH ? width : 0;
      const startY = flipV ? height : 0;
      const endX = flipH ? 0 : width;
      const endY = flipV ? 0 : height;

      // Draw L-shaped path
      d = `M ${startX} ${startY} L ${midX} ${startY} L ${midX} ${endY} L ${endX} ${endY}`;
      path.setAttribute('d', d);
      path.setAttribute('fill', 'none');
      return path;
    }

    case 'curvedConnector3': {
      // Curved connector using bezier curve
      const startX = flipH ? width : 0;
      const startY = flipV ? height : 0;
      const endX = flipH ? 0 : width;
      const endY = flipV ? 0 : height;

      // Control points for smooth S-curve
      const cp1x = (startX + endX) / 2;
      const cp1y = startY;
      const cp2x = (startX + endX) / 2;
      const cp2y = endY;

      d = `M ${startX} ${startY} C ${cp1x} ${cp1y}, ${cp2x} ${cp2y}, ${endX} ${endY}`;
      path.setAttribute('d', d);
      path.setAttribute('fill', 'none');
      return path;
    }

    case 'rect':
    default: {
      const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect.setAttribute('width', String(width));
      rect.setAttribute('height', String(height));
      return rect;
    }
  }
}

/**
 * Generates polygon points for regular polygons.
 */
function generatePolygonPoints(
  cx: number,
  cy: number,
  radius: number,
  sides: number,
  startAngle: number
): string[] {
  const points: string[] = [];
  const angleStep = (2 * Math.PI) / sides;
  const startRad = (startAngle * Math.PI) / 180;

  for (let i = 0; i < sides; i++) {
    const angle = startRad + i * angleStep;
    const x = cx + radius * Math.cos(angle);
    const y = cy + radius * Math.sin(angle);
    points.push(`${x},${y}`);
  }

  return points;
}

/**
 * Generates SVG path for star shapes.
 */
function generateStarPoints(
  cx: number,
  cy: number,
  outerRadius: number,
  innerRadius: number,
  points: number
): string {
  const pathParts: string[] = [];
  const angleStep = Math.PI / points;
  const startAngle = -Math.PI / 2; // Start at top

  for (let i = 0; i < points * 2; i++) {
    const radius = i % 2 === 0 ? outerRadius : innerRadius;
    const angle = startAngle + i * angleStep;
    const x = cx + radius * Math.cos(angle);
    const y = cy + radius * Math.sin(angle);

    if (i === 0) {
      pathParts.push(`M ${x} ${y}`);
    } else {
      pathParts.push(`L ${x} ${y}`);
    }
  }

  pathParts.push('Z');
  return pathParts.join(' ');
}

/**
 * Applies fill to an SVG element.
 */
function applyFill(element: SVGElement, fill: Fill, defs: SVGDefsElement): void {
  switch (fill.type) {
    case 'solid':
      element.setAttribute('fill', colorToCss(fill.color));
      if (fill.color.alpha < 1) {
        element.setAttribute('fill-opacity', String(fill.color.alpha));
      }
      break;

    case 'gradient': {
      const gradientId = `gradient_${++defIdCounter}`;

      if (fill.gradientType === 'radial') {
        // Radial gradient
        const gradient = document.createElementNS('http://www.w3.org/2000/svg', 'radialGradient');
        gradient.setAttribute('id', gradientId);

        // Calculate center from fillToRect (defaults to center)
        const rect = fill.fillToRect || { left: 0.5, top: 0.5, right: 0.5, bottom: 0.5 };
        const cx = (rect.left + (1 - rect.right)) / 2 * 100;
        const cy = (rect.top + (1 - rect.bottom)) / 2 * 100;

        gradient.setAttribute('cx', `${cx}%`);
        gradient.setAttribute('cy', `${cy}%`);
        gradient.setAttribute('r', '70.71%'); // sqrt(2)/2 * 100 to cover corners
        gradient.setAttribute('fx', `${cx}%`);
        gradient.setAttribute('fy', `${cy}%`);

        // Add stops
        for (const stop of fill.stops) {
          const stopEl = document.createElementNS('http://www.w3.org/2000/svg', 'stop');
          stopEl.setAttribute('offset', `${stop.position * 100}%`);
          stopEl.setAttribute('stop-color', stop.color.hex);
          if (stop.color.alpha < 1) {
            stopEl.setAttribute('stop-opacity', String(stop.color.alpha));
          }
          gradient.appendChild(stopEl);
        }

        defs.appendChild(gradient);
      } else {
        // Linear gradient
        const gradient = document.createElementNS('http://www.w3.org/2000/svg', 'linearGradient');
        gradient.setAttribute('id', gradientId);

        // Set gradient angle
        const angle = fill.angle || 0;
        const radians = (angle * Math.PI) / 180;
        gradient.setAttribute('x1', String(50 - 50 * Math.cos(radians)) + '%');
        gradient.setAttribute('y1', String(50 - 50 * Math.sin(radians)) + '%');
        gradient.setAttribute('x2', String(50 + 50 * Math.cos(radians)) + '%');
        gradient.setAttribute('y2', String(50 + 50 * Math.sin(radians)) + '%');

        // Add stops
        for (const stop of fill.stops) {
          const stopEl = document.createElementNS('http://www.w3.org/2000/svg', 'stop');
          stopEl.setAttribute('offset', `${stop.position * 100}%`);
          stopEl.setAttribute('stop-color', stop.color.hex);
          if (stop.color.alpha < 1) {
            stopEl.setAttribute('stop-opacity', String(stop.color.alpha));
          }
          gradient.appendChild(stopEl);
        }

        defs.appendChild(gradient);
      }

      element.setAttribute('fill', `url(#${gradientId})`);
      break;
    }

    case 'pattern': {
      const patternId = `pattern_${++defIdCounter}`;
      const pattern = document.createElementNS('http://www.w3.org/2000/svg', 'pattern');
      pattern.setAttribute('id', patternId);
      pattern.setAttribute('patternUnits', 'userSpaceOnUse');

      // Get pattern size based on pattern type
      const patternSize = getPatternSize(fill.pattern);
      pattern.setAttribute('width', String(patternSize));
      pattern.setAttribute('height', String(patternSize));

      // Create pattern content
      const patternContent = createPatternContent(fill.pattern, fill.foreground, fill.background, patternSize);
      pattern.appendChild(patternContent);

      defs.appendChild(pattern);
      element.setAttribute('fill', `url(#${patternId})`);
      break;
    }

    case 'image': {
      const patternId = `pattern_${++defIdCounter}`;
      const pattern = document.createElementNS('http://www.w3.org/2000/svg', 'pattern');
      pattern.setAttribute('id', patternId);
      pattern.setAttribute('patternUnits', 'objectBoundingBox');
      pattern.setAttribute('width', '1');
      pattern.setAttribute('height', '1');

      const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
      image.setAttribute('href', fill.src);
      image.setAttribute('width', '100%');
      image.setAttribute('height', '100%');
      image.setAttribute('preserveAspectRatio', 'xMidYMid slice');

      pattern.appendChild(image);
      defs.appendChild(pattern);
      element.setAttribute('fill', `url(#${patternId})`);
      break;
    }

    case 'none':
      element.setAttribute('fill', 'none');
      break;
  }
}

/**
 * Applies stroke to an SVG element.
 */
function applyStroke(element: SVGElement, stroke: Stroke, defs: SVGDefsElement): void {
  element.setAttribute('stroke', colorToCss(stroke.color));
  element.setAttribute('stroke-width', String(stroke.width));

  if (stroke.color.alpha < 1) {
    element.setAttribute('stroke-opacity', String(stroke.color.alpha));
  }

  // Apply dash pattern
  if (stroke.dashStyle && stroke.dashStyle !== 'solid') {
    let dashArray: string;
    switch (stroke.dashStyle) {
      case 'dash':
        dashArray = `${stroke.width * 4} ${stroke.width * 2}`;
        break;
      case 'dot':
        dashArray = `${stroke.width} ${stroke.width}`;
        break;
      case 'dashDot':
        dashArray = `${stroke.width * 4} ${stroke.width * 2} ${stroke.width} ${stroke.width * 2}`;
        break;
      default:
        dashArray = '';
    }
    if (dashArray) {
      element.setAttribute('stroke-dasharray', dashArray);
    }
  }

  // Apply arrow markers
  if (stroke.headEnd) {
    const markerId = createArrowMarker(stroke.headEnd, stroke, defs, 'start');
    element.setAttribute('marker-start', `url(#${markerId})`);
  }
  if (stroke.tailEnd) {
    const markerId = createArrowMarker(stroke.tailEnd, stroke, defs, 'end');
    element.setAttribute('marker-end', `url(#${markerId})`);
  }
}

/**
 * Creates an SVG marker for an arrow head.
 */
function createArrowMarker(
  arrow: ArrowHead,
  stroke: Stroke,
  defs: SVGDefsElement,
  position: 'start' | 'end'
): string {
  const markerId = `arrow_${++defIdCounter}`;
  const marker = document.createElementNS('http://www.w3.org/2000/svg', 'marker');
  marker.setAttribute('id', markerId);
  marker.setAttribute('markerUnits', 'strokeWidth');
  marker.setAttribute('orient', 'auto');

  // Size multipliers based on width and length
  const widthMult = arrow.width === 'sm' ? 0.6 : arrow.width === 'lg' ? 1.5 : 1;
  const lengthMult = arrow.length === 'sm' ? 0.6 : arrow.length === 'lg' ? 1.5 : 1;

  const baseSize = 4;
  const w = baseSize * widthMult;
  const h = baseSize * lengthMult;

  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
  path.setAttribute('fill', colorToCss(stroke.color));

  let d: string;
  switch (arrow.type) {
    case 'triangle':
      if (position === 'end') {
        marker.setAttribute('viewBox', `0 0 ${h} ${w * 2}`);
        marker.setAttribute('refX', String(h));
        marker.setAttribute('refY', String(w));
        marker.setAttribute('markerWidth', String(h));
        marker.setAttribute('markerHeight', String(w * 2));
        d = `M 0 0 L ${h} ${w} L 0 ${w * 2} Z`;
      } else {
        marker.setAttribute('viewBox', `0 0 ${h} ${w * 2}`);
        marker.setAttribute('refX', '0');
        marker.setAttribute('refY', String(w));
        marker.setAttribute('markerWidth', String(h));
        marker.setAttribute('markerHeight', String(w * 2));
        d = `M ${h} 0 L 0 ${w} L ${h} ${w * 2} Z`;
      }
      break;

    case 'stealth':
      if (position === 'end') {
        marker.setAttribute('viewBox', `0 0 ${h} ${w * 2}`);
        marker.setAttribute('refX', String(h));
        marker.setAttribute('refY', String(w));
        marker.setAttribute('markerWidth', String(h));
        marker.setAttribute('markerHeight', String(w * 2));
        d = `M 0 0 L ${h} ${w} L 0 ${w * 2} L ${h * 0.3} ${w} Z`;
      } else {
        marker.setAttribute('viewBox', `0 0 ${h} ${w * 2}`);
        marker.setAttribute('refX', '0');
        marker.setAttribute('refY', String(w));
        marker.setAttribute('markerWidth', String(h));
        marker.setAttribute('markerHeight', String(w * 2));
        d = `M ${h} 0 L 0 ${w} L ${h} ${w * 2} L ${h * 0.7} ${w} Z`;
      }
      break;

    case 'diamond':
      marker.setAttribute('viewBox', `0 0 ${h * 2} ${w * 2}`);
      marker.setAttribute('refX', String(h));
      marker.setAttribute('refY', String(w));
      marker.setAttribute('markerWidth', String(h * 2));
      marker.setAttribute('markerHeight', String(w * 2));
      d = `M 0 ${w} L ${h} 0 L ${h * 2} ${w} L ${h} ${w * 2} Z`;
      break;

    case 'oval':
      marker.setAttribute('viewBox', `0 0 ${w * 2} ${w * 2}`);
      marker.setAttribute('refX', String(w));
      marker.setAttribute('refY', String(w));
      marker.setAttribute('markerWidth', String(w * 2));
      marker.setAttribute('markerHeight', String(w * 2));
      const ellipse = document.createElementNS('http://www.w3.org/2000/svg', 'ellipse');
      ellipse.setAttribute('cx', String(w));
      ellipse.setAttribute('cy', String(w));
      ellipse.setAttribute('rx', String(w));
      ellipse.setAttribute('ry', String(w));
      ellipse.setAttribute('fill', colorToCss(stroke.color));
      marker.appendChild(ellipse);
      defs.appendChild(marker);
      return markerId;

    case 'arrow':
    default:
      if (position === 'end') {
        marker.setAttribute('viewBox', `0 0 ${h} ${w * 2}`);
        marker.setAttribute('refX', String(h));
        marker.setAttribute('refY', String(w));
        marker.setAttribute('markerWidth', String(h));
        marker.setAttribute('markerHeight', String(w * 2));
        d = `M 0 0 L ${h} ${w} L 0 ${w * 2}`;
        path.setAttribute('fill', 'none');
        path.setAttribute('stroke', colorToCss(stroke.color));
        path.setAttribute('stroke-width', '1');
      } else {
        marker.setAttribute('viewBox', `0 0 ${h} ${w * 2}`);
        marker.setAttribute('refX', '0');
        marker.setAttribute('refY', String(w));
        marker.setAttribute('markerWidth', String(h));
        marker.setAttribute('markerHeight', String(w * 2));
        d = `M ${h} 0 L 0 ${w} L ${h} ${w * 2}`;
        path.setAttribute('fill', 'none');
        path.setAttribute('stroke', colorToCss(stroke.color));
        path.setAttribute('stroke-width', '1');
      }
      break;
  }

  path.setAttribute('d', d);
  marker.appendChild(path);
  defs.appendChild(marker);
  return markerId;
}

/**
 * Creates an SVG filter for shadow effects and returns the filter ID.
 */
function applyShadowFilter(shadow: Shadow, defs: SVGDefsElement): string {
  const filterId = `shadow_${++defIdCounter}`;
  const filter = document.createElementNS('http://www.w3.org/2000/svg', 'filter');
  filter.setAttribute('id', filterId);

  // Extend filter region to accommodate shadow offset and blur
  filter.setAttribute('x', '-50%');
  filter.setAttribute('y', '-50%');
  filter.setAttribute('width', '200%');
  filter.setAttribute('height', '200%');

  // Calculate offset from angle and distance
  const angleRad = (shadow.angle * Math.PI) / 180;
  const dx = Math.cos(angleRad) * shadow.distance;
  const dy = Math.sin(angleRad) * shadow.distance;

  if (shadow.type === 'outer') {
    // Use feDropShadow for outer shadows (simpler and more performant)
    const dropShadow = document.createElementNS('http://www.w3.org/2000/svg', 'feDropShadow');
    dropShadow.setAttribute('dx', String(dx));
    dropShadow.setAttribute('dy', String(dy));
    dropShadow.setAttribute('stdDeviation', String(shadow.blurRadius / 2));
    dropShadow.setAttribute('flood-color', shadow.color.hex);
    dropShadow.setAttribute('flood-opacity', String(shadow.color.alpha));
    filter.appendChild(dropShadow);
  } else {
    // Inner shadow: more complex filter
    // 1. Create inverted shape
    const feComponentTransfer = document.createElementNS('http://www.w3.org/2000/svg', 'feComponentTransfer');
    feComponentTransfer.setAttribute('in', 'SourceAlpha');
    feComponentTransfer.setAttribute('result', 'invert');
    const feFuncA = document.createElementNS('http://www.w3.org/2000/svg', 'feFuncA');
    feFuncA.setAttribute('type', 'table');
    feFuncA.setAttribute('tableValues', '1 0');
    feComponentTransfer.appendChild(feFuncA);
    filter.appendChild(feComponentTransfer);

    // 2. Offset the inverted shape
    const feOffset = document.createElementNS('http://www.w3.org/2000/svg', 'feOffset');
    feOffset.setAttribute('dx', String(dx));
    feOffset.setAttribute('dy', String(dy));
    feOffset.setAttribute('in', 'invert');
    feOffset.setAttribute('result', 'offsetInvert');
    filter.appendChild(feOffset);

    // 3. Blur the offset shape
    const feGaussianBlur = document.createElementNS('http://www.w3.org/2000/svg', 'feGaussianBlur');
    feGaussianBlur.setAttribute('stdDeviation', String(shadow.blurRadius / 2));
    feGaussianBlur.setAttribute('in', 'offsetInvert');
    feGaussianBlur.setAttribute('result', 'blur');
    filter.appendChild(feGaussianBlur);

    // 4. Clip to original shape
    const feComposite = document.createElementNS('http://www.w3.org/2000/svg', 'feComposite');
    feComposite.setAttribute('operator', 'in');
    feComposite.setAttribute('in', 'blur');
    feComposite.setAttribute('in2', 'SourceAlpha');
    feComposite.setAttribute('result', 'innerShadow');
    filter.appendChild(feComposite);

    // 5. Color the shadow
    const feFlood = document.createElementNS('http://www.w3.org/2000/svg', 'feFlood');
    feFlood.setAttribute('flood-color', shadow.color.hex);
    feFlood.setAttribute('flood-opacity', String(shadow.color.alpha));
    feFlood.setAttribute('result', 'color');
    filter.appendChild(feFlood);

    // 6. Apply color to shadow
    const feComposite2 = document.createElementNS('http://www.w3.org/2000/svg', 'feComposite');
    feComposite2.setAttribute('operator', 'in');
    feComposite2.setAttribute('in', 'color');
    feComposite2.setAttribute('in2', 'innerShadow');
    feComposite2.setAttribute('result', 'coloredShadow');
    filter.appendChild(feComposite2);

    // 7. Merge with original
    const feMerge = document.createElementNS('http://www.w3.org/2000/svg', 'feMerge');
    const feMergeNode1 = document.createElementNS('http://www.w3.org/2000/svg', 'feMergeNode');
    feMergeNode1.setAttribute('in', 'SourceGraphic');
    const feMergeNode2 = document.createElementNS('http://www.w3.org/2000/svg', 'feMergeNode');
    feMergeNode2.setAttribute('in', 'coloredShadow');
    feMerge.appendChild(feMergeNode1);
    feMerge.appendChild(feMergeNode2);
    filter.appendChild(feMerge);
  }

  defs.appendChild(filter);
  return filterId;
}

/**
 * Renders a text box element.
 */
function renderTextBox(textEl: TextElement, group: SVGGElement, defs: SVGDefsElement): void {
  const { bounds, fill, stroke, shadow, text } = textEl;

  // If there's a fill or stroke, add a background rect
  if (fill || stroke) {
    const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    rect.setAttribute('width', String(bounds.width));
    rect.setAttribute('height', String(bounds.height));

    if (fill) {
      applyFill(rect, fill, defs);
    } else {
      rect.setAttribute('fill', 'none');
    }

    if (stroke) {
      applyStroke(rect, stroke, defs);
    }

    // Apply shadow filter
    if (shadow) {
      const filterId = applyShadowFilter(shadow, defs);
      rect.setAttribute('filter', `url(#${filterId})`);
    }

    group.appendChild(rect);
  }

  // Render text content
  if (text && text.paragraphs.length > 0) {
    const textFo = renderTextBodyToSvg(text, bounds.width, bounds.height);
    group.appendChild(textFo);
  }
}

/**
 * Renders an image element.
 */
function renderImage(imageEl: ImageElement, group: SVGGElement, defs: SVGDefsElement): void {
  const { bounds, src, altText, shadow, crop } = imageEl;

  const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
  image.setAttribute('href', src);

  if (crop) {
    // Calculate the visible portion of the image
    // crop values are percentages (0-100) of how much to crop from each edge
    const visibleWidth = 100 - crop.left - crop.right;
    const visibleHeight = 100 - crop.top - crop.bottom;

    // Calculate the full image size needed to show the cropped portion at the desired bounds
    const fullWidth = (bounds.width * 100) / visibleWidth;
    const fullHeight = (bounds.height * 100) / visibleHeight;

    // Calculate the offset to position the visible portion correctly
    const offsetX = -(fullWidth * crop.left) / 100;
    const offsetY = -(fullHeight * crop.top) / 100;

    // Create clip path to show only the desired portion
    const clipId = `clip_${++defIdCounter}`;
    const clipPath = document.createElementNS('http://www.w3.org/2000/svg', 'clipPath');
    clipPath.setAttribute('id', clipId);

    const clipRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
    clipRect.setAttribute('x', '0');
    clipRect.setAttribute('y', '0');
    clipRect.setAttribute('width', String(bounds.width));
    clipRect.setAttribute('height', String(bounds.height));
    clipPath.appendChild(clipRect);
    defs.appendChild(clipPath);

    // Set up the image with offset and scaling
    image.setAttribute('x', String(offsetX));
    image.setAttribute('y', String(offsetY));
    image.setAttribute('width', String(fullWidth));
    image.setAttribute('height', String(fullHeight));
    image.setAttribute('preserveAspectRatio', 'none');
    image.setAttribute('clip-path', `url(#${clipId})`);
  } else {
    // No cropping - simple case
    image.setAttribute('width', String(bounds.width));
    image.setAttribute('height', String(bounds.height));
    image.setAttribute('preserveAspectRatio', 'xMidYMid meet');
  }

  if (altText) {
    const title = document.createElementNS('http://www.w3.org/2000/svg', 'title');
    title.textContent = altText;
    image.appendChild(title);
  }

  // Apply shadow filter
  if (shadow) {
    const filterId = applyShadowFilter(shadow, defs);
    image.setAttribute('filter', `url(#${filterId})`);
  }

  group.appendChild(image);
}

/**
 * Renders a group element.
 */
function renderGroup(groupEl: GroupElement, group: SVGGElement, defs: SVGDefsElement): void {
  for (const child of groupEl.children) {
    const childGroup = renderElement(child, defs);
    group.appendChild(childGroup);
  }
}

/**
 * Renders a table element.
 */
function renderTableElement(tableEl: TableElement, group: SVGGElement): void {
  const tableFo = renderTable(tableEl);
  group.appendChild(tableFo);
}

/**
 * Gets the pattern tile size based on pattern type.
 */
function getPatternSize(pattern: PatternType): number {
  // Most patterns use an 8x8 tile
  const smallPatterns = ['pct5', 'pct10', 'smCheck', 'smGrid', 'smConfetti', 'dotGrid'];
  const largePatterns = ['lgCheck', 'lgGrid', 'lgConfetti', 'horzBrick', 'diagBrick', 'plaid'];

  if (smallPatterns.includes(pattern)) return 4;
  if (largePatterns.includes(pattern)) return 16;
  return 8;
}

/**
 * Creates SVG content for a pattern.
 */
function createPatternContent(
  pattern: PatternType,
  foreground: Color,
  background: Color,
  size: number
): SVGGElement {
  const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');

  // Background rectangle
  const bg = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  bg.setAttribute('width', String(size));
  bg.setAttribute('height', String(size));
  bg.setAttribute('fill', background.hex);
  if (background.alpha < 1) {
    bg.setAttribute('fill-opacity', String(background.alpha));
  }
  g.appendChild(bg);

  const fgColor = foreground.hex;
  const fgOpacity = foreground.alpha < 1 ? foreground.alpha : undefined;

  // Create pattern-specific shapes
  switch (pattern) {
    // Percentage patterns (dots)
    case 'pct5':
    case 'pct10':
    case 'pct20':
    case 'pct25':
    case 'pct30':
    case 'pct40':
    case 'pct50':
    case 'pct60':
    case 'pct70':
    case 'pct75':
    case 'pct80':
    case 'pct90': {
      const pct = parseInt(pattern.replace('pct', ''), 10);
      const dotCount = Math.ceil((pct / 100) * (size * size / 4));
      for (let i = 0; i < dotCount; i++) {
        const dot = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        const x = (i * 3) % size;
        const y = Math.floor((i * 3) / size) * 2 % size;
        dot.setAttribute('x', String(x));
        dot.setAttribute('y', String(y));
        dot.setAttribute('width', '1');
        dot.setAttribute('height', '1');
        dot.setAttribute('fill', fgColor);
        if (fgOpacity) dot.setAttribute('fill-opacity', String(fgOpacity));
        g.appendChild(dot);
      }
      break;
    }

    // Horizontal lines
    case 'horz':
    case 'ltHorz':
    case 'dkHorz':
    case 'narHorz':
    case 'wdHorz':
    case 'dashHorz': {
      const lineWidth = pattern.includes('lt') ? 1 : pattern.includes('dk') ? 3 : 2;
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      line.setAttribute('x', '0');
      line.setAttribute('y', String((size - lineWidth) / 2));
      line.setAttribute('width', String(size));
      line.setAttribute('height', String(lineWidth));
      line.setAttribute('fill', fgColor);
      if (fgOpacity) line.setAttribute('fill-opacity', String(fgOpacity));
      g.appendChild(line);
      break;
    }

    // Vertical lines
    case 'vert':
    case 'ltVert':
    case 'dkVert':
    case 'narVert':
    case 'wdVert':
    case 'dashVert': {
      const lineWidth = pattern.includes('lt') ? 1 : pattern.includes('dk') ? 3 : 2;
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      line.setAttribute('x', String((size - lineWidth) / 2));
      line.setAttribute('y', '0');
      line.setAttribute('width', String(lineWidth));
      line.setAttribute('height', String(size));
      line.setAttribute('fill', fgColor);
      if (fgOpacity) line.setAttribute('fill-opacity', String(fgOpacity));
      g.appendChild(line);
      break;
    }

    // Cross/grid patterns
    case 'cross':
    case 'smGrid':
    case 'lgGrid': {
      // Horizontal line
      const hLine = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      hLine.setAttribute('x', '0');
      hLine.setAttribute('y', String(size / 2 - 0.5));
      hLine.setAttribute('width', String(size));
      hLine.setAttribute('height', '1');
      hLine.setAttribute('fill', fgColor);
      if (fgOpacity) hLine.setAttribute('fill-opacity', String(fgOpacity));
      g.appendChild(hLine);
      // Vertical line
      const vLine = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      vLine.setAttribute('x', String(size / 2 - 0.5));
      vLine.setAttribute('y', '0');
      vLine.setAttribute('width', '1');
      vLine.setAttribute('height', String(size));
      vLine.setAttribute('fill', fgColor);
      if (fgOpacity) vLine.setAttribute('fill-opacity', String(fgOpacity));
      g.appendChild(vLine);
      break;
    }

    // Diagonal patterns
    case 'dnDiag':
    case 'ltDnDiag':
    case 'dkDnDiag':
    case 'wdDnDiag': {
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line.setAttribute('x1', '0');
      line.setAttribute('y1', '0');
      line.setAttribute('x2', String(size));
      line.setAttribute('y2', String(size));
      line.setAttribute('stroke', fgColor);
      line.setAttribute('stroke-width', pattern.includes('dk') ? '2' : '1');
      if (fgOpacity) line.setAttribute('stroke-opacity', String(fgOpacity));
      g.appendChild(line);
      break;
    }

    case 'upDiag':
    case 'ltUpDiag':
    case 'dkUpDiag':
    case 'wdUpDiag': {
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line.setAttribute('x1', '0');
      line.setAttribute('y1', String(size));
      line.setAttribute('x2', String(size));
      line.setAttribute('y2', '0');
      line.setAttribute('stroke', fgColor);
      line.setAttribute('stroke-width', pattern.includes('dk') ? '2' : '1');
      if (fgOpacity) line.setAttribute('stroke-opacity', String(fgOpacity));
      g.appendChild(line);
      break;
    }

    // Diagonal cross
    case 'diagCross': {
      const line1 = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line1.setAttribute('x1', '0');
      line1.setAttribute('y1', '0');
      line1.setAttribute('x2', String(size));
      line1.setAttribute('y2', String(size));
      line1.setAttribute('stroke', fgColor);
      line1.setAttribute('stroke-width', '1');
      if (fgOpacity) line1.setAttribute('stroke-opacity', String(fgOpacity));
      g.appendChild(line1);

      const line2 = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line2.setAttribute('x1', '0');
      line2.setAttribute('y1', String(size));
      line2.setAttribute('x2', String(size));
      line2.setAttribute('y2', '0');
      line2.setAttribute('stroke', fgColor);
      line2.setAttribute('stroke-width', '1');
      if (fgOpacity) line2.setAttribute('stroke-opacity', String(fgOpacity));
      g.appendChild(line2);
      break;
    }

    // Checkerboard patterns
    case 'smCheck':
    case 'lgCheck': {
      const halfSize = size / 2;
      const rect1 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect1.setAttribute('x', '0');
      rect1.setAttribute('y', '0');
      rect1.setAttribute('width', String(halfSize));
      rect1.setAttribute('height', String(halfSize));
      rect1.setAttribute('fill', fgColor);
      if (fgOpacity) rect1.setAttribute('fill-opacity', String(fgOpacity));
      g.appendChild(rect1);

      const rect2 = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
      rect2.setAttribute('x', String(halfSize));
      rect2.setAttribute('y', String(halfSize));
      rect2.setAttribute('width', String(halfSize));
      rect2.setAttribute('height', String(halfSize));
      rect2.setAttribute('fill', fgColor);
      if (fgOpacity) rect2.setAttribute('fill-opacity', String(fgOpacity));
      g.appendChild(rect2);
      break;
    }

    // Diamond patterns
    case 'solidDmnd':
    case 'openDmnd':
    case 'dotDmnd': {
      const half = size / 2;
      const diamond = document.createElementNS('http://www.w3.org/2000/svg', 'polygon');
      diamond.setAttribute('points', `${half},0 ${size},${half} ${half},${size} 0,${half}`);
      if (pattern === 'openDmnd') {
        diamond.setAttribute('fill', 'none');
        diamond.setAttribute('stroke', fgColor);
        diamond.setAttribute('stroke-width', '1');
      } else {
        diamond.setAttribute('fill', fgColor);
      }
      if (fgOpacity) {
        diamond.setAttribute(pattern === 'openDmnd' ? 'stroke-opacity' : 'fill-opacity', String(fgOpacity));
      }
      g.appendChild(diamond);
      break;
    }

    // Default: simple dot pattern
    default: {
      const dot = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      dot.setAttribute('cx', String(size / 2));
      dot.setAttribute('cy', String(size / 2));
      dot.setAttribute('r', String(size / 4));
      dot.setAttribute('fill', fgColor);
      if (fgOpacity) dot.setAttribute('fill-opacity', String(fgOpacity));
      g.appendChild(dot);
    }
  }

  return g;
}
