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
  Fill,
  Stroke,
  ShapeType,
  Bounds,
} from '../core/types';
import { colorToCss } from '../utils/color';
import { renderTextBodyToSvg } from './TextRenderer';

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

  switch (element.type) {
    case 'shape':
      renderShape(element, group, defs);
      break;
    case 'text':
      renderTextBox(element, group, defs);
      break;
    case 'image':
      renderImage(element, group);
      break;
    case 'group':
      renderGroup(element, group, defs);
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
  const { bounds, shapeType, fill, stroke, text } = shape;

  // Check if shape has any visible fill or stroke
  const hasVisibleFill = fill && fill.type !== 'none';
  const hasVisibleStroke = stroke && stroke.width > 0;

  // Only create shape element if it has visible fill or stroke
  if (hasVisibleFill || hasVisibleStroke) {
    const shapeEl = createShapeElement(shapeType, bounds.width, bounds.height);

    // Apply fill
    if (hasVisibleFill) {
      applyFill(shapeEl, fill, defs);
    } else {
      shapeEl.setAttribute('fill', 'none');
    }

    // Apply stroke
    if (hasVisibleStroke) {
      applyStroke(shapeEl, stroke);
    } else {
      shapeEl.setAttribute('stroke', 'none');
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
function createShapeElement(shapeType: ShapeType, width: number, height: number): SVGElement {
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
      const radius = Math.min(width, height) / 8;
      rect.setAttribute('rx', String(radius));
      rect.setAttribute('ry', String(radius));
      return rect;
    }

    case 'snip1Rect': {
      const snip = Math.min(width, height) * 0.15;
      d = `M 0 0 L ${width - snip} 0 L ${width} ${snip} L ${width} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'snip2Rect': {
      const snip = Math.min(width, height) * 0.15;
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
      const innerR = outerR * 0.4;
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
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line.setAttribute('x1', '0');
      line.setAttribute('y1', '0');
      line.setAttribute('x2', String(width));
      line.setAttribute('y2', String(height));
      return line;
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
      element.setAttribute('fill', `url(#${gradientId})`);
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
function applyStroke(element: SVGElement, stroke: Stroke): void {
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
}

/**
 * Renders a text box element.
 */
function renderTextBox(textEl: TextElement, group: SVGGElement, defs: SVGDefsElement): void {
  const { bounds, fill, stroke, text } = textEl;

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
      applyStroke(rect, stroke);
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
function renderImage(imageEl: ImageElement, group: SVGGElement): void {
  const { bounds, src, altText } = imageEl;

  const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
  image.setAttribute('href', src);
  image.setAttribute('width', String(bounds.width));
  image.setAttribute('height', String(bounds.height));
  image.setAttribute('preserveAspectRatio', 'xMidYMid meet');

  if (altText) {
    const title = document.createElementNS('http://www.w3.org/2000/svg', 'title');
    title.textContent = altText;
    image.appendChild(title);
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
