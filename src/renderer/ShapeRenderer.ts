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
      // Default corner radius - roughly 1/8 of smaller dimension
      const radius = Math.min(width, height) / 8;
      rect.setAttribute('rx', String(radius));
      rect.setAttribute('ry', String(radius));
      return rect;
    }

    case 'triangle': {
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      const d = `M ${width / 2} 0 L ${width} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'diamond': {
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      const d = `M ${width / 2} 0 L ${width} ${height / 2} L ${width / 2} ${height} L 0 ${height / 2} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'parallelogram': {
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      const offset = width * 0.2;
      const d = `M ${offset} 0 L ${width} 0 L ${width - offset} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'trapezoid': {
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      const offset = width * 0.15;
      const d = `M ${offset} 0 L ${width - offset} 0 L ${width} ${height} L 0 ${height} Z`;
      path.setAttribute('d', d);
      return path;
    }

    case 'pentagon': {
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      const cx = width / 2;
      const cy = height / 2;
      const r = Math.min(width, height) / 2;
      const points = generatePolygonPoints(cx, cy, r, 5, -90);
      path.setAttribute('d', `M ${points.join(' L ')} Z`);
      return path;
    }

    case 'hexagon': {
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      const cx = width / 2;
      const cy = height / 2;
      const r = Math.min(width, height) / 2;
      const points = generatePolygonPoints(cx, cy, r, 6, -90);
      path.setAttribute('d', `M ${points.join(' L ')} Z`);
      return path;
    }

    case 'arrow': {
      const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
      const arrowWidth = width * 0.6;
      const arrowHead = height * 0.3;
      const d = `
        M 0 ${height * 0.3}
        L ${arrowWidth} ${height * 0.3}
        L ${arrowWidth} 0
        L ${width} ${height / 2}
        L ${arrowWidth} ${height}
        L ${arrowWidth} ${height * 0.7}
        L 0 ${height * 0.7}
        Z
      `;
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
