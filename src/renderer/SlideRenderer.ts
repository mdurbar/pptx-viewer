/**
 * Renderer for complete slides.
 *
 * Orchestrates rendering of backgrounds and all slide elements.
 */

import type { Slide, Size, Fill, SlideLayout, SlideMaster, Background, SlideElement } from '../core/types';
import { colorToCss } from '../utils/color';
import { renderElement } from './ShapeRenderer';

/**
 * Options for rendering a slide.
 */
export interface SlideRenderOptions {
  /** Target width (will scale proportionally) */
  width?: number;
  /** Target height (will scale proportionally) */
  height?: number;
}

/**
 * Renders a slide to an SVG element.
 *
 * @param slide - The slide to render
 * @param slideSize - Original slide dimensions
 * @param options - Rendering options
 * @returns SVG element containing the rendered slide
 */
export function renderSlide(
  slide: Slide,
  slideSize: Size,
  options: SlideRenderOptions = {}
): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');

  // Set viewBox to original slide size for proper scaling
  svg.setAttribute('viewBox', `0 0 ${slideSize.width} ${slideSize.height}`);

  // Set display size
  if (options.width) {
    svg.setAttribute('width', String(options.width));
  }
  if (options.height) {
    svg.setAttribute('height', String(options.height));
  }

  // Preserve aspect ratio
  svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');

  // Create defs for gradients/patterns
  const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
  svg.appendChild(defs);

  // Render background
  const bgRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  bgRect.setAttribute('width', String(slideSize.width));
  bgRect.setAttribute('height', String(slideSize.height));

  try {
    if (slide.background?.fill) {
      applyBackgroundFill(bgRect, slide.background.fill, defs);
    } else {
      // Default white background
      bgRect.setAttribute('fill', '#FFFFFF');
    }
  } catch (error) {
    // Fall back to white background on error
    console.warn('Failed to render background:', error);
    bgRect.setAttribute('fill', '#FFFFFF');
  }

  svg.appendChild(bgRect);

  // Render elements with error recovery
  for (const element of slide.elements) {
    try {
      const elementGroup = renderElement(element, defs);
      svg.appendChild(elementGroup);
    } catch (error) {
      // Log but don't fail on individual element render errors
      console.warn(`Failed to render element ${element.id}:`, error);
    }
  }

  return svg;
}

/**
 * Renders a slide with full master/layout inheritance.
 *
 * This function implements proper layering:
 * 1. Background (resolved from slide → layout → master)
 * 2. Master shapes (if showMasterShapes is true)
 * 3. Layout shapes (non-placeholder elements)
 * 4. Slide content
 *
 * @param slide - The slide to render
 * @param slideSize - Original slide dimensions
 * @param layout - Optional slide layout
 * @param master - Optional slide master
 * @param options - Rendering options
 * @returns SVG element containing the rendered slide
 */
export function renderSlideWithInheritance(
  slide: Slide,
  slideSize: Size,
  layout?: SlideLayout,
  master?: SlideMaster,
  options: SlideRenderOptions = {}
): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');

  // Set viewBox to original slide size for proper scaling
  svg.setAttribute('viewBox', `0 0 ${slideSize.width} ${slideSize.height}`);

  // Set display size
  if (options.width) {
    svg.setAttribute('width', String(options.width));
  }
  if (options.height) {
    svg.setAttribute('height', String(options.height));
  }

  // Preserve aspect ratio
  svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');

  // Create defs for gradients/patterns
  const defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
  svg.appendChild(defs);

  // LAYER 1: Background (resolved from inheritance chain)
  const resolvedBackground = resolveBackground(slide, layout, master);
  const bgRect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  bgRect.setAttribute('width', String(slideSize.width));
  bgRect.setAttribute('height', String(slideSize.height));

  try {
    if (resolvedBackground?.fill) {
      applyBackgroundFill(bgRect, resolvedBackground.fill, defs);
    } else {
      bgRect.setAttribute('fill', '#FFFFFF');
    }
  } catch (error) {
    console.warn('Failed to render background:', error);
    bgRect.setAttribute('fill', '#FFFFFF');
  }
  svg.appendChild(bgRect);

  // LAYER 2: Master shapes (if layout allows)
  if (master && (layout?.showMasterShapes !== false)) {
    const masterGroup = createLayerGroup('master-shapes');
    renderElementsToGroup(master.elements, masterGroup, defs, true);
    svg.appendChild(masterGroup);
  }

  // LAYER 3: Layout shapes (non-placeholder elements only)
  if (layout) {
    const layoutGroup = createLayerGroup('layout-shapes');
    renderElementsToGroup(layout.elements, layoutGroup, defs, true);
    svg.appendChild(layoutGroup);
  }

  // LAYER 4: Slide content
  const contentGroup = createLayerGroup('slide-content');
  renderElementsToGroup(slide.elements, contentGroup, defs, false);
  svg.appendChild(contentGroup);

  return svg;
}

/**
 * Resolves background from inheritance chain.
 * Priority: Slide > Layout > Master
 */
function resolveBackground(
  slide: Slide,
  layout?: SlideLayout,
  master?: SlideMaster
): Background | undefined {
  // Slide's own background takes priority
  if (slide.background) {
    return slide.background;
  }
  // Then layout background
  if (layout?.background) {
    return layout.background;
  }
  // Finally master background
  if (master?.background) {
    return master.background;
  }
  return undefined;
}

/**
 * Creates a group element for layering.
 */
function createLayerGroup(name: string): SVGGElement {
  const group = document.createElementNS('http://www.w3.org/2000/svg', 'g');
  group.setAttribute('data-layer', name);
  return group;
}

/**
 * Renders elements to a group.
 * @param skipPlaceholders - If true, skip elements that are placeholders (for master/layout layers)
 */
function renderElementsToGroup(
  elements: SlideElement[],
  group: SVGGElement,
  defs: SVGDefsElement,
  skipPlaceholders: boolean
): void {
  for (const element of elements) {
    // Skip placeholder shapes in master/layout layers
    // They define positions but actual content comes from slides
    if (skipPlaceholders && element.placeholder) {
      continue;
    }

    try {
      const elementGroup = renderElement(element, defs);
      group.appendChild(elementGroup);
    } catch (error) {
      console.warn(`Failed to render element ${element.id}:`, error);
    }
  }
}

/**
 * Applies a fill to the background rect.
 */
function applyBackgroundFill(rect: SVGRectElement, fill: Fill, defs: SVGDefsElement): void {
  switch (fill.type) {
    case 'solid':
      rect.setAttribute('fill', colorToCss(fill.color));
      if (fill.color.alpha < 1) {
        rect.setAttribute('fill-opacity', String(fill.color.alpha));
      }
      break;

    case 'gradient': {
      const gradientId = `bg_gradient_${Date.now()}`;

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

        const angle = fill.angle || 0;
        const radians = (angle * Math.PI) / 180;
        gradient.setAttribute('x1', String(50 - 50 * Math.cos(radians)) + '%');
        gradient.setAttribute('y1', String(50 - 50 * Math.sin(radians)) + '%');
        gradient.setAttribute('x2', String(50 + 50 * Math.cos(radians)) + '%');
        gradient.setAttribute('y2', String(50 + 50 * Math.sin(radians)) + '%');

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

      rect.setAttribute('fill', `url(#${gradientId})`);
      break;
    }

    case 'image': {
      const patternId = `bg_pattern_${Date.now()}`;
      const pattern = document.createElementNS('http://www.w3.org/2000/svg', 'pattern');
      pattern.setAttribute('id', patternId);
      pattern.setAttribute('patternUnits', 'objectBoundingBox');
      pattern.setAttribute('width', '1');
      pattern.setAttribute('height', '1');

      const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
      image.setAttribute('href', fill.src);
      image.setAttribute('width', '100%');
      image.setAttribute('height', '100%');

      // Apply fill mode
      switch (fill.mode) {
        case 'contain':
          image.setAttribute('preserveAspectRatio', 'xMidYMid meet');
          break;
        case 'cover':
          image.setAttribute('preserveAspectRatio', 'xMidYMid slice');
          break;
        case 'stretch':
          image.setAttribute('preserveAspectRatio', 'none');
          break;
        case 'tile':
          // For tiling, we'd need to know the image dimensions
          // For now, treat as cover
          image.setAttribute('preserveAspectRatio', 'xMidYMid slice');
          break;
      }

      pattern.appendChild(image);
      defs.appendChild(pattern);
      rect.setAttribute('fill', `url(#${patternId})`);
      break;
    }

    case 'none':
      rect.setAttribute('fill', 'none');
      break;
  }
}

/**
 * Renders a slide thumbnail (smaller, simplified).
 *
 * @param slide - The slide to render
 * @param slideSize - Original slide dimensions
 * @param thumbnailWidth - Target thumbnail width
 * @returns SVG element
 */
export function renderSlideThumbnail(
  slide: Slide,
  slideSize: Size,
  thumbnailWidth: number
): SVGSVGElement {
  const aspectRatio = slideSize.height / slideSize.width;
  const thumbnailHeight = thumbnailWidth * aspectRatio;

  return renderSlide(slide, slideSize, {
    width: thumbnailWidth,
    height: thumbnailHeight,
  });
}

/**
 * Creates an empty slide placeholder.
 *
 * @param slideSize - Slide dimensions
 * @param message - Optional message to display
 * @returns SVG element
 */
export function createEmptySlide(slideSize: Size, message?: string): SVGSVGElement {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('viewBox', `0 0 ${slideSize.width} ${slideSize.height}`);
  svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');

  // White background
  const bg = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  bg.setAttribute('width', String(slideSize.width));
  bg.setAttribute('height', String(slideSize.height));
  bg.setAttribute('fill', '#FFFFFF');
  svg.appendChild(bg);

  // Optional message
  if (message) {
    const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    text.setAttribute('x', String(slideSize.width / 2));
    text.setAttribute('y', String(slideSize.height / 2));
    text.setAttribute('text-anchor', 'middle');
    text.setAttribute('dominant-baseline', 'middle');
    text.setAttribute('fill', '#666666');
    text.setAttribute('font-family', 'sans-serif');
    text.setAttribute('font-size', '24');
    text.textContent = message;
    svg.appendChild(text);
  }

  return svg;
}
