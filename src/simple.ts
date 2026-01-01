/**
 * Simplified API for common use cases.
 *
 * These functions provide a streamlined interface for users who don't need
 * the full PPTXViewer component and want to handle rendering themselves.
 */

import type { Presentation } from './core/types';
import { extractPPTX, type PPTXArchive } from './core/unzip';
import { parsePPTX } from './parser/PPTXParser';
import { renderSlide } from './renderer/SlideRenderer';

/**
 * Holds a loaded presentation and its archive for cleanup.
 */
export interface LoadedPresentation extends Presentation {
  /**
   * Cleans up resources (blob URLs for images).
   * Call this when you're done with the presentation.
   */
  cleanup: () => void;
}

/**
 * Loads and parses a PPTX file in one step.
 *
 * @param source - PPTX file as File, ArrayBuffer, Uint8Array, or URL string
 * @returns Parsed presentation with cleanup function
 *
 * @example
 * ```typescript
 * import { loadPresentation } from 'pptx-viewer';
 *
 * // From file input
 * const input = document.querySelector('input[type="file"]');
 * input.addEventListener('change', async (e) => {
 *   const file = e.target.files[0];
 *   const presentation = await loadPresentation(file);
 *
 *   console.log(`Loaded ${presentation.slides.length} slides`);
 *
 *   // When done, clean up blob URLs
 *   presentation.cleanup();
 * });
 *
 * // From URL
 * const presentation = await loadPresentation('/path/to/file.pptx');
 * ```
 */
export async function loadPresentation(
  source: File | ArrayBuffer | Uint8Array | string
): Promise<LoadedPresentation> {
  const archive = await extractPPTX(source);
  const presentation = await parsePPTX(archive);

  return {
    ...presentation,
    cleanup: () => archive.cleanup(),
  };
}

/**
 * Renders a slide directly into a container element.
 *
 * @param presentation - Loaded presentation
 * @param slideIndex - 0-based slide index
 * @param container - HTML element to render into
 * @param options - Optional width/height overrides
 *
 * @example
 * ```typescript
 * import { loadPresentation, renderSlideToElement } from 'pptx-viewer';
 *
 * const presentation = await loadPresentation(file);
 * const container = document.getElementById('slide-container');
 *
 * // Render first slide
 * renderSlideToElement(presentation, 0, container);
 *
 * // Render with custom size
 * renderSlideToElement(presentation, 0, container, { width: 800 });
 * ```
 */
export function renderSlideToElement(
  presentation: Presentation,
  slideIndex: number,
  container: HTMLElement,
  options: { width?: number; height?: number } = {}
): void {
  // Validate index
  if (slideIndex < 0 || slideIndex >= presentation.slides.length) {
    throw new Error(`Invalid slide index: ${slideIndex}. Presentation has ${presentation.slides.length} slides.`);
  }

  // Clear container
  container.innerHTML = '';

  // Calculate size if not specified
  const containerRect = container.getBoundingClientRect();
  const slideAspectRatio = presentation.slideSize.width / presentation.slideSize.height;

  let width = options.width;
  let height = options.height;

  if (!width && !height) {
    // Fit to container
    width = containerRect.width;
    height = width / slideAspectRatio;

    if (height > containerRect.height) {
      height = containerRect.height;
      width = height * slideAspectRatio;
    }
  } else if (width && !height) {
    height = width / slideAspectRatio;
  } else if (height && !width) {
    width = height * slideAspectRatio;
  }

  // Render slide
  const slide = presentation.slides[slideIndex];
  const svg = renderSlide(slide, presentation.slideSize, { width, height });

  container.appendChild(svg);
}

/**
 * Renders a slide to a canvas element.
 * Useful for generating thumbnails or when you need raster output.
 *
 * @param presentation - Loaded presentation
 * @param slideIndex - 0-based slide index
 * @param canvas - Canvas element to render into
 * @returns Promise that resolves when rendering is complete
 *
 * @example
 * ```typescript
 * import { loadPresentation, renderSlideToCanvas } from 'pptx-viewer';
 *
 * const presentation = await loadPresentation(file);
 * const canvas = document.getElementById('thumbnail') as HTMLCanvasElement;
 * canvas.width = 320;
 * canvas.height = 240;
 *
 * await renderSlideToCanvas(presentation, 0, canvas);
 * ```
 */
export async function renderSlideToCanvas(
  presentation: Presentation,
  slideIndex: number,
  canvas: HTMLCanvasElement
): Promise<void> {
  // Validate index
  if (slideIndex < 0 || slideIndex >= presentation.slides.length) {
    throw new Error(`Invalid slide index: ${slideIndex}. Presentation has ${presentation.slides.length} slides.`);
  }

  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('Could not get canvas 2D context');
  }

  // Render to SVG first
  const slide = presentation.slides[slideIndex];
  const svg = renderSlide(slide, presentation.slideSize, {
    width: canvas.width,
    height: canvas.height,
  });

  // Convert SVG to data URL
  const svgString = new XMLSerializer().serializeToString(svg);
  const svgBlob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(svgBlob);

  // Draw to canvas
  return new Promise((resolve, reject) => {
    const img = new Image();

    img.onload = () => {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0);
      URL.revokeObjectURL(url);
      resolve();
    };

    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error('Failed to render slide to canvas'));
    };

    img.src = url;
  });
}

/**
 * Gets all slides as thumbnail images.
 *
 * @param presentation - Loaded presentation
 * @param thumbnailWidth - Width of each thumbnail
 * @returns Array of SVG elements
 *
 * @example
 * ```typescript
 * import { loadPresentation, getThumbnails } from 'pptx-viewer';
 *
 * const presentation = await loadPresentation(file);
 * const thumbnails = getThumbnails(presentation, 200);
 *
 * thumbnails.forEach((svg, index) => {
 *   const wrapper = document.createElement('div');
 *   wrapper.appendChild(svg);
 *   wrapper.onclick = () => goToSlide(index);
 *   sidebar.appendChild(wrapper);
 * });
 * ```
 */
export function getThumbnails(
  presentation: Presentation,
  thumbnailWidth: number = 200
): SVGSVGElement[] {
  const aspectRatio = presentation.slideSize.height / presentation.slideSize.width;
  const thumbnailHeight = thumbnailWidth * aspectRatio;

  return presentation.slides.map((slide) =>
    renderSlide(slide, presentation.slideSize, {
      width: thumbnailWidth,
      height: thumbnailHeight,
    })
  );
}
