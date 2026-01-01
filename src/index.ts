/**
 * PPTX Viewer - A lightweight PowerPoint viewer for the browser.
 *
 * @packageDocumentation
 *
 * ## Quick Start - Full Viewer
 * ```typescript
 * import { PPTXViewer } from 'pptx-viewer';
 *
 * const viewer = new PPTXViewer('#container');
 * await viewer.load(file);
 * ```
 *
 * ## Custom Rendering - Just the Parser & Renderer
 * ```typescript
 * import { loadPresentation, renderSlideToElement } from 'pptx-viewer';
 *
 * const presentation = await loadPresentation(file);
 * renderSlideToElement(presentation, 0, document.getElementById('slide'));
 * ```
 */

// =============================================================================
// High-Level API (Most users start here)
// =============================================================================

// Full viewer with controls
export { PPTXViewer } from './PPTXViewer';

// Simple functions for custom implementations
export {
  loadPresentation,
  renderSlideToElement,
  renderSlideToCanvas,
  getThumbnails,
  type LoadedPresentation,
} from './simple';

// =============================================================================
// Types
// =============================================================================

export type {
  // Presentation structure
  Presentation,
  PresentationMetadata,
  Slide,
  SlideElement,
  Size,
  Bounds,

  // Elements
  ShapeElement,
  TextElement,
  ImageElement,
  GroupElement,
  ElementType,

  // Styling
  Fill,
  SolidFill,
  GradientFill,
  ImageFill,
  NoFill,
  Stroke,
  Color,
  GradientStop,
  Background,
  ShapeType,

  // Text
  TextBody,
  Paragraph,
  TextRun,
  BulletStyle,

  // Theme
  Theme,
  ThemeColors,
  ThemeFonts,

  // Viewer
  ViewerOptions,
  ViewerEvents,
  ViewerEventType,
} from './core/types';

// =============================================================================
// Low-Level API (For advanced customization)
// =============================================================================

// Archive extraction
export { extractPPTX, type PPTXArchive } from './core/unzip';

// Parsing
export { parsePPTX, isValidPPTX, getSlideCount } from './parser/PPTXParser';

// Rendering
export { renderSlide, renderSlideThumbnail, createEmptySlide } from './renderer/SlideRenderer';
export { renderElement } from './renderer/ShapeRenderer';
