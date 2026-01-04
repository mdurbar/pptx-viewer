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
  ImageCrop,
  GroupElement,
  TableElement,
  TableRow,
  TableCell,
  TableStyle,
  CellBorders,
  ElementType,

  // Styling
  Fill,
  SolidFill,
  GradientFill,
  PatternFill,
  PatternType,
  ImageFill,
  NoFill,
  Stroke,
  ArrowHead,
  Shadow,
  Color,
  GradientStop,
  Background,
  ShapeType,

  // Text
  TextBody,
  TextAutofit,
  Paragraph,
  TextRun,
  BulletStyle,
  TextGlow,
  TextReflection,
  TextOutline,

  // Theme
  Theme,
  ThemeColors,
  ThemeFonts,

  // Slide Masters & Layouts
  SlideMaster,
  SlideLayout,
  PlaceholderType,
  PlaceholderInfo,
  ColorMap,

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
export { renderSlide, renderSlideWithInheritance, renderSlideThumbnail, createEmptySlide } from './renderer/SlideRenderer';
export { renderElement } from './renderer/ShapeRenderer';

// =============================================================================
// Errors
// =============================================================================

export {
  PPTXError,
  InvalidFileError,
  MissingFileError,
  XMLParseError,
  FetchError,
  RenderError,
  isPPTXError,
} from './core/errors';
