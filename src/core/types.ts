/**
 * Core type definitions for the PPTX Viewer library.
 * These types represent the parsed structure of a PPTX presentation.
 */

// ============================================================================
// Presentation Structure
// ============================================================================

/**
 * Represents a complete parsed PPTX presentation.
 */
export interface Presentation {
  /** Presentation metadata */
  metadata: PresentationMetadata;
  /** Slide dimensions in pixels */
  slideSize: Size;
  /** All slides in presentation order */
  slides: Slide[];
  /** Theme information (colors, fonts) */
  theme: Theme;
}

/**
 * Presentation metadata extracted from docProps.
 */
export interface PresentationMetadata {
  title?: string;
  author?: string;
  createdAt?: string;
  modifiedAt?: string;
}

/**
 * Represents a single slide in the presentation.
 */
export interface Slide {
  /** Slide index (0-based) */
  index: number;
  /** Slide background */
  background?: Background;
  /** All elements on the slide */
  elements: SlideElement[];
  /** Notes for this slide */
  notes?: string;
}

// ============================================================================
// Slide Elements
// ============================================================================

/**
 * Base type for all slide elements.
 */
export interface BaseElement {
  /** Unique identifier */
  id: string;
  /** Element type discriminator */
  type: ElementType;
  /** Position and size */
  bounds: Bounds;
  /** Rotation in degrees */
  rotation?: number;
}

/**
 * Discriminator for element types.
 */
export type ElementType = 'shape' | 'text' | 'image' | 'group';

/**
 * Union type of all possible slide elements.
 */
export type SlideElement = ShapeElement | TextElement | ImageElement | GroupElement;

/**
 * A shape element (rectangle, ellipse, custom path, etc.).
 */
export interface ShapeElement extends BaseElement {
  type: 'shape';
  /** Shape geometry type */
  shapeType: ShapeType;
  /** Fill style */
  fill?: Fill;
  /** Stroke/outline style */
  stroke?: Stroke;
  /** Text content inside the shape */
  text?: TextBody;
  /** Custom geometry path (for freeform shapes) */
  path?: string;
}

/**
 * A standalone text element (text box).
 */
export interface TextElement extends BaseElement {
  type: 'text';
  /** Text content and formatting */
  text: TextBody;
  /** Fill style for text box background */
  fill?: Fill;
  /** Stroke for text box border */
  stroke?: Stroke;
}

/**
 * An image element.
 */
export interface ImageElement extends BaseElement {
  type: 'image';
  /** Image source (data URL or blob URL) */
  src: string;
  /** Original image MIME type */
  mimeType: string;
  /** Alt text for accessibility */
  altText?: string;
}

/**
 * A group of elements.
 */
export interface GroupElement extends BaseElement {
  type: 'group';
  /** Child elements in this group */
  children: SlideElement[];
}

// ============================================================================
// Geometry & Positioning
// ============================================================================

/**
 * Width and height dimensions.
 */
export interface Size {
  width: number;
  height: number;
}

/**
 * Position and size bounds.
 */
export interface Bounds {
  x: number;
  y: number;
  width: number;
  height: number;
}

/**
 * Common shape types supported by PPTX.
 */
export type ShapeType =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'diamond'
  | 'parallelogram'
  | 'trapezoid'
  | 'pentagon'
  | 'hexagon'
  | 'arrow'
  | 'line'
  | 'custom';

// ============================================================================
// Styling
// ============================================================================

/**
 * Fill style for shapes and backgrounds.
 */
export type Fill = SolidFill | GradientFill | ImageFill | NoFill;

export interface SolidFill {
  type: 'solid';
  color: Color;
}

export interface GradientFill {
  type: 'gradient';
  /** Gradient angle in degrees */
  angle: number;
  /** Gradient color stops */
  stops: GradientStop[];
}

export interface GradientStop {
  /** Position from 0 to 1 */
  position: number;
  color: Color;
}

export interface ImageFill {
  type: 'image';
  src: string;
  /** How to fit the image */
  mode: 'stretch' | 'tile' | 'cover' | 'contain';
}

export interface NoFill {
  type: 'none';
}

/**
 * Stroke/outline style.
 */
export interface Stroke {
  color: Color;
  /** Width in pixels */
  width: number;
  /** Dash pattern */
  dashStyle?: 'solid' | 'dash' | 'dot' | 'dashDot';
}

/**
 * Color representation.
 */
export interface Color {
  /** Hex color value (e.g., "#FF0000") */
  hex: string;
  /** Opacity from 0 to 1 */
  alpha: number;
}

/**
 * Slide background.
 */
export interface Background {
  fill: Fill;
}

// ============================================================================
// Text & Typography
// ============================================================================

/**
 * Text body containing paragraphs.
 */
export interface TextBody {
  paragraphs: Paragraph[];
  /** Vertical alignment */
  verticalAlign?: 'top' | 'middle' | 'bottom';
  /** Text padding/margins */
  padding?: {
    top: number;
    right: number;
    bottom: number;
    left: number;
  };
}

/**
 * A paragraph containing text runs.
 */
export interface Paragraph {
  /** Text runs with consistent formatting */
  runs: TextRun[];
  /** Paragraph alignment */
  align?: 'left' | 'center' | 'right' | 'justify';
  /** Line spacing multiplier */
  lineSpacing?: number;
  /** Space before paragraph in pixels */
  spaceBefore?: number;
  /** Space after paragraph in pixels */
  spaceAfter?: number;
  /** Bullet point style */
  bullet?: BulletStyle;
  /** Indentation level (for lists) */
  level?: number;
}

/**
 * A run of text with consistent formatting.
 */
export interface TextRun {
  /** The text content */
  text: string;
  /** Font family */
  fontFamily?: string;
  /** Font size in pixels */
  fontSize?: number;
  /** Font color */
  color?: Color;
  /** Bold */
  bold?: boolean;
  /** Italic */
  italic?: boolean;
  /** Underline */
  underline?: boolean;
  /** Strikethrough */
  strikethrough?: boolean;
  /** Hyperlink URL */
  link?: string;
}

/**
 * Bullet point style.
 */
export interface BulletStyle {
  type: 'bullet' | 'number';
  /** Custom bullet character */
  char?: string;
  /** Starting number for numbered lists */
  startAt?: number;
}

// ============================================================================
// Theme
// ============================================================================

/**
 * Presentation theme with colors and fonts.
 */
export interface Theme {
  /** Theme name */
  name?: string;
  /** Color scheme */
  colors: ThemeColors;
  /** Font scheme */
  fonts: ThemeFonts;
}

/**
 * Theme color palette.
 */
export interface ThemeColors {
  /** Dark 1 (usually background) */
  dark1: string;
  /** Light 1 (usually text) */
  light1: string;
  /** Dark 2 */
  dark2: string;
  /** Light 2 */
  light2: string;
  /** Accent colors 1-6 */
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  /** Hyperlink color */
  hlink: string;
  /** Followed hyperlink color */
  folHlink: string;
}

/**
 * Theme font definitions.
 */
export interface ThemeFonts {
  /** Major font (headings) */
  major: string;
  /** Minor font (body text) */
  minor: string;
}

// ============================================================================
// Viewer Configuration
// ============================================================================

/**
 * Configuration options for the PPTX Viewer.
 */
export interface ViewerOptions {
  /** Initial slide index (0-based) */
  initialSlide?: number;
  /** Enable keyboard navigation */
  keyboardNavigation?: boolean;
  /** Show slide navigation controls */
  showControls?: boolean;
  /** Custom width (defaults to container width) */
  width?: number;
  /** Custom height (defaults to maintain aspect ratio) */
  height?: number;
  /** Callback when slide changes */
  onSlideChange?: (index: number) => void;
  /** Callback when presentation loads */
  onLoad?: (presentation: Presentation) => void;
  /** Callback on error */
  onError?: (error: Error) => void;
}

// ============================================================================
// Events
// ============================================================================

/**
 * Events emitted by the viewer.
 */
export interface ViewerEvents {
  slidechange: number;
  load: Presentation;
  error: Error;
  fullscreenchange: boolean;
}

export type ViewerEventType = keyof ViewerEvents;
