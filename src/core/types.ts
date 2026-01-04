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
  /** Slide masters indexed by relationship ID */
  slideMasters: Map<string, SlideMaster>;
  /** Slide layouts indexed by relationship ID */
  slideLayouts: Map<string, SlideLayout>;
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
  /** Layout relationship ID (for inheritance) */
  layoutId?: string;
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
  /** Placeholder info (if this element is a placeholder) */
  placeholder?: PlaceholderInfo;
}

/**
 * Discriminator for element types.
 */
export type ElementType = 'shape' | 'text' | 'image' | 'group' | 'table';

/**
 * Union type of all possible slide elements.
 */
export type SlideElement = ShapeElement | TextElement | ImageElement | GroupElement | TableElement;

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
  /** Shadow effect */
  shadow?: Shadow;
  /** Text content inside the shape */
  text?: TextBody;
  /** Custom geometry path (for freeform shapes) */
  path?: string;
  /** Shape adjustment values (e.g., corner radius for rounded rectangles) */
  adjustments?: Map<string, number>;
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
  /** Shadow effect */
  shadow?: Shadow;
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
  /** Shadow effect */
  shadow?: Shadow;
}

/**
 * A group of elements.
 */
export interface GroupElement extends BaseElement {
  type: 'group';
  /** Child elements in this group */
  children: SlideElement[];
}

/**
 * A table element.
 */
export interface TableElement extends BaseElement {
  type: 'table';
  /** Table rows */
  rows: TableRow[];
  /** Column widths in pixels */
  columnWidths: number[];
  /** Table-level styling */
  style?: TableStyle;
}

/**
 * A row in a table.
 */
export interface TableRow {
  /** Row height in pixels */
  height: number;
  /** Cells in this row */
  cells: TableCell[];
}

/**
 * A cell in a table.
 */
export interface TableCell {
  /** Text content */
  text?: TextBody;
  /** Cell fill */
  fill?: Fill;
  /** Cell borders */
  borders?: CellBorders;
  /** Number of columns this cell spans */
  colSpan?: number;
  /** Number of rows this cell spans */
  rowSpan?: number;
  /** Vertical alignment */
  verticalAlign?: 'top' | 'middle' | 'bottom';
}

/**
 * Cell border definitions.
 */
export interface CellBorders {
  top?: Stroke;
  right?: Stroke;
  bottom?: Stroke;
  left?: Stroke;
}

/**
 * Table-level styling.
 */
export interface TableStyle {
  /** First row has special styling */
  firstRow?: boolean;
  /** Last row has special styling */
  lastRow?: boolean;
  /** First column has special styling */
  firstCol?: boolean;
  /** Last column has special styling */
  lastCol?: boolean;
  /** Alternating row bands */
  bandRow?: boolean;
  /** Alternating column bands */
  bandCol?: boolean;
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
  // Basic shapes
  | 'rect'
  | 'roundRect'
  | 'snip1Rect'
  | 'snip2Rect'
  | 'ellipse'
  | 'triangle'
  | 'rtTriangle'
  | 'diamond'
  | 'parallelogram'
  | 'trapezoid'
  | 'pentagon'
  | 'hexagon'
  | 'heptagon'
  | 'octagon'
  | 'decagon'
  | 'dodecagon'
  // Stars
  | 'star4'
  | 'star5'
  | 'star6'
  | 'star8'
  | 'star10'
  | 'star12'
  // Arrows
  | 'arrow'
  | 'leftArrow'
  | 'rightArrow'
  | 'upArrow'
  | 'downArrow'
  | 'leftRightArrow'
  | 'upDownArrow'
  | 'chevron'
  | 'homePlate'
  | 'notchedRightArrow'
  // Callouts
  | 'wedgeRectCallout'
  | 'wedgeRoundRectCallout'
  | 'wedgeEllipseCallout'
  | 'cloudCallout'
  // Block shapes
  | 'cube'
  | 'can'
  | 'lightningBolt'
  | 'heart'
  | 'sun'
  | 'moon'
  | 'cloud'
  | 'arc'
  | 'donut'
  | 'noSmoking'
  | 'blockArc'
  | 'foldedCorner'
  | 'frame'
  | 'halfFrame'
  | 'corner'
  | 'plus'
  | 'cross'
  // Lines and connectors
  | 'line'
  | 'bentConnector3'
  | 'curvedConnector3'
  // Other
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
  /** Gradient type (linear or radial) */
  gradientType: 'linear' | 'radial';
  /** Gradient angle in degrees (for linear gradients) */
  angle: number;
  /** Gradient color stops */
  stops: GradientStop[];
  /** Radial gradient path type (for radial gradients) */
  path?: 'circle' | 'rect';
  /** Fill to rect - defines the rectangle that bounds the gradient (for radial) */
  fillToRect?: {
    left: number;
    top: number;
    right: number;
    bottom: number;
  };
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
  /** Arrow head at start of line */
  headEnd?: ArrowHead;
  /** Arrow head at end of line */
  tailEnd?: ArrowHead;
}

/**
 * Arrow head style for line endings.
 */
export interface ArrowHead {
  /** Arrow type */
  type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
  /** Arrow width: small, medium, large */
  width?: 'sm' | 'med' | 'lg';
  /** Arrow length: small, medium, large */
  length?: 'sm' | 'med' | 'lg';
}

/**
 * Shadow effect.
 */
export interface Shadow {
  /** Shadow type */
  type: 'outer' | 'inner';
  /** Shadow color */
  color: Color;
  /** Blur radius in pixels */
  blurRadius: number;
  /** Distance from shape in pixels */
  distance: number;
  /** Angle in degrees (0 = right, 90 = down) */
  angle: number;
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
  /** Text autofit settings */
  autofit?: TextAutofit;
}

/**
 * Text autofit settings for shrinking text to fit container.
 */
export interface TextAutofit {
  /** Type of autofit */
  type: 'normal' | 'shape' | 'none';
  /** Font scale as decimal (e.g., 0.925 = 92.5%) */
  fontScale?: number;
  /** Line spacing reduction as decimal (e.g., 0.1 = 10% reduction) */
  lineSpacingReduction?: number;
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
  /** Left margin in pixels */
  marginLeft?: number;
  /** First line indent (negative for hanging indent) */
  indent?: number;
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
  /** Baseline offset as percentage (positive = superscript, negative = subscript) */
  baseline?: number;
  /** Hyperlink URL */
  link?: string;
  /** Character spacing in pixels (can be negative for tighter spacing) */
  characterSpacing?: number;
  /** Text capitalization */
  capitalization?: 'none' | 'allCaps' | 'smallCaps';
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
  /** Number format type (e.g., 'arabicPeriod', 'alphaLcParenR') */
  numberType?: string;
  /** Bullet color */
  color?: Color;
  /** Bullet font family */
  font?: string;
  /** Bullet size as percentage of text size (e.g., 100 = same size) */
  sizePercent?: number;
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
// Slide Masters & Layouts
// ============================================================================

/**
 * Placeholder type from OOXML ph@type attribute.
 */
export type PlaceholderType =
  | 'ctrTitle'   // Center title
  | 'title'      // Title
  | 'body'       // Body text
  | 'subTitle'   // Subtitle
  | 'dt'         // Date/time
  | 'ftr'        // Footer
  | 'sldNum'     // Slide number
  | 'hdr'        // Header
  | 'pic'        // Picture
  | 'chart'      // Chart
  | 'tbl'        // Table
  | 'dgm'        // Diagram/SmartArt
  | 'media'      // Media
  | 'clipArt'    // Clip art
  | 'obj';       // Generic object

/**
 * Placeholder information for an element.
 */
export interface PlaceholderInfo {
  /** Placeholder type */
  type: PlaceholderType;
  /** Placeholder index (for matching with layout) */
  idx?: number;
}

/**
 * Color mapping from scheme colors to theme colors.
 * Maps logical color names to theme color slots.
 */
export interface ColorMap {
  /** Background 1 */
  bg1?: string;
  /** Background 2 */
  bg2?: string;
  /** Text 1 */
  tx1?: string;
  /** Text 2 */
  tx2?: string;
  /** Accent colors 1-6 */
  accent1?: string;
  accent2?: string;
  accent3?: string;
  accent4?: string;
  accent5?: string;
  accent6?: string;
  /** Hyperlink */
  hlink?: string;
  /** Followed hyperlink */
  folHlink?: string;
}

/**
 * Represents a slide layout.
 * Layouts define placeholder positions and can override master properties.
 */
export interface SlideLayout {
  /** Layout relationship ID */
  id: string;
  /** Layout name */
  name?: string;
  /** Layout type (e.g., "title", "obj", "twoObj") */
  type?: string;
  /** Parent master relationship ID */
  masterId: string;
  /** Background override (if different from master) */
  background?: Background;
  /** Layout-specific elements and placeholders */
  elements: SlideElement[];
  /** Whether to show master shapes on slides using this layout */
  showMasterShapes: boolean;
  /** Color map overrides */
  colorMap?: ColorMap;
}

/**
 * Represents a slide master.
 * Masters define the base styling for all slides using them.
 */
export interface SlideMaster {
  /** Master relationship ID */
  id: string;
  /** Master name */
  name?: string;
  /** Master background */
  background?: Background;
  /** Master-level elements (logos, decorations visible on all slides) */
  elements: SlideElement[];
  /** Color mapping for this master */
  colorMap: ColorMap;
  /** Associated layout relationship IDs */
  layoutIds: string[];
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
