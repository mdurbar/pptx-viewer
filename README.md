# PPTX Viewer

[![npm version](https://img.shields.io/npm/v/pptx-viewer.svg)](https://www.npmjs.com/package/pptx-viewer)
[![bundle size](https://img.shields.io/bundlephobia/minzip/pptx-viewer)](https://bundlephobia.com/package/pptx-viewer)
[![license](https://img.shields.io/npm/l/pptx-viewer.svg)](https://github.com/mdurbar/pptx-viewer/blob/main/LICENSE)

A lightweight, dependency-minimal library for viewing PowerPoint (PPTX) files in the browser.

## Features

- **Lightweight** - Only one dependency (fflate for ZIP decompression, ~8KB gzipped)
- **No server required** - Runs entirely in the browser
- **TypeScript** - Full type definitions included
- **Flexible API** - Use the full viewer or just the parser/renderer
- **Presentation mode** - Fullscreen support with keyboard navigation
- **Good fidelity** - Renders shapes, text, images, and basic styling

## Installation

```bash
npm install pptx-viewer
```

## Quick Start

### Option 1: Full Viewer (with controls)

```typescript
import { PPTXViewer } from 'pptx-viewer';

const viewer = new PPTXViewer('#container');
await viewer.load(file);

viewer.next();
viewer.previous();
viewer.goToSlide(2);
viewer.enterFullscreen();
```

### Option 2: Custom Rendering (bring your own UI)

```typescript
import { loadPresentation, renderSlideToElement } from 'pptx-viewer';

// Load the presentation
const presentation = await loadPresentation(file);

// Render a specific slide to your container
const container = document.getElementById('my-slide');
renderSlideToElement(presentation, 0, container);

// Navigate by rendering different slides
function goToSlide(index: number) {
  renderSlideToElement(presentation, index, container);
}

// Clean up when done
presentation.cleanup();
```

### Option 3: Generate Thumbnails

```typescript
import { loadPresentation, getThumbnails } from 'pptx-viewer';

const presentation = await loadPresentation(file);
const thumbnails = getThumbnails(presentation, 200); // 200px wide

thumbnails.forEach((svg, index) => {
  svg.onclick = () => goToSlide(index);
  sidebar.appendChild(svg);
});
```

## Usage Examples

### From File Input

```html
<input type="file" id="file-input" accept=".pptx">
<div id="viewer" style="width: 100%; height: 600px;"></div>

<script type="module">
  import { PPTXViewer } from 'pptx-viewer';

  const viewer = new PPTXViewer('#viewer');

  document.getElementById('file-input').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    await viewer.load(file);
  });
</script>
```

### From URL

```typescript
const viewer = new PPTXViewer('#viewer');
await viewer.load('/presentations/demo.pptx');
```

### With Options

```typescript
const viewer = new PPTXViewer('#viewer', {
  initialSlide: 0,           // Start at first slide
  showControls: true,        // Show navigation controls
  keyboardNavigation: true,  // Enable keyboard shortcuts
  onSlideChange: (index) => {
    console.log(`Now on slide ${index + 1}`);
  },
  onLoad: (presentation) => {
    console.log(`Loaded ${presentation.slides.length} slides`);
  },
  onError: (error) => {
    console.error('Failed to load:', error);
  },
});
```

## API Reference

### PPTXViewer

#### Constructor

```typescript
new PPTXViewer(container: string | HTMLElement, options?: ViewerOptions)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `container` | `string \| HTMLElement` | CSS selector or element to render into |
| `options` | `ViewerOptions` | Optional configuration |

#### Methods

| Method | Returns | Description |
|--------|---------|-------------|
| `load(source)` | `Promise<void>` | Load a PPTX file (File, ArrayBuffer, or URL) |
| `next()` | `number` | Go to next slide, returns new index or -1 |
| `previous()` | `number` | Go to previous slide, returns new index or -1 |
| `goToSlide(index)` | `boolean` | Go to specific slide (0-based index) |
| `getCurrentSlide()` | `number` | Get current slide index |
| `getSlideCount()` | `number` | Get total number of slides |
| `getPresentation()` | `Presentation \| null` | Get parsed presentation data |
| `enterFullscreen()` | `Promise<void>` | Enter fullscreen mode |
| `exitFullscreen()` | `Promise<void>` | Exit fullscreen mode |
| `toggleFullscreen()` | `Promise<void>` | Toggle fullscreen mode |
| `on(event, listener)` | `void` | Add event listener |
| `off(event, listener)` | `void` | Remove event listener |
| `destroy()` | `void` | Clean up and release resources |

#### Events

```typescript
viewer.on('slidechange', (index: number) => { });
viewer.on('load', (presentation: Presentation) => { });
viewer.on('error', (error: Error) => { });
viewer.on('fullscreenchange', (isFullscreen: boolean) => { });
```

### ViewerOptions

```typescript
interface ViewerOptions {
  initialSlide?: number;           // Default: 0
  keyboardNavigation?: boolean;    // Default: true
  showControls?: boolean;          // Default: true
  width?: number;                  // Custom width
  height?: number;                 // Custom height
  onSlideChange?: (index: number) => void;
  onLoad?: (presentation: Presentation) => void;
  onError?: (error: Error) => void;
}
```

## Keyboard Shortcuts

| Key | Action |
|-----|--------|
| `→` `↓` `Space` `PageDown` | Next slide |
| `←` `↑` `PageUp` | Previous slide |
| `Home` | First slide |
| `End` | Last slide |
| `F` | Toggle fullscreen |
| `Escape` | Exit fullscreen |

## Advanced Usage

### Custom Rendering

For more control, you can use the parser and renderer separately:

```typescript
import { extractPPTX, parsePPTX, renderSlide } from 'pptx-viewer';

// Extract the PPTX archive
const archive = await extractPPTX(file);

// Parse into a Presentation object
const presentation = await parsePPTX(archive);

// Render a specific slide
const svg = renderSlide(
  presentation.slides[0],
  presentation.slideSize,
  { width: 800 }
);

document.body.appendChild(svg);

// Clean up when done
archive.cleanup();
```

### Accessing Presentation Data

```typescript
const viewer = new PPTXViewer('#viewer');
await viewer.load(file);

const presentation = viewer.getPresentation();

// Access metadata
console.log(presentation.metadata.title);
console.log(presentation.metadata.author);

// Access theme
console.log(presentation.theme.colors);
console.log(presentation.theme.fonts);

// Access slides
presentation.slides.forEach((slide, index) => {
  console.log(`Slide ${index + 1}: ${slide.elements.length} elements`);
});
```

## Supported Features

### Fully Supported

- **Slide masters & layouts** - Background and element inheritance from masters and layouts
- **Tables** - Full table rendering with cell styling, borders, and text formatting
- **50+ shape types** - Rectangles, ellipses, triangles, stars (4-12 points), arrows (all directions), callouts, hearts, clouds, and more
- **Text with formatting** - Font family, size, color, bold, italic, underline, strikethrough, subscript, superscript, highlight, outline
- **Text effects** - Glow and reflection effects on text
- **Element opacity** - Transparency support for images and other elements
- **Character spacing & capitalization** - Letter-spacing, all caps, and small caps
- **Text autofit** - Automatic font scaling to fit text in containers
- **Bullet and numbered lists** - Multiple formats (arabic, alpha, roman) with proper indentation
- **Shadow effects** - Outer and inner shadows on shapes, images, and text boxes
- **Line arrow heads** - Triangle, stealth, diamond, oval, and arrow markers
- **Connector lines** - Straight, bent (elbow), and curved connectors with flip support
- **Shape adjustments** - Custom corner radius, snip sizes, star point depths, and more
- **Hyperlinks** - Clickable links in text
- **Images** - PNG, JPEG, GIF with cropping support
- **Solid, gradient, and pattern fills** - Linear/radial gradients, 40+ pattern types, theme and preset colors
- **Stroke/outline styles** - Color, width, dash patterns
- **Grouped shapes** - Nested groups supported
- **Theme colors and fonts** - Full theme support
- **Slide backgrounds** - Solid and gradient
- **Charts** - Native SVG rendering for bar, column, stacked column, pie, doughnut, line, area, and scatter charts
- **SmartArt diagrams** - Native rendering using pre-computed DrawingML shapes, with fallback to embedded images

### Not Supported

- Animations and transitions
- Videos and audio
- 3D effects
- Embedded fonts
- Interactive elements

## Browser Support

- Chrome 80+
- Firefox 75+
- Safari 13+
- Edge 80+

Requires support for:
- ES2020
- SVG foreignObject
- Fullscreen API

## Development

```bash
# Install dependencies
npm install

# Start development server
npm run dev

# Build for production
npm run build

# Type check
npm run typecheck

# Run tests
npm test

# Run tests once (CI mode)
npm run test:run

# Run tests with coverage
npm run test:coverage
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Areas for Contribution

- Performance optimizations
- Animation and transition support
- Video and audio playback

## License

MIT - see [LICENSE](LICENSE) for details.
