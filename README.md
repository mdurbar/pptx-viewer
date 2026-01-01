# PPTX Viewer

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

- Basic shapes (rectangles, ellipses, triangles, etc.)
- Text with formatting (font, size, color, bold, italic, underline)
- Images (PNG, JPEG, GIF, etc.)
- Solid and gradient fills
- Stroke/outline styles
- Grouped shapes
- Theme colors and fonts
- Slide backgrounds

### Partially Supported

- Charts (rendered as fallback images when available)
- SmartArt (rendered as fallback images when available)
- Complex shapes (simplified to basic approximations)

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
```

## License

MIT
