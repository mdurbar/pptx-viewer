/**
 * Main PPTX Viewer class.
 *
 * Provides a simple API for loading and viewing PPTX presentations.
 *
 * @example
 * ```typescript
 * import { PPTXViewer } from 'pptx-viewer';
 *
 * const viewer = new PPTXViewer('#container');
 * await viewer.load(file);
 *
 * viewer.next();
 * viewer.previous();
 * viewer.goToSlide(3);
 * ```
 */

import type {
  Presentation,
  Slide,
  ViewerOptions,
  ViewerEvents,
  ViewerEventType,
} from './core/types';
import { extractPPTX, type PPTXArchive } from './core/unzip';
import { parsePPTX } from './parser/PPTXParser';
import { renderSlide, createEmptySlide } from './renderer/SlideRenderer';

/**
 * Event listener function type.
 */
type EventListener<T> = (data: T) => void;

/**
 * PPTX Viewer - A lightweight PowerPoint viewer for the browser.
 *
 * @example
 * ```typescript
 * // Basic usage
 * const viewer = new PPTXViewer('#container');
 * await viewer.load(file);
 *
 * // With options
 * const viewer = new PPTXViewer('#container', {
 *   initialSlide: 0,
 *   keyboardNavigation: true,
 *   showControls: true,
 *   onSlideChange: (index) => console.log(`Slide ${index + 1}`),
 * });
 * ```
 */
export class PPTXViewer {
  private container: HTMLElement;
  private options: Required<ViewerOptions>;
  private presentation: Presentation | null = null;
  private archive: PPTXArchive | null = null;
  private currentSlideIndex: number = 0;
  private slideContainer: HTMLElement | null = null;
  private controlsContainer: HTMLElement | null = null;
  private listeners: Map<ViewerEventType, Set<EventListener<any>>> = new Map();
  private isFullscreen: boolean = false;

  /**
   * Creates a new PPTX Viewer instance.
   *
   * @param container - CSS selector or HTMLElement to render the viewer into
   * @param options - Viewer configuration options
   *
   * @throws Error if the container is not found
   */
  constructor(container: string | HTMLElement, options: ViewerOptions = {}) {
    // Resolve container
    if (typeof container === 'string') {
      const el = document.querySelector(container);
      if (!el || !(el instanceof HTMLElement)) {
        throw new Error(`Container not found: ${container}`);
      }
      this.container = el;
    } else {
      this.container = container;
    }

    // Merge options with defaults
    this.options = {
      initialSlide: options.initialSlide ?? 0,
      keyboardNavigation: options.keyboardNavigation ?? true,
      showControls: options.showControls ?? true,
      width: options.width ?? undefined as any,
      height: options.height ?? undefined as any,
      onSlideChange: options.onSlideChange ?? (() => {}),
      onLoad: options.onLoad ?? (() => {}),
      onError: options.onError ?? (() => {}),
    };

    // Set up container
    this.setupContainer();

    // Set up keyboard navigation
    if (this.options.keyboardNavigation) {
      this.setupKeyboardNavigation();
    }
  }

  /**
   * Loads a PPTX file into the viewer.
   *
   * @param source - The PPTX file to load. Accepts:
   *   - `File` object from file input
   *   - `ArrayBuffer` of PPTX data
   *   - `string` URL to fetch PPTX from
   *
   * @returns Promise that resolves when slides are ready
   *
   * @example
   * ```typescript
   * // From file input
   * const file = inputElement.files[0];
   * await viewer.load(file);
   *
   * // From URL
   * await viewer.load('/presentations/demo.pptx');
   * ```
   */
  async load(source: File | ArrayBuffer | string): Promise<void> {
    try {
      // Clean up previous presentation
      if (this.archive) {
        this.archive.cleanup();
      }

      // Extract and parse
      this.archive = await extractPPTX(source);
      this.presentation = await parsePPTX(this.archive);

      // Go to initial slide
      this.currentSlideIndex = Math.min(
        this.options.initialSlide,
        this.presentation.slides.length - 1
      );

      // Render current slide
      this.renderCurrentSlide();

      // Update controls
      this.updateControls();

      // Emit events
      this.emit('load', this.presentation);
      this.options.onLoad(this.presentation);
    } catch (error) {
      const err = error instanceof Error ? error : new Error(String(error));
      this.emit('error', err);
      this.options.onError(err);
      throw err;
    }
  }

  /**
   * Advances to the next slide.
   *
   * @returns The new slide index, or -1 if already at the last slide
   */
  next(): number {
    if (!this.presentation) return -1;

    if (this.currentSlideIndex < this.presentation.slides.length - 1) {
      this.currentSlideIndex++;
      this.renderCurrentSlide();
      this.updateControls();
      this.emitSlideChange();
      return this.currentSlideIndex;
    }

    return -1;
  }

  /**
   * Goes to the previous slide.
   *
   * @returns The new slide index, or -1 if already at the first slide
   */
  previous(): number {
    if (!this.presentation) return -1;

    if (this.currentSlideIndex > 0) {
      this.currentSlideIndex--;
      this.renderCurrentSlide();
      this.updateControls();
      this.emitSlideChange();
      return this.currentSlideIndex;
    }

    return -1;
  }

  /**
   * Goes to a specific slide.
   *
   * @param index - The 0-based slide index
   * @returns True if navigation was successful
   */
  goToSlide(index: number): boolean {
    if (!this.presentation) return false;

    if (index >= 0 && index < this.presentation.slides.length) {
      this.currentSlideIndex = index;
      this.renderCurrentSlide();
      this.updateControls();
      this.emitSlideChange();
      return true;
    }

    return false;
  }

  /**
   * Enters fullscreen presentation mode.
   */
  async enterFullscreen(): Promise<void> {
    try {
      await this.container.requestFullscreen();
      this.isFullscreen = true;
      this.container.classList.add('pptx-fullscreen');
      this.emit('fullscreenchange', true);
    } catch (error) {
      console.warn('Fullscreen not supported:', error);
    }
  }

  /**
   * Exits fullscreen mode.
   */
  async exitFullscreen(): Promise<void> {
    try {
      if (document.fullscreenElement) {
        await document.exitFullscreen();
      }
      this.isFullscreen = false;
      this.container.classList.remove('pptx-fullscreen');
      this.emit('fullscreenchange', false);
    } catch (error) {
      console.warn('Exit fullscreen failed:', error);
    }
  }

  /**
   * Toggles fullscreen mode.
   */
  async toggleFullscreen(): Promise<void> {
    if (this.isFullscreen) {
      await this.exitFullscreen();
    } else {
      await this.enterFullscreen();
    }
  }

  /**
   * Gets the current slide index (0-based).
   */
  getCurrentSlide(): number {
    return this.currentSlideIndex;
  }

  /**
   * Gets the total number of slides.
   */
  getSlideCount(): number {
    return this.presentation?.slides.length ?? 0;
  }

  /**
   * Gets the loaded presentation data.
   */
  getPresentation(): Presentation | null {
    return this.presentation;
  }

  /**
   * Registers an event listener.
   *
   * @param event - Event type to listen for
   * @param listener - Callback function
   *
   * @example
   * ```typescript
   * viewer.on('slidechange', (index) => {
   *   console.log(`Now on slide ${index + 1}`);
   * });
   * ```
   */
  on<K extends ViewerEventType>(event: K, listener: EventListener<ViewerEvents[K]>): void {
    if (!this.listeners.has(event)) {
      this.listeners.set(event, new Set());
    }
    this.listeners.get(event)!.add(listener);
  }

  /**
   * Removes an event listener.
   *
   * @param event - Event type
   * @param listener - Callback function to remove
   */
  off<K extends ViewerEventType>(event: K, listener: EventListener<ViewerEvents[K]>): void {
    this.listeners.get(event)?.delete(listener);
  }

  /**
   * Cleans up the viewer and releases resources.
   * Should be called when the viewer is no longer needed.
   */
  destroy(): void {
    // Clean up archive (revokes blob URLs)
    if (this.archive) {
      this.archive.cleanup();
      this.archive = null;
    }

    // Clear presentation
    this.presentation = null;

    // Remove event listeners
    this.listeners.clear();

    // Clear container
    this.container.innerHTML = '';

    // Remove keyboard listeners
    document.removeEventListener('keydown', this.handleKeydown);
    document.removeEventListener('fullscreenchange', this.handleFullscreenChange);
  }

  // =====================================================================
  // Private Methods
  // =====================================================================

  /**
   * Sets up the container structure.
   */
  private setupContainer(): void {
    this.container.classList.add('pptx-viewer');

    // Apply base styles
    Object.assign(this.container.style, {
      position: 'relative',
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      overflow: 'hidden',
      backgroundColor: '#1a1a1a',
    });

    // Create slide container
    this.slideContainer = document.createElement('div');
    this.slideContainer.className = 'pptx-slide-container';
    Object.assign(this.slideContainer.style, {
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      flex: '1',
      width: '100%',
      padding: '20px',
      boxSizing: 'border-box',
    });
    this.container.appendChild(this.slideContainer);

    // Create controls container
    if (this.options.showControls) {
      this.controlsContainer = document.createElement('div');
      this.controlsContainer.className = 'pptx-controls';
      Object.assign(this.controlsContainer.style, {
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        gap: '16px',
        padding: '12px',
        backgroundColor: 'rgba(0, 0, 0, 0.5)',
        borderRadius: '8px',
        marginBottom: '20px',
      });
      this.container.appendChild(this.controlsContainer);
      this.createControls();
    }

    // Listen for fullscreen changes
    document.addEventListener('fullscreenchange', this.handleFullscreenChange);
  }

  /**
   * Creates navigation controls.
   */
  private createControls(): void {
    if (!this.controlsContainer) return;

    const buttonStyle = {
      padding: '8px 16px',
      border: 'none',
      borderRadius: '4px',
      backgroundColor: '#4a4a4a',
      color: '#ffffff',
      cursor: 'pointer',
      fontSize: '14px',
      fontFamily: 'sans-serif',
    };

    // Previous button
    const prevButton = document.createElement('button');
    prevButton.className = 'pptx-prev';
    prevButton.textContent = '← Previous';
    Object.assign(prevButton.style, buttonStyle);
    prevButton.addEventListener('click', () => this.previous());
    this.controlsContainer.appendChild(prevButton);

    // Slide counter
    const counter = document.createElement('span');
    counter.className = 'pptx-counter';
    Object.assign(counter.style, {
      color: '#ffffff',
      fontSize: '14px',
      fontFamily: 'sans-serif',
      minWidth: '80px',
      textAlign: 'center',
    });
    counter.textContent = '0 / 0';
    this.controlsContainer.appendChild(counter);

    // Next button
    const nextButton = document.createElement('button');
    nextButton.className = 'pptx-next';
    nextButton.textContent = 'Next →';
    Object.assign(nextButton.style, buttonStyle);
    nextButton.addEventListener('click', () => this.next());
    this.controlsContainer.appendChild(nextButton);

    // Fullscreen button
    const fsButton = document.createElement('button');
    fsButton.className = 'pptx-fullscreen-btn';
    fsButton.textContent = '⛶';
    Object.assign(fsButton.style, {
      ...buttonStyle,
      fontSize: '18px',
      padding: '8px 12px',
    });
    fsButton.addEventListener('click', () => this.toggleFullscreen());
    this.controlsContainer.appendChild(fsButton);
  }

  /**
   * Updates the controls state.
   */
  private updateControls(): void {
    if (!this.controlsContainer || !this.presentation) return;

    const counter = this.controlsContainer.querySelector('.pptx-counter');
    if (counter) {
      counter.textContent = `${this.currentSlideIndex + 1} / ${this.presentation.slides.length}`;
    }

    const prevButton = this.controlsContainer.querySelector('.pptx-prev') as HTMLButtonElement;
    const nextButton = this.controlsContainer.querySelector('.pptx-next') as HTMLButtonElement;

    if (prevButton) {
      prevButton.disabled = this.currentSlideIndex === 0;
      prevButton.style.opacity = prevButton.disabled ? '0.5' : '1';
    }

    if (nextButton) {
      nextButton.disabled = this.currentSlideIndex === this.presentation.slides.length - 1;
      nextButton.style.opacity = nextButton.disabled ? '0.5' : '1';
    }
  }

  /**
   * Renders the current slide.
   */
  private renderCurrentSlide(): void {
    if (!this.slideContainer || !this.presentation) return;

    // Clear existing content
    this.slideContainer.innerHTML = '';

    const slide = this.presentation.slides[this.currentSlideIndex];
    const slideSize = this.presentation.slideSize;

    // Calculate display size
    const containerRect = this.slideContainer.getBoundingClientRect();
    const containerWidth = containerRect.width - 40; // Account for padding
    const containerHeight = containerRect.height - 40;

    const aspectRatio = slideSize.width / slideSize.height;
    let displayWidth = containerWidth;
    let displayHeight = displayWidth / aspectRatio;

    if (displayHeight > containerHeight) {
      displayHeight = containerHeight;
      displayWidth = displayHeight * aspectRatio;
    }

    // Render slide
    const svg = renderSlide(slide, slideSize, {
      width: displayWidth,
      height: displayHeight,
    });

    // Add shadow effect
    Object.assign(svg.style, {
      boxShadow: '0 4px 20px rgba(0, 0, 0, 0.5)',
      borderRadius: '2px',
    });

    this.slideContainer.appendChild(svg);
  }

  /**
   * Sets up keyboard navigation.
   */
  private setupKeyboardNavigation(): void {
    document.addEventListener('keydown', this.handleKeydown);
  }

  /**
   * Handles keyboard events.
   */
  private handleKeydown = (event: KeyboardEvent): void => {
    // Only handle if viewer is focused or in fullscreen
    if (!this.container.contains(document.activeElement) && !this.isFullscreen) {
      return;
    }

    switch (event.key) {
      case 'ArrowRight':
      case 'ArrowDown':
      case ' ':
      case 'PageDown':
        event.preventDefault();
        this.next();
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
      case 'PageUp':
        event.preventDefault();
        this.previous();
        break;
      case 'Home':
        event.preventDefault();
        this.goToSlide(0);
        break;
      case 'End':
        event.preventDefault();
        if (this.presentation) {
          this.goToSlide(this.presentation.slides.length - 1);
        }
        break;
      case 'Escape':
        if (this.isFullscreen) {
          this.exitFullscreen();
        }
        break;
      case 'f':
      case 'F':
        this.toggleFullscreen();
        break;
    }
  };

  /**
   * Handles fullscreen change events.
   */
  private handleFullscreenChange = (): void => {
    const wasFullscreen = this.isFullscreen;
    this.isFullscreen = document.fullscreenElement === this.container;

    if (wasFullscreen !== this.isFullscreen) {
      if (this.isFullscreen) {
        this.container.classList.add('pptx-fullscreen');
      } else {
        this.container.classList.remove('pptx-fullscreen');
      }
      this.emit('fullscreenchange', this.isFullscreen);

      // Re-render slide to fit new size
      setTimeout(() => this.renderCurrentSlide(), 100);
    }
  };

  /**
   * Emits an event to listeners.
   */
  private emit<K extends ViewerEventType>(event: K, data: ViewerEvents[K]): void {
    this.listeners.get(event)?.forEach((listener) => listener(data));
  }

  /**
   * Emits a slide change event.
   */
  private emitSlideChange(): void {
    this.emit('slidechange', this.currentSlideIndex);
    this.options.onSlideChange(this.currentSlideIndex);
  }
}
