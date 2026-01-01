/**
 * ZIP extraction utilities for PPTX files.
 *
 * PPTX files are ZIP archives containing XML and media files.
 * Uses fflate for efficient decompression.
 */

import { unzipSync, type Unzipped } from 'fflate';

/**
 * Represents a file extracted from the PPTX archive.
 */
export interface ZipEntry {
  /** File path within the archive */
  path: string;
  /** File content as Uint8Array */
  data: Uint8Array;
}

/**
 * Extracted PPTX archive contents.
 */
export interface PPTXArchive {
  /** All files in the archive, keyed by path */
  files: Map<string, Uint8Array>;

  /**
   * Gets a file's content as a string (UTF-8).
   * @param path - File path within the archive
   * @returns File content as string, or null if not found
   */
  getText(path: string): string | null;

  /**
   * Gets a file's content as raw bytes.
   * @param path - File path within the archive
   * @returns File content as Uint8Array, or null if not found
   */
  getBytes(path: string): Uint8Array | null;

  /**
   * Gets a file as a Blob URL (for images/media).
   * @param path - File path within the archive
   * @param mimeType - MIME type for the blob
   * @returns Blob URL, or null if not found
   */
  getBlobUrl(path: string, mimeType: string): string | null;

  /**
   * Lists all files in the archive.
   * @returns Array of file paths
   */
  listFiles(): string[];

  /**
   * Checks if a file exists in the archive.
   * @param path - File path to check
   * @returns True if file exists
   */
  hasFile(path: string): boolean;

  /**
   * Cleans up blob URLs to prevent memory leaks.
   * Should be called when done with the archive.
   */
  cleanup(): void;
}

/**
 * Extracts a PPTX file from various input sources.
 *
 * @param source - PPTX file as File, ArrayBuffer, Uint8Array, or URL
 * @returns Extracted archive contents
 *
 * @example
 * // From File input
 * const file = inputElement.files[0];
 * const archive = await extractPPTX(file);
 *
 * @example
 * // From URL
 * const archive = await extractPPTX('/presentations/demo.pptx');
 *
 * @example
 * // From ArrayBuffer
 * const buffer = await fetch(url).then(r => r.arrayBuffer());
 * const archive = await extractPPTX(buffer);
 */
export async function extractPPTX(
  source: File | ArrayBuffer | Uint8Array | string
): Promise<PPTXArchive> {
  let data: Uint8Array;

  if (typeof source === 'string') {
    // URL - fetch the file
    const response = await fetch(source);
    if (!response.ok) {
      throw new Error(`Failed to fetch PPTX: ${response.status} ${response.statusText}`);
    }
    const buffer = await response.arrayBuffer();
    data = new Uint8Array(buffer);
  } else if (source instanceof File) {
    // File object - read as ArrayBuffer
    const buffer = await source.arrayBuffer();
    data = new Uint8Array(buffer);
  } else if (source instanceof ArrayBuffer) {
    // ArrayBuffer - convert to Uint8Array
    data = new Uint8Array(source);
  } else if (source instanceof Uint8Array) {
    // Already a Uint8Array
    data = source;
  } else {
    throw new Error('Invalid source type. Expected File, ArrayBuffer, Uint8Array, or URL string.');
  }

  // Extract the ZIP archive
  const unzipped = unzipSync(data);

  return createArchive(unzipped);
}

/**
 * Creates a PPTXArchive from unzipped data.
 */
function createArchive(unzipped: Unzipped): PPTXArchive {
  const files = new Map<string, Uint8Array>();
  const blobUrls = new Set<string>();

  // Normalize paths (remove leading slash if present)
  for (const [path, data] of Object.entries(unzipped)) {
    const normalizedPath = path.startsWith('/') ? path.slice(1) : path;
    files.set(normalizedPath, data);
  }

  const decoder = new TextDecoder('utf-8');

  return {
    files,

    getText(path: string): string | null {
      const data = files.get(path);
      if (!data) return null;
      return decoder.decode(data);
    },

    getBytes(path: string): Uint8Array | null {
      return files.get(path) || null;
    },

    getBlobUrl(path: string, mimeType: string): string | null {
      const data = files.get(path);
      if (!data) return null;

      // Create a copy to ensure we have a standard ArrayBuffer
      const buffer = new ArrayBuffer(data.length);
      new Uint8Array(buffer).set(data);
      const blob = new Blob([buffer], { type: mimeType });
      const url = URL.createObjectURL(blob);
      blobUrls.add(url);
      return url;
    },

    listFiles(): string[] {
      return Array.from(files.keys());
    },

    hasFile(path: string): boolean {
      return files.has(path);
    },

    cleanup(): void {
      for (const url of blobUrls) {
        URL.revokeObjectURL(url);
      }
      blobUrls.clear();
    },
  };
}

/**
 * Common PPTX file paths.
 */
export const PPTX_PATHS = {
  /** Content types definition */
  CONTENT_TYPES: '[Content_Types].xml',
  /** Main relationships */
  RELS: '_rels/.rels',
  /** Presentation relationships */
  PRESENTATION_RELS: 'ppt/_rels/presentation.xml.rels',
  /** Main presentation file */
  PRESENTATION: 'ppt/presentation.xml',
  /** Core properties (metadata) */
  CORE_PROPS: 'docProps/core.xml',
  /** App properties */
  APP_PROPS: 'docProps/app.xml',
} as const;

/**
 * Gets the path to a slide XML file.
 * @param slideNumber - 1-based slide number
 */
export function getSlidePath(slideNumber: number): string {
  return `ppt/slides/slide${slideNumber}.xml`;
}

/**
 * Gets the path to a slide's relationships file.
 * @param slideNumber - 1-based slide number
 */
export function getSlideRelsPath(slideNumber: number): string {
  return `ppt/slides/_rels/slide${slideNumber}.xml.rels`;
}

/**
 * Determines MIME type from file extension.
 * @param path - File path
 * @returns MIME type
 */
export function getMimeType(path: string): string {
  const ext = path.split('.').pop()?.toLowerCase();

  const mimeTypes: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    tiff: 'image/tiff',
    tif: 'image/tiff',
    svg: 'image/svg+xml',
    emf: 'image/emf',
    wmf: 'image/wmf',
    webp: 'image/webp',
    mp4: 'video/mp4',
    webm: 'video/webm',
    mp3: 'audio/mpeg',
    wav: 'audio/wav',
  };

  return mimeTypes[ext || ''] || 'application/octet-stream';
}
