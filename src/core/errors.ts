/**
 * Custom error types for PPTX Viewer.
 *
 * These provide more specific error information for debugging and handling.
 */

/**
 * Base error class for all PPTX Viewer errors.
 */
export class PPTXError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'PPTXError';
  }
}

/**
 * Error thrown when the input is not a valid ZIP/PPTX file.
 */
export class InvalidFileError extends PPTXError {
  constructor(message: string = 'The file is not a valid PPTX file') {
    super(message);
    this.name = 'InvalidFileError';
  }
}

/**
 * Error thrown when a required file is missing from the archive.
 */
export class MissingFileError extends PPTXError {
  public readonly filePath: string;

  constructor(filePath: string) {
    super(`Required file not found in PPTX: ${filePath}`);
    this.name = 'MissingFileError';
    this.filePath = filePath;
  }
}

/**
 * Error thrown when XML parsing fails.
 */
export class XMLParseError extends PPTXError {
  public readonly filePath?: string;

  constructor(message: string, filePath?: string) {
    super(filePath ? `Failed to parse XML in ${filePath}: ${message}` : `XML parse error: ${message}`);
    this.name = 'XMLParseError';
    this.filePath = filePath;
  }
}

/**
 * Error thrown when fetching a remote file fails.
 */
export class FetchError extends PPTXError {
  public readonly url: string;
  public readonly status?: number;

  constructor(url: string, status?: number, statusText?: string) {
    const message = status
      ? `Failed to fetch PPTX from ${url}: ${status} ${statusText}`
      : `Failed to fetch PPTX from ${url}`;
    super(message);
    this.name = 'FetchError';
    this.url = url;
    this.status = status;
  }
}

/**
 * Error thrown when rendering fails.
 */
export class RenderError extends PPTXError {
  public readonly slideIndex?: number;
  public readonly elementId?: string;

  constructor(message: string, slideIndex?: number, elementId?: string) {
    super(message);
    this.name = 'RenderError';
    this.slideIndex = slideIndex;
    this.elementId = elementId;
  }
}

/**
 * Wraps an error with additional context.
 *
 * @param error - Original error
 * @param context - Additional context message
 * @returns Wrapped error
 */
export function wrapError(error: unknown, context: string): PPTXError {
  const message = error instanceof Error ? error.message : String(error);
  return new PPTXError(`${context}: ${message}`);
}

/**
 * Type guard to check if an error is a PPTXError.
 */
export function isPPTXError(error: unknown): error is PPTXError {
  return error instanceof PPTXError;
}
