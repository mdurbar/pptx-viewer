import { describe, it, expect } from 'vitest';
import {
  PPTXError,
  InvalidFileError,
  MissingFileError,
  XMLParseError,
  FetchError,
  RenderError,
  wrapError,
  isPPTXError,
} from '../src/core/errors';

describe('Error Types', () => {
  describe('PPTXError', () => {
    it('creates error with message', () => {
      const error = new PPTXError('Something went wrong');
      expect(error.message).toBe('Something went wrong');
      expect(error.name).toBe('PPTXError');
      expect(error instanceof Error).toBe(true);
    });
  });

  describe('InvalidFileError', () => {
    it('creates error with default message', () => {
      const error = new InvalidFileError();
      expect(error.message).toBe('The file is not a valid PPTX file');
      expect(error.name).toBe('InvalidFileError');
    });

    it('creates error with custom message', () => {
      const error = new InvalidFileError('File is corrupted');
      expect(error.message).toBe('File is corrupted');
    });

    it('is instance of PPTXError', () => {
      const error = new InvalidFileError();
      expect(error instanceof PPTXError).toBe(true);
    });
  });

  describe('MissingFileError', () => {
    it('creates error with file path', () => {
      const error = new MissingFileError('ppt/slides/slide1.xml');
      expect(error.message).toContain('ppt/slides/slide1.xml');
      expect(error.filePath).toBe('ppt/slides/slide1.xml');
      expect(error.name).toBe('MissingFileError');
    });

    it('is instance of PPTXError', () => {
      const error = new MissingFileError('test.xml');
      expect(error instanceof PPTXError).toBe(true);
    });
  });

  describe('XMLParseError', () => {
    it('creates error with message only', () => {
      const error = new XMLParseError('Invalid XML syntax');
      expect(error.message).toBe('XML parse error: Invalid XML syntax');
      expect(error.filePath).toBeUndefined();
      expect(error.name).toBe('XMLParseError');
    });

    it('creates error with file path', () => {
      const error = new XMLParseError('Missing closing tag', 'ppt/presentation.xml');
      expect(error.message).toContain('ppt/presentation.xml');
      expect(error.message).toContain('Missing closing tag');
      expect(error.filePath).toBe('ppt/presentation.xml');
    });

    it('is instance of PPTXError', () => {
      const error = new XMLParseError('test');
      expect(error instanceof PPTXError).toBe(true);
    });
  });

  describe('FetchError', () => {
    it('creates error with URL only', () => {
      const error = new FetchError('https://example.com/presentation.pptx');
      expect(error.message).toContain('https://example.com/presentation.pptx');
      expect(error.url).toBe('https://example.com/presentation.pptx');
      expect(error.status).toBeUndefined();
      expect(error.name).toBe('FetchError');
    });

    it('creates error with status code', () => {
      const error = new FetchError('https://example.com/file.pptx', 404, 'Not Found');
      expect(error.message).toContain('404');
      expect(error.message).toContain('Not Found');
      expect(error.status).toBe(404);
    });

    it('is instance of PPTXError', () => {
      const error = new FetchError('http://test.com');
      expect(error instanceof PPTXError).toBe(true);
    });
  });

  describe('RenderError', () => {
    it('creates error with message only', () => {
      const error = new RenderError('Failed to render shape');
      expect(error.message).toBe('Failed to render shape');
      expect(error.slideIndex).toBeUndefined();
      expect(error.elementId).toBeUndefined();
      expect(error.name).toBe('RenderError');
    });

    it('creates error with slide index', () => {
      const error = new RenderError('Render failed', 2);
      expect(error.slideIndex).toBe(2);
    });

    it('creates error with element ID', () => {
      const error = new RenderError('Render failed', 0, 'shape-123');
      expect(error.slideIndex).toBe(0);
      expect(error.elementId).toBe('shape-123');
    });

    it('is instance of PPTXError', () => {
      const error = new RenderError('test');
      expect(error instanceof PPTXError).toBe(true);
    });
  });

  describe('wrapError', () => {
    it('wraps Error instance', () => {
      const original = new Error('Original error');
      const wrapped = wrapError(original, 'Context');
      expect(wrapped.message).toBe('Context: Original error');
      expect(wrapped instanceof PPTXError).toBe(true);
    });

    it('wraps string', () => {
      const wrapped = wrapError('String error', 'Context');
      expect(wrapped.message).toBe('Context: String error');
    });

    it('wraps unknown types', () => {
      const wrapped = wrapError(42, 'Context');
      expect(wrapped.message).toBe('Context: 42');
    });
  });

  describe('isPPTXError', () => {
    it('returns true for PPTXError', () => {
      expect(isPPTXError(new PPTXError('test'))).toBe(true);
    });

    it('returns true for subclasses', () => {
      expect(isPPTXError(new InvalidFileError())).toBe(true);
      expect(isPPTXError(new MissingFileError('test'))).toBe(true);
      expect(isPPTXError(new XMLParseError('test'))).toBe(true);
      expect(isPPTXError(new FetchError('test'))).toBe(true);
      expect(isPPTXError(new RenderError('test'))).toBe(true);
    });

    it('returns false for regular Error', () => {
      expect(isPPTXError(new Error('test'))).toBe(false);
    });

    it('returns false for non-errors', () => {
      expect(isPPTXError('string')).toBe(false);
      expect(isPPTXError(null)).toBe(false);
      expect(isPPTXError(undefined)).toBe(false);
      expect(isPPTXError({})).toBe(false);
    });
  });
});
