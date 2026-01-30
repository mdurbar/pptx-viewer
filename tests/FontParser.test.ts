import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  parseEmbeddedFonts,
  extractFontName,
  deobfuscateOdttf,
  parseGuidToBytes,
  cleanupFontUrls,
} from '../src/parser/FontParser';
import type { PPTXArchive } from '../src/core/unzip';
import type { EmbeddedFont } from '../src/core/types';

// Mock URL.createObjectURL and URL.revokeObjectURL
const mockBlobUrls: string[] = [];
let blobUrlCounter = 0;

beforeEach(() => {
  mockBlobUrls.length = 0;
  blobUrlCounter = 0;

  vi.stubGlobal('URL', {
    createObjectURL: (blob: Blob) => {
      const url = `blob:test-${blobUrlCounter++}`;
      mockBlobUrls.push(url);
      return url;
    },
    revokeObjectURL: (url: string) => {
      const index = mockBlobUrls.indexOf(url);
      if (index >= 0) {
        mockBlobUrls.splice(index, 1);
      }
    },
  });
});

afterEach(() => {
  vi.unstubAllGlobals();
});

// Helper to create a mock archive
function createMockArchive(files: Record<string, Uint8Array | null>): PPTXArchive {
  const fileMap = new Map(
    Object.entries(files).filter(([_, data]) => data !== null) as [string, Uint8Array][]
  );

  return {
    files: fileMap,
    getText: () => null,
    getBytes: (path: string) => fileMap.get(path) || null,
    getBlobUrl: () => null,
    listFiles: () => Array.from(fileMap.keys()),
    hasFile: (path: string) => fileMap.has(path),
    cleanup: () => {},
  };
}

// Sample TTF header (simplified - just needs to be non-empty)
function createMockTtfData(): Uint8Array {
  // TTF magic number: 00 01 00 00
  return new Uint8Array([0x00, 0x01, 0x00, 0x00, ...Array(100).fill(0)]);
}

describe('FontParser', () => {
  describe('extractFontName', () => {
    it('extracts font name from simple filename', () => {
      expect(extractFontName('Arial.ttf')).toBe('Arial');
      expect(extractFontName('Calibri.otf')).toBe('Calibri');
    });

    it('extracts font name from filename with GUID', () => {
      expect(extractFontName('Arial+{12345678-1234-1234-1234-123456789ABC}.odttf')).toBe('Arial');
      expect(extractFontName('Calibri_{ABCDEF12-3456-7890-ABCD-EF1234567890}.odttf')).toBe('Calibri');
    });

    it('removes weight/style suffixes', () => {
      expect(extractFontName('Arial-Regular.ttf')).toBe('Arial');
      expect(extractFontName('Calibri-Bold.ttf')).toBe('Calibri');
      expect(extractFontName('OpenSans_Italic.ttf')).toBe('OpenSans');
      expect(extractFontName('Roboto-BoldItalic.ttf')).toBe('Roboto');
    });

    it('replaces underscores with spaces', () => {
      expect(extractFontName('Open_Sans.ttf')).toBe('Open Sans');
    });

    it('handles complex filenames', () => {
      expect(extractFontName('Open_Sans-Bold+{12345678-1234-1234-1234-123456789ABC}.odttf')).toBe('Open Sans');
    });

    it('returns null for empty filename', () => {
      expect(extractFontName('')).toBeNull();
    });
  });

  describe('parseGuidToBytes', () => {
    it('parses GUID into bytes with correct endianness', () => {
      // Test GUID: 12345678-ABCD-EF12-3456-789ABCDEF012
      const bytes = parseGuidToBytes('12345678-ABCD-EF12-3456-789ABCDEF012');

      // First 4 bytes: 12345678 -> 78 56 34 12 (little-endian)
      expect(bytes[0]).toBe(0x78);
      expect(bytes[1]).toBe(0x56);
      expect(bytes[2]).toBe(0x34);
      expect(bytes[3]).toBe(0x12);

      // Next 2 bytes: ABCD -> CD AB (little-endian)
      expect(bytes[4]).toBe(0xcd);
      expect(bytes[5]).toBe(0xab);

      // Next 2 bytes: EF12 -> 12 EF (little-endian)
      expect(bytes[6]).toBe(0x12);
      expect(bytes[7]).toBe(0xef);

      // Remaining 8 bytes: 3456789ABCDEF012 (big-endian, as-is)
      expect(bytes[8]).toBe(0x34);
      expect(bytes[9]).toBe(0x56);
      expect(bytes[10]).toBe(0x78);
      expect(bytes[11]).toBe(0x9a);
      expect(bytes[12]).toBe(0xbc);
      expect(bytes[13]).toBe(0xde);
      expect(bytes[14]).toBe(0xf0);
      expect(bytes[15]).toBe(0x12);
    });
  });

  describe('deobfuscateOdttf', () => {
    it('returns data unchanged if no GUID in filename', () => {
      const data = new Uint8Array([1, 2, 3, 4, 5]);
      const result = deobfuscateOdttf(data, 'Arial.ttf');

      expect(result).toEqual(data);
    });

    it('XORs first 32 bytes with GUID bytes', () => {
      // Create test data
      const data = new Uint8Array(64);
      for (let i = 0; i < 64; i++) {
        data[i] = i;
      }

      // Use a simple GUID for testing
      const filename = 'Test+{00000000-0000-0000-0000-000000000000}.odttf';
      const result = deobfuscateOdttf(data, filename);

      // With all-zero GUID, XOR should leave first 32 bytes unchanged
      for (let i = 0; i < 32; i++) {
        expect(result[i]).toBe(i);
      }

      // Bytes 32+ should be unchanged
      for (let i = 32; i < 64; i++) {
        expect(result[i]).toBe(i);
      }
    });

    it('correctly deobfuscates with non-zero GUID', () => {
      // Test with known values
      const data = new Uint8Array([
        0x78, 0x56, 0x34, 0x12, // Will XOR with first 4 GUID bytes (reversed)
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, // Rest of first 16 bytes
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, // Second 16 bytes
        0xff, 0xff, // After first 32 bytes - unchanged
      ]);

      const filename = 'Font+{12345678-0000-0000-0000-000000000000}.odttf';
      const result = deobfuscateOdttf(data, filename);

      // First 4 bytes: 78 56 34 12 XOR 78 56 34 12 = 00 00 00 00
      expect(result[0]).toBe(0);
      expect(result[1]).toBe(0);
      expect(result[2]).toBe(0);
      expect(result[3]).toBe(0);

      // Bytes after 32 should be unchanged
      expect(result[32]).toBe(0xff);
      expect(result[33]).toBe(0xff);
    });

    it('handles data shorter than 32 bytes', () => {
      const data = new Uint8Array([1, 2, 3, 4, 5]);
      const filename = 'Font+{00000000-0000-0000-0000-000000000000}.odttf';

      // Should not throw
      const result = deobfuscateOdttf(data, filename);
      expect(result.length).toBe(5);
    });
  });

  describe('parseEmbeddedFonts', () => {
    it('returns empty map when no fonts directory', () => {
      const archive = createMockArchive({
        'ppt/presentation.xml': new Uint8Array([1, 2, 3]),
      });

      const { fonts } = parseEmbeddedFonts(archive);

      expect(fonts.size).toBe(0);
    });

    it('extracts TTF fonts', () => {
      const archive = createMockArchive({
        'ppt/fonts/Arial.ttf': createMockTtfData(),
      });

      const { fonts } = parseEmbeddedFonts(archive);

      expect(fonts.size).toBe(1);
      expect(fonts.has('Arial')).toBe(true);

      const font = fonts.get('Arial')!;
      expect(font.name).toBe('Arial');
      expect(font.format).toBe('truetype');
      expect(font.path).toBe('ppt/fonts/Arial.ttf');
      expect(font.url).toMatch(/^blob:/);
    });

    it('extracts OTF fonts', () => {
      const archive = createMockArchive({
        'ppt/fonts/Roboto.otf': createMockTtfData(),
      });

      const { fonts } = parseEmbeddedFonts(archive);

      expect(fonts.size).toBe(1);

      const font = fonts.get('Roboto')!;
      expect(font.format).toBe('opentype');
    });

    it('extracts and deobfuscates ODTTF fonts', () => {
      const archive = createMockArchive({
        'ppt/fonts/CustomFont+{12345678-1234-1234-1234-123456789ABC}.odttf': createMockTtfData(),
      });

      const { fonts } = parseEmbeddedFonts(archive);

      expect(fonts.size).toBe(1);
      expect(fonts.has('CustomFont')).toBe(true);

      const font = fonts.get('CustomFont')!;
      expect(font.format).toBe('truetype');
    });

    it('extracts multiple fonts', () => {
      const archive = createMockArchive({
        'ppt/fonts/Arial.ttf': createMockTtfData(),
        'ppt/fonts/Calibri.ttf': createMockTtfData(),
        'ppt/fonts/CustomFont+{12345678-1234-1234-1234-123456789ABC}.odttf': createMockTtfData(),
      });

      const { fonts } = parseEmbeddedFonts(archive);

      expect(fonts.size).toBe(3);
      expect(fonts.has('Arial')).toBe(true);
      expect(fonts.has('Calibri')).toBe(true);
      expect(fonts.has('CustomFont')).toBe(true);
    });

    it('ignores non-font files in fonts directory', () => {
      const archive = createMockArchive({
        'ppt/fonts/Arial.ttf': createMockTtfData(),
        'ppt/fonts/readme.txt': new Uint8Array([1, 2, 3]),
        'ppt/fonts/.DS_Store': new Uint8Array([1, 2, 3]),
      });

      const { fonts } = parseEmbeddedFonts(archive);

      expect(fonts.size).toBe(1);
      expect(fonts.has('Arial')).toBe(true);
    });

    it('handles font files that fail to extract', () => {
      // Create archive where getBytes returns null for one font
      const archive = createMockArchive({
        'ppt/fonts/Arial.ttf': createMockTtfData(),
      });

      // Override listFiles to include a font that doesn't exist
      archive.listFiles = () => ['ppt/fonts/Arial.ttf', 'ppt/fonts/Missing.ttf'];

      const { fonts } = parseEmbeddedFonts(archive);

      // Should still extract the valid font
      expect(fonts.size).toBe(1);
      expect(fonts.has('Arial')).toBe(true);
    });
  });

  describe('cleanupFontUrls', () => {
    it('revokes blob URLs for all fonts', () => {
      const fonts = new Map<string, EmbeddedFont>([
        ['Arial', { name: 'Arial', url: 'blob:test-0', format: 'truetype', path: 'ppt/fonts/Arial.ttf' }],
        ['Calibri', { name: 'Calibri', url: 'blob:test-1', format: 'truetype', path: 'ppt/fonts/Calibri.ttf' }],
      ]);

      // Simulate having these URLs tracked
      mockBlobUrls.push('blob:test-0', 'blob:test-1');

      cleanupFontUrls(fonts);

      expect(mockBlobUrls).toEqual([]);
    });

    it('handles empty font map', () => {
      const fonts = new Map<string, EmbeddedFont>();

      // Should not throw
      cleanupFontUrls(fonts);
    });
  });
});
