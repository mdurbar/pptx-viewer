/**
 * Parser for embedded fonts in PPTX files.
 *
 * PPTX files can embed fonts in the ppt/fonts/ directory.
 * Regular fonts are stored as .ttf files, while obfuscated fonts
 * use the .odttf extension (XOR encrypted with a GUID).
 */

import type { EmbeddedFont } from '../core/types';
import type { PPTXArchive } from '../core/unzip';

/**
 * Result of parsing embedded fonts.
 */
export interface FontParseResult {
  fonts: Map<string, EmbeddedFont>;
}

/**
 * Parses embedded fonts from a PPTX archive.
 *
 * @param archive - PPTX archive to extract fonts from
 * @returns Map of font names to EmbeddedFont objects
 */
export function parseEmbeddedFonts(archive: PPTXArchive): FontParseResult {
  const fonts = new Map<string, EmbeddedFont>();

  // List all files in the ppt/fonts/ directory
  const allFiles = archive.listFiles();
  const fontFiles = allFiles.filter(
    (path) =>
      path.startsWith('ppt/fonts/') &&
      (path.endsWith('.ttf') || path.endsWith('.odttf') || path.endsWith('.otf'))
  );

  for (const fontPath of fontFiles) {
    try {
      const font = extractFont(archive, fontPath);
      if (font) {
        fonts.set(font.name, font);
      }
    } catch (error) {
      console.warn(`Failed to extract font from ${fontPath}:`, error);
    }
  }

  return { fonts };
}

/**
 * Extracts a single font from the archive.
 *
 * @param archive - PPTX archive
 * @param fontPath - Path to the font file in the archive
 * @returns EmbeddedFont or null if extraction fails
 */
function extractFont(archive: PPTXArchive, fontPath: string): EmbeddedFont | null {
  let fontData = archive.getBytes(fontPath);
  if (!fontData) return null;

  const filename = fontPath.split('/').pop() || '';
  const isObfuscated = fontPath.endsWith('.odttf');

  // Deobfuscate if necessary
  if (isObfuscated) {
    fontData = deobfuscateOdttf(fontData, filename);
  }

  // Extract font name from filename
  const fontName = extractFontName(filename);
  if (!fontName) return null;

  // Determine format
  const format = fontPath.endsWith('.otf') ? 'opentype' : 'truetype';

  // Create blob URL
  // Copy to a standard ArrayBuffer to ensure compatibility
  const mimeType = format === 'opentype' ? 'font/otf' : 'font/ttf';
  const buffer = new ArrayBuffer(fontData.length);
  new Uint8Array(buffer).set(fontData);
  const blob = new Blob([buffer], { type: mimeType });
  const url = URL.createObjectURL(blob);

  return {
    name: fontName,
    url,
    format,
    path: fontPath,
  };
}

/**
 * Extracts the font name from a font filename.
 *
 * Font files in PPTX can have various naming patterns:
 * - "Arial.ttf" -> "Arial"
 * - "Arial+{GUID}.odttf" -> "Arial"
 * - "Calibri-Regular.ttf" -> "Calibri"
 * - "Open Sans_Regular.ttf" -> "Open Sans"
 *
 * @param filename - The font filename
 * @returns Extracted font name or null
 */
export function extractFontName(filename: string): string | null {
  if (!filename) return null;

  // Remove extension
  let name = filename.replace(/\.(ttf|otf|odttf)$/i, '');

  // Remove GUID pattern for obfuscated fonts: "FontName+{GUID}" or "FontName_{GUID}"
  name = name.replace(/[+_]\{[0-9A-Fa-f-]+\}$/, '');

  // Remove common weight/style suffixes
  name = name.replace(/[-_](Regular|Bold|Italic|Light|Medium|SemiBold|ExtraBold|Black|Thin|BoldItalic)$/i, '');

  // Replace underscores with spaces (some fonts use underscores)
  name = name.replace(/_/g, ' ');

  // Trim and validate
  name = name.trim();

  return name || null;
}

/**
 * Deobfuscates an ODTTF font file.
 *
 * ODTTF files are XOR-encrypted TTF files. The encryption key is derived
 * from a GUID embedded in the filename. The first 32 bytes of the file
 * are XORed with the GUID bytes (repeated twice).
 *
 * @param data - The obfuscated font data
 * @param filename - The font filename containing the GUID
 * @returns Deobfuscated font data
 */
export function deobfuscateOdttf(data: Uint8Array, filename: string): Uint8Array {
  // Extract GUID from filename: "FontName+{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}.odttf"
  const guidMatch = filename.match(/\{([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})\}/i);

  if (!guidMatch) {
    // No GUID found, return data as-is (might not be obfuscated)
    return data;
  }

  // Parse GUID into bytes
  // GUID format: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
  // The GUID bytes are in a specific order for ODTTF:
  // First 4 bytes are reversed, next 2 reversed, next 2 reversed, rest as-is
  const guidBytes = parseGuidToBytes(guidMatch[1]);

  // XOR the first 32 bytes with the GUID bytes (repeated twice)
  const result = new Uint8Array(data);
  const bytesToDecrypt = Math.min(32, data.length);

  for (let i = 0; i < bytesToDecrypt; i++) {
    result[i] = data[i] ^ guidBytes[i % 16];
  }

  return result;
}

/**
 * Parses a GUID string into bytes in the correct order for ODTTF deobfuscation.
 *
 * GUID byte order for ODTTF is little-endian for the first three groups:
 * - Bytes 0-3: First group (4 bytes) - reversed
 * - Bytes 4-5: Second group (2 bytes) - reversed
 * - Bytes 6-7: Third group (2 bytes) - reversed
 * - Bytes 8-15: Fourth and fifth groups (8 bytes) - as-is
 *
 * @param guid - GUID string (e.g., "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX")
 * @returns 16-byte array
 */
export function parseGuidToBytes(guid: string): Uint8Array {
  // Remove hyphens
  const hex = guid.replace(/-/g, '');

  // Parse into byte array with correct endianness
  const bytes = new Uint8Array(16);

  // First group: 4 bytes, little-endian (reversed)
  bytes[0] = parseInt(hex.substring(6, 8), 16);
  bytes[1] = parseInt(hex.substring(4, 6), 16);
  bytes[2] = parseInt(hex.substring(2, 4), 16);
  bytes[3] = parseInt(hex.substring(0, 2), 16);

  // Second group: 2 bytes, little-endian (reversed)
  bytes[4] = parseInt(hex.substring(10, 12), 16);
  bytes[5] = parseInt(hex.substring(8, 10), 16);

  // Third group: 2 bytes, little-endian (reversed)
  bytes[6] = parseInt(hex.substring(14, 16), 16);
  bytes[7] = parseInt(hex.substring(12, 14), 16);

  // Fourth and fifth groups: 8 bytes, big-endian (as-is)
  for (let i = 0; i < 8; i++) {
    bytes[8 + i] = parseInt(hex.substring(16 + i * 2, 18 + i * 2), 16);
  }

  return bytes;
}

/**
 * Cleans up font blob URLs.
 *
 * @param fonts - Map of embedded fonts to clean up
 */
export function cleanupFontUrls(fonts: Map<string, EmbeddedFont>): void {
  for (const font of fonts.values()) {
    URL.revokeObjectURL(font.url);
  }
}
