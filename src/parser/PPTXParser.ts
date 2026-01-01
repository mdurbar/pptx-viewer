/**
 * Main parser for PPTX files.
 *
 * Orchestrates the parsing of all components:
 * - Presentation structure (slide order, dimensions)
 * - Theme (colors, fonts)
 * - Individual slides
 */

import type { Presentation, PresentationMetadata, Size, Slide, Theme } from '../core/types';
import type { PPTXArchive } from '../core/unzip';
import { PPTX_PATHS, getSlidePath } from '../core/unzip';
import { MissingFileError, XMLParseError, PPTXError } from '../core/errors';
import { parseXml, findFirstByName, findChildrenByName, getAttribute, getNumberAttribute } from '../utils/xml';
import { emuToPixels } from '../utils/units';
import { parseRelationships, RELATIONSHIP_TYPES, isRelationshipType } from './RelationshipParser';
import { parseTheme, createDefaultTheme } from './ThemeParser';
import { parseSlide } from './SlideParser';

/**
 * Parses a PPTX archive into a Presentation object.
 *
 * @param archive - Extracted PPTX archive
 * @returns Parsed presentation
 * @throws Error if the archive is not a valid PPTX file
 *
 * @example
 * const archive = await extractPPTX(file);
 * const presentation = await parsePPTX(archive);
 * console.log(`${presentation.slides.length} slides`);
 */
export async function parsePPTX(archive: PPTXArchive): Promise<Presentation> {
  // Validate it's a PPTX file
  validateArchive(archive);

  // Parse presentation.xml to get slide order and dimensions
  const presentationXml = archive.getText(PPTX_PATHS.PRESENTATION);
  if (!presentationXml) {
    throw new MissingFileError(PPTX_PATHS.PRESENTATION);
  }

  let slideSize: Size;
  let slideRIds: string[];

  try {
    const result = parsePresentationXml(presentationXml);
    slideSize = result.slideSize;
    slideRIds = result.slideRIds;
  } catch (error) {
    if (error instanceof PPTXError) throw error;
    throw new XMLParseError(
      error instanceof Error ? error.message : 'Unknown error',
      PPTX_PATHS.PRESENTATION
    );
  }

  // Parse relationships to get actual slide paths
  const relsXml = archive.getText(PPTX_PATHS.PRESENTATION_RELS);
  if (!relsXml) {
    throw new MissingFileError(PPTX_PATHS.PRESENTATION_RELS);
  }

  let relationships;
  try {
    relationships = parseRelationships(relsXml);
  } catch (error) {
    throw new XMLParseError(
      error instanceof Error ? error.message : 'Unknown error',
      PPTX_PATHS.PRESENTATION_RELS
    );
  }

  // Find and parse the theme
  const themeRels = relationships.getByType(RELATIONSHIP_TYPES.THEME);
  let theme: Theme;

  if (themeRels.length > 0) {
    const themePath = `ppt/${themeRels[0].target.replace('../', '')}`;
    const themeXml = archive.getText(themePath);
    try {
      theme = themeXml ? parseTheme(themeXml) : createDefaultTheme();
    } catch (error) {
      // Theme parsing failure is non-fatal, use defaults
      console.warn(`Failed to parse theme, using defaults:`, error);
      theme = createDefaultTheme();
    }
  } else {
    theme = createDefaultTheme();
  }

  // Parse metadata (non-fatal if it fails)
  const metadata = parseMetadata(archive);

  // Parse each slide
  const slides: Slide[] = [];
  const parseErrors: Array<{ slideIndex: number; error: unknown }> = [];

  for (let i = 0; i < slideRIds.length; i++) {
    const rId = slideRIds[i];
    const rel = relationships.get(rId);

    if (!rel) {
      console.warn(`Slide relationship ${rId} not found, skipping slide ${i + 1}`);
      parseErrors.push({ slideIndex: i, error: new Error(`Relationship ${rId} not found`) });
      continue;
    }

    // Resolve slide path
    const slidePath = `ppt/${rel.target.replace('../', '')}`;
    const slideXml = archive.getText(slidePath);

    if (!slideXml) {
      console.warn(`Slide file not found: ${slidePath}, skipping slide ${i + 1}`);
      parseErrors.push({ slideIndex: i, error: new MissingFileError(slidePath) });
      continue;
    }

    try {
      const slide = parseSlide(slideXml, i, archive, theme.colors, slidePath);
      slides.push(slide);
    } catch (error) {
      console.warn(`Failed to parse slide ${i + 1}:`, error);
      parseErrors.push({ slideIndex: i, error });
      // Create an empty placeholder slide so indices stay consistent
      slides.push({
        index: i,
        elements: [],
      });
    }
  }

  // If no slides could be parsed at all, that's a fatal error
  if (slides.length === 0 && slideRIds.length > 0) {
    throw new PPTXError('Failed to parse any slides from the presentation');
  }

  return {
    metadata,
    slideSize,
    slides,
    theme,
  };
}

/**
 * Validates that the archive is a valid PPTX file.
 */
function validateArchive(archive: PPTXArchive): void {
  // Check for required files
  if (!archive.hasFile(PPTX_PATHS.CONTENT_TYPES)) {
    throw new MissingFileError(PPTX_PATHS.CONTENT_TYPES);
  }

  if (!archive.hasFile(PPTX_PATHS.PRESENTATION)) {
    throw new MissingFileError(PPTX_PATHS.PRESENTATION);
  }
}

/**
 * Parses the main presentation.xml file.
 */
function parsePresentationXml(xmlContent: string): {
  slideSize: Size;
  slideRIds: string[];
} {
  const doc = parseXml(xmlContent);
  const root = doc.documentElement;

  // Parse slide size
  const sldSz = findFirstByName(root, 'sldSz');
  const slideSize: Size = sldSz
    ? {
        width: emuToPixels(getNumberAttribute(sldSz, 'cx', 9144000)),
        height: emuToPixels(getNumberAttribute(sldSz, 'cy', 6858000)),
      }
    : { width: 960, height: 720 }; // Default slide size

  // Parse slide ID list to get order
  const sldIdLst = findFirstByName(root, 'sldIdLst');
  const slideRIds: string[] = [];

  if (sldIdLst) {
    const sldIds = findChildrenByName(sldIdLst, 'sldId');
    for (const sldId of sldIds) {
      const rId = getAttribute(sldId, 'r:id');
      if (rId) {
        slideRIds.push(rId);
      }
    }
  }

  return { slideSize, slideRIds };
}

/**
 * Parses presentation metadata from docProps.
 */
function parseMetadata(archive: PPTXArchive): PresentationMetadata {
  const metadata: PresentationMetadata = {};

  const coreXml = archive.getText(PPTX_PATHS.CORE_PROPS);
  if (!coreXml) return metadata;

  try {
    const doc = parseXml(coreXml);
    const root = doc.documentElement;

    // Title
    const title = findFirstByName(root, 'title');
    if (title?.textContent) {
      metadata.title = title.textContent.trim();
    }

    // Author/Creator
    const creator = findFirstByName(root, 'creator');
    if (creator?.textContent) {
      metadata.author = creator.textContent.trim();
    }

    // Created date
    const created = findFirstByName(root, 'created');
    if (created?.textContent) {
      metadata.createdAt = created.textContent.trim();
    }

    // Modified date
    const modified = findFirstByName(root, 'modified');
    if (modified?.textContent) {
      metadata.modifiedAt = modified.textContent.trim();
    }
  } catch (e) {
    // Metadata parsing is optional, don't fail if it errors
    console.warn('Failed to parse metadata:', e);
  }

  return metadata;
}

/**
 * Quick check if a file is a valid PPTX.
 *
 * @param archive - Extracted archive to check
 * @returns True if the archive appears to be a valid PPTX
 */
export function isValidPPTX(archive: PPTXArchive): boolean {
  return (
    archive.hasFile(PPTX_PATHS.CONTENT_TYPES) &&
    archive.hasFile(PPTX_PATHS.PRESENTATION)
  );
}

/**
 * Gets the slide count without fully parsing the presentation.
 *
 * @param archive - Extracted PPTX archive
 * @returns Number of slides
 */
export function getSlideCount(archive: PPTXArchive): number {
  const presentationXml = archive.getText(PPTX_PATHS.PRESENTATION);
  if (!presentationXml) return 0;

  try {
    const doc = parseXml(presentationXml);
    const sldIdLst = findFirstByName(doc.documentElement, 'sldIdLst');
    if (!sldIdLst) return 0;

    return findChildrenByName(sldIdLst, 'sldId').length;
  } catch {
    return 0;
  }
}
