/**
 * Parser for OOXML relationship files (.rels).
 *
 * Relationships define connections between parts of the PPTX package.
 * They map relationship IDs (rId1, rId2, etc.) to target file paths.
 */

import { parseXml, getAttribute, findChildrenByName } from '../utils/xml';

/**
 * A single relationship entry.
 */
export interface Relationship {
  /** Relationship ID (e.g., "rId1") */
  id: string;
  /** Relationship type URL */
  type: string;
  /** Target file path (relative to the .rels file's parent) */
  target: string;
  /** Whether target is external (URL) */
  external: boolean;
}

/**
 * Common relationship type suffixes.
 */
export const RELATIONSHIP_TYPES = {
  SLIDE: 'relationships/slide',
  SLIDE_LAYOUT: 'relationships/slideLayout',
  SLIDE_MASTER: 'relationships/slideMaster',
  THEME: 'relationships/theme',
  IMAGE: 'relationships/image',
  CHART: 'relationships/chart',
  OLE_OBJECT: 'relationships/oleObject',
  HYPERLINK: 'relationships/hyperlink',
  NOTES_SLIDE: 'relationships/notesSlide',
  OFFICE_DOCUMENT: 'relationships/officeDocument',
} as const;

/**
 * Parsed relationships file.
 */
export interface RelationshipMap {
  /** All relationships keyed by ID */
  byId: Map<string, Relationship>;
  /** Relationships grouped by type */
  byType: Map<string, Relationship[]>;

  /**
   * Gets a relationship by ID.
   * @param id - Relationship ID (e.g., "rId1")
   */
  get(id: string): Relationship | undefined;

  /**
   * Gets all relationships of a specific type.
   * @param type - Relationship type (can be partial match)
   */
  getByType(type: string): Relationship[];

  /**
   * Resolves a relationship ID to a full file path.
   * @param id - Relationship ID
   * @param basePath - Base path of the source file
   */
  resolvePath(id: string, basePath: string): string | null;
}

/**
 * Parses a relationships XML file.
 *
 * @param xmlContent - Raw XML content of the .rels file
 * @returns Parsed relationship map
 *
 * @example
 * const relsXml = archive.getText('ppt/_rels/presentation.xml.rels');
 * const rels = parseRelationships(relsXml);
 * const slideRel = rels.get('rId2');
 */
export function parseRelationships(xmlContent: string): RelationshipMap {
  const doc = parseXml(xmlContent);
  const byId = new Map<string, Relationship>();
  const byType = new Map<string, Relationship[]>();

  // Find all Relationship elements
  const relationshipElements = findChildrenByName(doc.documentElement, 'Relationship');

  for (const el of relationshipElements) {
    const id = getAttribute(el, 'Id');
    const type = getAttribute(el, 'Type');
    const target = getAttribute(el, 'Target');
    const targetMode = getAttribute(el, 'TargetMode');

    if (!id || !type || !target) continue;

    const relationship: Relationship = {
      id,
      type,
      target,
      external: targetMode === 'External',
    };

    byId.set(id, relationship);

    // Group by type
    const typeKey = extractTypeKey(type);
    const typeList = byType.get(typeKey) || [];
    typeList.push(relationship);
    byType.set(typeKey, typeList);
  }

  return {
    byId,
    byType,

    get(id: string): Relationship | undefined {
      return byId.get(id);
    },

    getByType(type: string): Relationship[] {
      // Check for exact match first
      if (byType.has(type)) {
        return byType.get(type) || [];
      }

      // Check for partial match
      for (const [key, rels] of byType) {
        if (key.includes(type) || type.includes(key)) {
          return rels;
        }
      }

      return [];
    },

    resolvePath(id: string, basePath: string): string | null {
      const rel = byId.get(id);
      if (!rel || rel.external) return null;

      // Get the directory of the base path
      const baseDir = basePath.substring(0, basePath.lastIndexOf('/'));

      // Resolve the target path relative to base
      return resolvePath(baseDir, rel.target);
    },
  };
}

/**
 * Extracts a simple type key from a full relationship type URL.
 * E.g., "http://...relationships/slide" -> "slide"
 */
function extractTypeKey(fullType: string): string {
  const parts = fullType.split('/');
  return parts[parts.length - 1];
}

/**
 * Resolves a relative path against a base directory.
 */
function resolvePath(baseDir: string, relativePath: string): string {
  // Handle absolute paths
  if (relativePath.startsWith('/')) {
    return relativePath.slice(1);
  }

  // Split paths into parts
  const baseParts = baseDir.split('/').filter(Boolean);
  const relativeParts = relativePath.split('/');

  for (const part of relativeParts) {
    if (part === '..') {
      baseParts.pop();
    } else if (part !== '.') {
      baseParts.push(part);
    }
  }

  return baseParts.join('/');
}

/**
 * Checks if a relationship type matches a specific category.
 *
 * @param type - Full relationship type URL
 * @param category - Category to check (from RELATIONSHIP_TYPES)
 */
export function isRelationshipType(type: string, category: string): boolean {
  return type.endsWith(category);
}
