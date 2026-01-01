/**
 * XML parsing utilities for OOXML documents.
 *
 * OOXML uses multiple namespaces. Common prefixes:
 * - a: DrawingML (http://schemas.openxmlformats.org/drawingml/2006/main)
 * - r: Relationships (http://schemas.openxmlformats.org/officeDocument/2006/relationships)
 * - p: PresentationML (http://schemas.openxmlformats.org/presentationml/2006/main)
 */

/**
 * Common OOXML namespace URIs.
 */
export const NAMESPACES = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  cp: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
} as const;

/**
 * Parses an XML string into a Document.
 *
 * @param xmlString - The XML content to parse
 * @returns Parsed XML Document
 * @throws Error if parsing fails
 */
export function parseXml(xmlString: string): Document {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlString, 'application/xml');

  // Check for parsing errors
  const parserError = doc.querySelector('parsererror');
  if (parserError) {
    throw new Error(`XML parsing error: ${parserError.textContent}`);
  }

  return doc;
}

/**
 * Gets the local name of an element (without namespace prefix).
 *
 * @param element - The element to get the name from
 * @returns Local name without prefix
 */
export function getLocalName(element: Element): string {
  return element.localName || element.nodeName.split(':').pop() || '';
}

/**
 * Gets an attribute value from an element, checking common prefixes.
 *
 * @param element - The element to get the attribute from
 * @param name - Attribute name (can include prefix like "r:embed")
 * @returns Attribute value or null if not found
 */
export function getAttribute(element: Element, name: string): string | null {
  // Try direct attribute first
  let value = element.getAttribute(name);
  if (value !== null) return value;

  // Handle namespaced attributes
  if (name.includes(':')) {
    const [prefix, localName] = name.split(':');
    const ns = NAMESPACES[prefix as keyof typeof NAMESPACES];
    if (ns) {
      value = element.getAttributeNS(ns, localName);
      if (value !== null) return value;
    }
  }

  return null;
}

/**
 * Gets an attribute value as a number.
 *
 * @param element - The element to get the attribute from
 * @param name - Attribute name
 * @param defaultValue - Default value if attribute is missing or invalid
 * @returns Parsed number or default value
 */
export function getNumberAttribute(
  element: Element,
  name: string,
  defaultValue: number = 0
): number {
  const value = getAttribute(element, name);
  if (value === null) return defaultValue;

  const parsed = parseFloat(value);
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * Gets an attribute value as a boolean.
 *
 * @param element - The element to get the attribute from
 * @param name - Attribute name
 * @param defaultValue - Default value if attribute is missing
 * @returns Boolean value
 */
export function getBooleanAttribute(
  element: Element,
  name: string,
  defaultValue: boolean = false
): boolean {
  const value = getAttribute(element, name);
  if (value === null) return defaultValue;

  return value === '1' || value === 'true';
}

/**
 * Finds all child elements with a specific local name.
 *
 * @param parent - Parent element to search in
 * @param localName - Local name to find (without namespace prefix)
 * @returns Array of matching elements
 */
export function findChildrenByName(parent: Element, localName: string): Element[] {
  const results: Element[] = [];
  for (const child of Array.from(parent.children)) {
    if (getLocalName(child) === localName) {
      results.push(child);
    }
  }
  return results;
}

/**
 * Finds the first child element with a specific local name.
 *
 * @param parent - Parent element to search in
 * @param localName - Local name to find (without namespace prefix)
 * @returns Matching element or null
 */
export function findChildByName(parent: Element, localName: string): Element | null {
  for (const child of Array.from(parent.children)) {
    if (getLocalName(child) === localName) {
      return child;
    }
  }
  return null;
}

/**
 * Recursively finds all descendants with a specific local name.
 *
 * @param parent - Parent element to search in
 * @param localName - Local name to find
 * @returns Array of matching elements
 */
export function findAllByName(parent: Element, localName: string): Element[] {
  const results: Element[] = [];

  function search(element: Element) {
    for (const child of Array.from(element.children)) {
      if (getLocalName(child) === localName) {
        results.push(child);
      }
      search(child);
    }
  }

  search(parent);
  return results;
}

/**
 * Finds the first descendant with a specific local name.
 *
 * @param parent - Parent element to search in
 * @param localName - Local name to find
 * @returns Matching element or null
 */
export function findFirstByName(parent: Element, localName: string): Element | null {
  function search(element: Element): Element | null {
    for (const child of Array.from(element.children)) {
      if (getLocalName(child) === localName) {
        return child;
      }
      const found = search(child);
      if (found) return found;
    }
    return null;
  }

  return search(parent);
}

/**
 * Gets the text content of an element, trimmed.
 *
 * @param element - Element to get text from
 * @returns Trimmed text content
 */
export function getTextContent(element: Element): string {
  return element.textContent?.trim() || '';
}

/**
 * Creates a simple path through child elements.
 * Useful for navigating deep structures.
 *
 * @param element - Starting element
 * @param path - Array of local names to traverse
 * @returns Final element or null if path doesn't exist
 *
 * @example
 * traversePath(spPr, ['xfrm', 'off']); // Gets spPr/xfrm/off
 */
export function traversePath(element: Element, path: string[]): Element | null {
  let current: Element | null = element;

  for (const name of path) {
    if (!current) return null;
    current = findChildByName(current, name);
  }

  return current;
}
