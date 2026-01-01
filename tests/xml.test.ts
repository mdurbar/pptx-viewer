import { describe, it, expect, beforeEach } from 'vitest';
import {
  parseXml,
  getLocalName,
  getAttribute,
  getNumberAttribute,
  getBooleanAttribute,
  findChildrenByName,
  findChildByName,
  findAllByName,
  findFirstByName,
  getTextContent,
  traversePath,
  NAMESPACES,
} from '../src/utils/xml';

describe('XML Utilities', () => {
  describe('parseXml', () => {
    it('parses valid XML', () => {
      const doc = parseXml('<root><child>text</child></root>');
      expect(doc.documentElement.nodeName).toBe('root');
    });

    it('throws on invalid XML', () => {
      expect(() => parseXml('<root><unclosed>')).toThrow('XML parsing error');
    });

    it('handles empty elements', () => {
      const doc = parseXml('<root/>');
      expect(doc.documentElement.nodeName).toBe('root');
    });
  });

  describe('getLocalName', () => {
    it('returns local name without prefix', () => {
      const doc = parseXml('<a:shape xmlns:a="http://example.com"/>');
      expect(getLocalName(doc.documentElement)).toBe('shape');
    });

    it('returns name for non-prefixed elements', () => {
      const doc = parseXml('<shape/>');
      expect(getLocalName(doc.documentElement)).toBe('shape');
    });
  });

  describe('getAttribute', () => {
    it('gets simple attribute', () => {
      const doc = parseXml('<root attr="value"/>');
      expect(getAttribute(doc.documentElement, 'attr')).toBe('value');
    });

    it('returns null for missing attribute', () => {
      const doc = parseXml('<root/>');
      expect(getAttribute(doc.documentElement, 'missing')).toBeNull();
    });

    it('handles namespaced attributes', () => {
      const doc = parseXml(
        '<root xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'
      );
      expect(getAttribute(doc.documentElement, 'r:id')).toBe('rId1');
    });
  });

  describe('getNumberAttribute', () => {
    it('parses integer attribute', () => {
      const doc = parseXml('<root count="42"/>');
      expect(getNumberAttribute(doc.documentElement, 'count')).toBe(42);
    });

    it('parses float attribute', () => {
      const doc = parseXml('<root value="3.14"/>');
      expect(getNumberAttribute(doc.documentElement, 'value')).toBeCloseTo(3.14);
    });

    it('returns default for missing attribute', () => {
      const doc = parseXml('<root/>');
      expect(getNumberAttribute(doc.documentElement, 'missing', 100)).toBe(100);
    });

    it('returns default for non-numeric attribute', () => {
      const doc = parseXml('<root value="abc"/>');
      expect(getNumberAttribute(doc.documentElement, 'value', 0)).toBe(0);
    });
  });

  describe('getBooleanAttribute', () => {
    it('parses "1" as true', () => {
      const doc = parseXml('<root enabled="1"/>');
      expect(getBooleanAttribute(doc.documentElement, 'enabled')).toBe(true);
    });

    it('parses "true" as true', () => {
      const doc = parseXml('<root enabled="true"/>');
      expect(getBooleanAttribute(doc.documentElement, 'enabled')).toBe(true);
    });

    it('parses "0" as false', () => {
      const doc = parseXml('<root enabled="0"/>');
      expect(getBooleanAttribute(doc.documentElement, 'enabled')).toBe(false);
    });

    it('returns default for missing attribute', () => {
      const doc = parseXml('<root/>');
      expect(getBooleanAttribute(doc.documentElement, 'enabled', true)).toBe(true);
    });
  });

  describe('findChildrenByName', () => {
    it('finds all matching children', () => {
      const doc = parseXml('<root><item/><item/><other/></root>');
      const items = findChildrenByName(doc.documentElement, 'item');
      expect(items).toHaveLength(2);
    });

    it('returns empty array when no matches', () => {
      const doc = parseXml('<root><other/></root>');
      const items = findChildrenByName(doc.documentElement, 'item');
      expect(items).toHaveLength(0);
    });

    it('only searches direct children', () => {
      const doc = parseXml('<root><parent><item/></parent></root>');
      const items = findChildrenByName(doc.documentElement, 'item');
      expect(items).toHaveLength(0);
    });
  });

  describe('findChildByName', () => {
    it('finds first matching child', () => {
      const doc = parseXml('<root><item id="1"/><item id="2"/></root>');
      const item = findChildByName(doc.documentElement, 'item');
      expect(item).not.toBeNull();
      expect(getAttribute(item!, 'id')).toBe('1');
    });

    it('returns null when no match', () => {
      const doc = parseXml('<root><other/></root>');
      expect(findChildByName(doc.documentElement, 'item')).toBeNull();
    });
  });

  describe('findAllByName', () => {
    it('finds all descendants with matching name', () => {
      const doc = parseXml(`
        <root>
          <item/>
          <parent>
            <item/>
            <child><item/></child>
          </parent>
        </root>
      `);
      const items = findAllByName(doc.documentElement, 'item');
      expect(items).toHaveLength(3);
    });
  });

  describe('findFirstByName', () => {
    it('finds first descendant with matching name', () => {
      const doc = parseXml(`
        <root>
          <parent>
            <item id="1"/>
          </parent>
          <item id="2"/>
        </root>
      `);
      const item = findFirstByName(doc.documentElement, 'item');
      expect(item).not.toBeNull();
      expect(getAttribute(item!, 'id')).toBe('1');
    });

    it('returns null when no match', () => {
      const doc = parseXml('<root><other/></root>');
      expect(findFirstByName(doc.documentElement, 'item')).toBeNull();
    });
  });

  describe('getTextContent', () => {
    it('gets text content from element', () => {
      const doc = parseXml('<root>Hello World</root>');
      expect(getTextContent(doc.documentElement)).toBe('Hello World');
    });

    it('trims whitespace', () => {
      const doc = parseXml('<root>  trimmed  </root>');
      expect(getTextContent(doc.documentElement)).toBe('trimmed');
    });

    it('returns empty string for empty element', () => {
      const doc = parseXml('<root/>');
      expect(getTextContent(doc.documentElement)).toBe('');
    });
  });

  describe('traversePath', () => {
    it('traverses path to nested element', () => {
      const doc = parseXml(`
        <root>
          <a>
            <b>
              <c id="target"/>
            </b>
          </a>
        </root>
      `);
      const result = traversePath(doc.documentElement, ['a', 'b', 'c']);
      expect(result).not.toBeNull();
      expect(getAttribute(result!, 'id')).toBe('target');
    });

    it('returns null for invalid path', () => {
      const doc = parseXml('<root><a/></root>');
      expect(traversePath(doc.documentElement, ['a', 'b', 'c'])).toBeNull();
    });

    it('handles empty path', () => {
      const doc = parseXml('<root/>');
      expect(traversePath(doc.documentElement, [])).toBe(doc.documentElement);
    });
  });

  describe('NAMESPACES', () => {
    it('contains DrawingML namespace', () => {
      expect(NAMESPACES.a).toBe('http://schemas.openxmlformats.org/drawingml/2006/main');
    });

    it('contains Relationships namespace', () => {
      expect(NAMESPACES.r).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships');
    });

    it('contains PresentationML namespace', () => {
      expect(NAMESPACES.p).toBe('http://schemas.openxmlformats.org/presentationml/2006/main');
    });
  });
});
