import { describe, it, expect } from 'vitest';
import { parseDiagram } from '../src/parser/DiagramParser';
import type { ThemeColors } from '../src/core/types';
import type { PPTXArchive } from '../src/core/unzip';
import { parseRelationships } from '../src/parser/RelationshipParser';

const mockTheme: ThemeColors = {
  dark1: '#000000',
  dark2: '#444444',
  light1: '#FFFFFF',
  light2: '#EEEEEE',
  accent1: '#FF0000',
  accent2: '#00FF00',
  accent3: '#0000FF',
  accent4: '#FFFF00',
  accent5: '#FF00FF',
  accent6: '#00FFFF',
  hlink: '#0000CC',
  folHlink: '#660066',
};

const mockRelationships = parseRelationships(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>`);

// Helper to create a mock archive with diagram XML
function createMockArchive(drawingXml: string, relsXml?: string): PPTXArchive {
  return {
    files: new Map(),
    getText: (path: string) => {
      if (path === 'ppt/diagrams/drawing1.xml') {
        return drawingXml;
      }
      if (path === 'ppt/diagrams/_rels/drawing1.xml.rels' && relsXml) {
        return relsXml;
      }
      return null;
    },
    getBytes: () => null,
    getBlobUrl: () => null,
    listFiles: () => [],
    hasFile: () => false,
    cleanup: () => {},
  };
}

describe('DiagramParser', () => {
  describe('parseDiagram', () => {
    it('returns null for missing diagram drawing file', () => {
      const archive = createMockArchive('');
      archive.getText = () => null;

      const result = parseDiagram(
        'ppt/diagrams/drawing1.xml',
        archive,
        mockTheme,
        mockRelationships,
        'ppt/slides/slide1.xml'
      );
      expect(result).toBeNull();
    });

    it('returns null for invalid XML', () => {
      const archive = createMockArchive('<invalid');

      const result = parseDiagram(
        'ppt/diagrams/drawing1.xml',
        archive,
        mockTheme,
        mockRelationships,
        'ppt/slides/slide1.xml'
      );
      expect(result).toBeNull();
    });

    it('returns null when no spTree element found', () => {
      const drawingXml = `
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram">
          <dsp:noSpTree/>
        </dsp:drawing>
      `;
      const archive = createMockArchive(drawingXml);

      const result = parseDiagram(
        'ppt/diagrams/drawing1.xml',
        archive,
        mockTheme,
        mockRelationships,
        'ppt/slides/slide1.xml'
      );
      expect(result).toBeNull();
    });

    it('parses a diagram with a shape tree containing shapes', () => {
      const drawingXml = `
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dsp:spTree>
            <dsp:sp>
              <dsp:nvSpPr>
                <dsp:cNvPr id="1" name="Shape 1"/>
                <dsp:cNvSpPr/>
              </dsp:nvSpPr>
              <dsp:spPr>
                <a:xfrm>
                  <a:off x="100000" y="100000"/>
                  <a:ext cx="500000" cy="300000"/>
                </a:xfrm>
                <a:prstGeom prst="rect"/>
                <a:solidFill>
                  <a:srgbClr val="FF0000"/>
                </a:solidFill>
              </dsp:spPr>
            </dsp:sp>
          </dsp:spTree>
        </dsp:drawing>
      `;
      const archive = createMockArchive(drawingXml);

      const result = parseDiagram(
        'ppt/diagrams/drawing1.xml',
        archive,
        mockTheme,
        mockRelationships,
        'ppt/slides/slide1.xml'
      );

      expect(result).not.toBeNull();
      expect(result?.children).toBeInstanceOf(Array);
      // The shape should be parsed (has spPr with xfrm, off, ext, and prstGeom)
      expect(result?.children.length).toBeGreaterThanOrEqual(1);
    });

    it('returns result with empty children for shape tree with no parseable shapes', () => {
      // A valid spTree but shapes without required properties
      const drawingXml = `
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram">
          <dsp:spTree>
            <dsp:sp>
              <dsp:nvSpPr>
                <dsp:cNvPr id="1" name="Empty Shape"/>
              </dsp:nvSpPr>
            </dsp:sp>
          </dsp:spTree>
        </dsp:drawing>
      `;
      const archive = createMockArchive(drawingXml);

      const result = parseDiagram(
        'ppt/diagrams/drawing1.xml',
        archive,
        mockTheme,
        mockRelationships,
        'ppt/slides/slide1.xml'
      );

      // Should return a result but with empty children (shapes couldn't be parsed)
      expect(result).not.toBeNull();
      expect(result?.children).toEqual([]);
    });

    it('uses diagram-specific relationships when available', () => {
      const drawingXml = `
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dsp:spTree>
            <dsp:sp>
              <dsp:nvSpPr>
                <dsp:cNvPr id="1" name="Shape 1"/>
                <dsp:cNvSpPr/>
              </dsp:nvSpPr>
              <dsp:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="100000" cy="100000"/>
                </a:xfrm>
                <a:prstGeom prst="rect"/>
              </dsp:spPr>
            </dsp:sp>
          </dsp:spTree>
        </dsp:drawing>
      `;

      const diagramRelsXml = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/diagram_image.png"/>
</Relationships>`;

      const archive = createMockArchive(drawingXml, diagramRelsXml);

      const result = parseDiagram(
        'ppt/diagrams/drawing1.xml',
        archive,
        mockTheme,
        mockRelationships,
        'ppt/slides/slide1.xml'
      );

      expect(result).not.toBeNull();
    });

    it('extracts diagram type from root element name attribute', () => {
      const drawingXml = `
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                     name="Process Diagram"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dsp:spTree>
            <dsp:sp>
              <dsp:nvSpPr>
                <dsp:cNvPr id="1" name="Shape"/>
                <dsp:cNvSpPr/>
              </dsp:nvSpPr>
              <dsp:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="100000" cy="100000"/>
                </a:xfrm>
                <a:prstGeom prst="rect"/>
              </dsp:spPr>
            </dsp:sp>
          </dsp:spTree>
        </dsp:drawing>
      `;
      const archive = createMockArchive(drawingXml);

      const result = parseDiagram(
        'ppt/diagrams/drawing1.xml',
        archive,
        mockTheme,
        mockRelationships,
        'ppt/slides/slide1.xml'
      );

      expect(result).not.toBeNull();
      expect(result?.diagramType).toBe('Process Diagram');
    });

    it('parses multiple shapes in diagram', () => {
      const drawingXml = `
        <dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"
                     xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <dsp:spTree>
            <dsp:sp>
              <dsp:nvSpPr>
                <dsp:cNvPr id="1" name="Shape 1"/>
                <dsp:cNvSpPr/>
              </dsp:nvSpPr>
              <dsp:spPr>
                <a:xfrm>
                  <a:off x="0" y="0"/>
                  <a:ext cx="100000" cy="100000"/>
                </a:xfrm>
                <a:prstGeom prst="rect"/>
              </dsp:spPr>
            </dsp:sp>
            <dsp:sp>
              <dsp:nvSpPr>
                <dsp:cNvPr id="2" name="Shape 2"/>
                <dsp:cNvSpPr/>
              </dsp:nvSpPr>
              <dsp:spPr>
                <a:xfrm>
                  <a:off x="200000" y="0"/>
                  <a:ext cx="100000" cy="100000"/>
                </a:xfrm>
                <a:prstGeom prst="ellipse"/>
              </dsp:spPr>
            </dsp:sp>
            <dsp:sp>
              <dsp:nvSpPr>
                <dsp:cNvPr id="3" name="Shape 3"/>
                <dsp:cNvSpPr/>
              </dsp:nvSpPr>
              <dsp:spPr>
                <a:xfrm>
                  <a:off x="400000" y="0"/>
                  <a:ext cx="100000" cy="100000"/>
                </a:xfrm>
                <a:prstGeom prst="triangle"/>
              </dsp:spPr>
            </dsp:sp>
          </dsp:spTree>
        </dsp:drawing>
      `;
      const archive = createMockArchive(drawingXml);

      const result = parseDiagram(
        'ppt/diagrams/drawing1.xml',
        archive,
        mockTheme,
        mockRelationships,
        'ppt/slides/slide1.xml'
      );

      expect(result).not.toBeNull();
      expect(result?.children.length).toBe(3);
    });
  });
});
