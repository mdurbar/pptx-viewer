import { describe, it, expect } from 'vitest';
import { parseShapeTree } from '../src/parser/ShapeParser';
import { parseXml } from '../src/utils/xml';
import type { ThemeColors } from '../src/core/types';
import type { PPTXArchive } from '../src/core/unzip';

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

// Minimal mock archive for testing
const mockArchive: PPTXArchive = {
  files: new Map(),
  getText: async () => '',
  getBytes: async () => new Uint8Array(),
  getBlobUrl: async () => '',
  cleanup: () => {},
};

const mockContext = {
  themeColors: mockTheme,
  relationships: new Map(),
  archive: mockArchive,
  basePath: 'ppt/slides',
};

describe('ShapeParser', () => {
  describe('parseShapeTree', () => {
    it('parses empty shape tree', () => {
      const xml = parseXml('<spTree/>');
      const elements = parseShapeTree(xml.documentElement, mockContext);

      expect(elements).toHaveLength(0);
    });

    it('parses basic rectangle shape', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr>
              <cNvPr id="2" name="Rectangle 1"/>
            </nvSpPr>
            <spPr>
              <xfrm>
                <off x="914400" y="914400"/>
                <ext cx="1828800" cy="914400"/>
              </xfrm>
              <prstGeom prst="rect"/>
              <solidFill>
                <srgbClr val="FF0000"/>
              </solidFill>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      expect(elements).toHaveLength(1);
      expect(elements[0].type).toBe('shape');

      const shape = elements[0] as any;
      expect(shape.shapeType).toBe('rect');
      expect(shape.bounds.x).toBeCloseTo(96);  // 914400 EMU = 1 inch = 96px
      expect(shape.bounds.y).toBeCloseTo(96);
      expect(shape.bounds.width).toBeCloseTo(192);  // 1828800 EMU = 2 inches = 192px
      expect(shape.bounds.height).toBeCloseTo(96);
    });

    it('parses ellipse shape', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="3" name="Ellipse"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="0" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="ellipse"/>
              <solidFill><srgbClr val="00FF00"/></solidFill>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      expect(elements).toHaveLength(1);
      const shape = elements[0] as any;
      expect(shape.shapeType).toBe('ellipse');
    });

    it('parses rounded rectangle', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="4" name="Rounded Rectangle"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="0" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="roundRect"/>
              <solidFill><srgbClr val="0000FF"/></solidFill>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      const shape = elements[0] as any;
      expect(shape.shapeType).toBe('roundRect');
    });

    it('parses shape with rotation', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="5" name="Rotated"/></nvSpPr>
            <spPr>
              <xfrm rot="5400000">
                <off x="0" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="rect"/>
              <solidFill><srgbClr val="FF0000"/></solidFill>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      const shape = elements[0] as any;
      expect(shape.rotation).toBe(90);  // 5400000 / 60000 = 90 degrees
    });

    it('parses shape with solid fill', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="6" name="Filled"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="0" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="rect"/>
              <solidFill>
                <srgbClr val="AABBCC"/>
              </solidFill>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      const shape = elements[0] as any;
      expect(shape.fill.type).toBe('solid');
      expect(shape.fill.color.hex).toBe('#AABBCC');
    });

    it('parses shape with theme color fill', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="7" name="Theme Filled"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="0" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="rect"/>
              <solidFill>
                <schemeClr val="accent1"/>
              </solidFill>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      const shape = elements[0] as any;
      expect(shape.fill.type).toBe('solid');
      expect(shape.fill.color.hex).toBe('#FF0000');  // accent1 from mockTheme
    });

    it('parses shape with stroke', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="8" name="Stroked"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="0" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="rect"/>
              <noFill/>
              <ln w="25400">
                <solidFill><srgbClr val="000000"/></solidFill>
              </ln>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      const shape = elements[0] as any;
      expect(shape.stroke).toBeDefined();
      expect(shape.stroke.width).toBeCloseTo(2.67, 1);  // 25400 EMU ≈ 2.67px
      expect(shape.stroke.color.hex).toBe('#000000');
    });

    it('parses text box as TextElement when no visible fill/stroke', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="9" name="Text Box"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="0" y="0"/>
                <ext cx="914400" cy="457200"/>
              </xfrm>
              <prstGeom prst="rect"/>
              <noFill/>
            </spPr>
            <txBody>
              <bodyPr/>
              <p><r><t>Hello</t></r></p>
            </txBody>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      expect(elements).toHaveLength(1);
      expect(elements[0].type).toBe('text');
    });

    it('parses arrow shapes', () => {
      const arrowTypes = ['rightArrow', 'leftArrow', 'upArrow', 'downArrow'];

      for (const arrowType of arrowTypes) {
        const xml = parseXml(`
          <spTree>
            <sp>
              <nvSpPr><cNvPr id="10" name="Arrow"/></nvSpPr>
              <spPr>
                <xfrm>
                  <off x="0" y="0"/>
                  <ext cx="914400" cy="914400"/>
                </xfrm>
                <prstGeom prst="${arrowType}"/>
                <solidFill><srgbClr val="FF0000"/></solidFill>
              </spPr>
            </sp>
          </spTree>
        `);
        const elements = parseShapeTree(xml.documentElement, mockContext);

        const shape = elements[0] as any;
        expect(shape.shapeType).toBe(arrowType);
      }
    });

    it('parses star shapes', () => {
      const starTypes = ['star4', 'star5', 'star6'];

      for (const starType of starTypes) {
        const xml = parseXml(`
          <spTree>
            <sp>
              <nvSpPr><cNvPr id="11" name="Star"/></nvSpPr>
              <spPr>
                <xfrm>
                  <off x="0" y="0"/>
                  <ext cx="914400" cy="914400"/>
                </xfrm>
                <prstGeom prst="${starType}"/>
                <solidFill><srgbClr val="FFD700"/></solidFill>
              </spPr>
            </sp>
          </spTree>
        `);
        const elements = parseShapeTree(xml.documentElement, mockContext);

        const shape = elements[0] as any;
        expect(shape.shapeType).toBe(starType);
      }
    });

    it('parses multiple shapes', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="1" name="Shape 1"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="0" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="rect"/>
              <solidFill><srgbClr val="FF0000"/></solidFill>
            </spPr>
          </sp>
          <sp>
            <nvSpPr><cNvPr id="2" name="Shape 2"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="914400" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="ellipse"/>
              <solidFill><srgbClr val="00FF00"/></solidFill>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      expect(elements).toHaveLength(2);
      expect((elements[0] as any).shapeType).toBe('rect');
      expect((elements[1] as any).shapeType).toBe('ellipse');
    });

    it('parses shape with gradient fill', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="12" name="Gradient"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="0" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="rect"/>
              <gradFill>
                <gsLst>
                  <gs pos="0"><srgbClr val="FF0000"/></gs>
                  <gs pos="100000"><srgbClr val="0000FF"/></gs>
                </gsLst>
                <lin ang="5400000"/>
              </gradFill>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      const shape = elements[0] as any;
      expect(shape.fill.type).toBe('gradient');
      expect(shape.fill.stops).toHaveLength(2);
      expect(shape.fill.stops[0].color.hex).toBe('#FF0000');
      expect(shape.fill.stops[1].color.hex).toBe('#0000FF');
    });

    it('parses shape with noFill', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="13" name="No Fill"/></nvSpPr>
            <spPr>
              <xfrm>
                <off x="0" y="0"/>
                <ext cx="914400" cy="914400"/>
              </xfrm>
              <prstGeom prst="rect"/>
              <noFill/>
              <ln w="25400">
                <solidFill><srgbClr val="000000"/></solidFill>
              </ln>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      const shape = elements[0] as any;
      expect(shape.fill.type).toBe('none');
    });

    it('skips shape without spPr', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="14" name="Bad Shape"/></nvSpPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      expect(elements).toHaveLength(0);
    });

    it('skips shape without bounds', () => {
      const xml = parseXml(`
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="15" name="No Bounds"/></nvSpPr>
            <spPr>
              <prstGeom prst="rect"/>
            </spPr>
          </sp>
        </spTree>
      `);
      const elements = parseShapeTree(xml.documentElement, mockContext);

      expect(elements).toHaveLength(0);
    });
  });
});
