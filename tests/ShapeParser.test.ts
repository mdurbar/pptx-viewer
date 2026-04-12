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

  describe('placeholder inheritance', () => {
    // A slide-level placeholder shape with empty spPr (no xfrm). This is the
    // exact shape reported in the bug — agent-generated PPTX files lean on
    // layout/master inheritance for bounds.
    const placeholderSlideXml = `
      <spTree>
        <sp>
          <nvSpPr>
            <cNvPr id="100" name="Title"/>
            <cNvSpPr txBox="1"/>
            <nvPr><ph type="ctrTitle"/></nvPr>
          </nvSpPr>
          <spPr>
            <prstGeom prst="rect"><avLst/></prstGeom>
          </spPr>
          <txBody>
            <bodyPr/><lstStyle/>
            <p><r><t>Northstar Labs</t></r></p>
          </txBody>
        </sp>
      </spTree>
    `;

    function makeTextElement(
      phType: any,
      idx: number | undefined,
      x: number,
      y: number,
      w: number,
      h: number
    ): any {
      return {
        id: `ph-${phType}-${idx ?? 'noidx'}`,
        type: 'text',
        bounds: { x, y, width: w, height: h },
        placeholder: idx !== undefined ? { type: phType, idx } : { type: phType },
        text: { paragraphs: [] },
      };
    }

    function makeLayout(elements: any[]): any {
      return {
        id: 'rId1',
        masterId: 'rId1',
        elements,
        showMasterShapes: true,
      };
    }

    function makeMaster(elements: any[]): any {
      return {
        id: 'rId1',
        elements,
        colorMap: {},
        layoutIds: [],
      };
    }

    it('inherits bounds from layout placeholder (type match)', () => {
      const layout = makeLayout([makeTextElement('ctrTitle', undefined, 100, 50, 800, 200)]);
      const xml = parseXml(placeholderSlideXml);
      const elements = parseShapeTree(xml.documentElement, { ...mockContext, layout });

      expect(elements).toHaveLength(1);
      expect(elements[0].bounds).toEqual({ x: 100, y: 50, width: 800, height: 200 });
    });

    it('inherits bounds from master when layout has no matching placeholder', () => {
      const layout = makeLayout([makeTextElement('body', 1, 10, 10, 20, 20)]);
      const master = makeMaster([makeTextElement('ctrTitle', undefined, 42, 24, 600, 100)]);
      const xml = parseXml(placeholderSlideXml);
      const elements = parseShapeTree(xml.documentElement, { ...mockContext, layout, master });

      expect(elements).toHaveLength(1);
      expect(elements[0].bounds).toEqual({ x: 42, y: 24, width: 600, height: 100 });
    });

    it('drops the shape when neither layout nor master has matching bounds', () => {
      const layout = makeLayout([makeTextElement('body', 1, 10, 10, 20, 20)]);
      const master = makeMaster([makeTextElement('ftr', undefined, 0, 700, 960, 40)]);
      const xml = parseXml(placeholderSlideXml);
      const elements = parseShapeTree(xml.documentElement, { ...mockContext, layout, master });

      expect(elements).toHaveLength(0);
    });

    it('matches ctrTitle on the slide against a layout title placeholder', () => {
      // Slide has ctrTitle, layout only has plain title — they should be
      // considered equivalent per ECMA-376 §19.3.1.36.
      const layout = makeLayout([makeTextElement('title', undefined, 50, 50, 500, 80)]);
      const xml = parseXml(placeholderSlideXml);
      const elements = parseShapeTree(xml.documentElement, { ...mockContext, layout });

      expect(elements).toHaveLength(1);
      expect(elements[0].bounds).toEqual({ x: 50, y: 50, width: 500, height: 80 });
    });

    it('matches ctrTitle on the layout against a slide title placeholder', () => {
      // Reverse: slide has plain title, layout has ctrTitle — still a match.
      const layout = makeLayout([makeTextElement('ctrTitle', undefined, 12, 34, 300, 60)]);
      const slideXml = `
        <spTree>
          <sp>
            <nvSpPr>
              <cNvPr id="101" name="Title"/>
              <nvPr><ph type="title"/></nvPr>
            </nvSpPr>
            <spPr><prstGeom prst="rect"/></spPr>
            <txBody><p><r><t>Hello</t></r></p></txBody>
          </sp>
        </spTree>
      `;
      const xml = parseXml(slideXml);
      const elements = parseShapeTree(xml.documentElement, { ...mockContext, layout });

      expect(elements).toHaveLength(1);
      expect(elements[0].bounds).toEqual({ x: 12, y: 34, width: 300, height: 60 });
    });

    it('prefers idx match over type match', () => {
      // Layout has two body placeholders — idx=1 and idx=2. Slide requests
      // idx=2, so we must get the idx=2 bounds (not the idx=1 bounds even
      // though both type-match).
      const layout = makeLayout([
        makeTextElement('body', 1, 100, 100, 200, 200),
        makeTextElement('body', 2, 500, 500, 300, 300),
      ]);
      const slideXml = `
        <spTree>
          <sp>
            <nvSpPr>
              <cNvPr id="102" name="Body"/>
              <nvPr><ph type="body" idx="2"/></nvPr>
            </nvSpPr>
            <spPr><prstGeom prst="rect"/></spPr>
            <txBody><p><r><t>Point</t></r></p></txBody>
          </sp>
        </spTree>
      `;
      const xml = parseXml(slideXml);
      const elements = parseShapeTree(xml.documentElement, { ...mockContext, layout });

      expect(elements).toHaveLength(1);
      expect(elements[0].bounds).toEqual({ x: 500, y: 500, width: 300, height: 300 });
    });

    it('does not inherit when no placeholder is declared on the slide shape', () => {
      // Shape has no <ph> element at all — not a placeholder, so no layout
      // lookup should happen. It should be dropped as before.
      const layout = makeLayout([makeTextElement('ctrTitle', undefined, 100, 50, 800, 200)]);
      const slideXml = `
        <spTree>
          <sp>
            <nvSpPr><cNvPr id="103" name="Plain"/></nvSpPr>
            <spPr><prstGeom prst="rect"/></spPr>
          </sp>
        </spTree>
      `;
      const xml = parseXml(slideXml);
      const elements = parseShapeTree(xml.documentElement, { ...mockContext, layout });

      expect(elements).toHaveLength(0);
    });
  });
});
