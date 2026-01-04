import { describe, it, expect } from 'vitest';
import { parseTextBody, parseColorElement } from '../src/parser/TextParser';
import { parseXml } from '../src/utils/xml';
import type { ThemeColors } from '../src/core/types';

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

describe('TextParser', () => {
  describe('parseTextBody', () => {
    it('parses empty text body', () => {
      const xml = parseXml('<txBody><bodyPr/></txBody>');
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs).toHaveLength(0);
      expect(result.verticalAlign).toBe('top');
    });

    it('parses single paragraph with text', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r><t>Hello World</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs).toHaveLength(1);
      expect(result.paragraphs[0].runs).toHaveLength(1);
      expect(result.paragraphs[0].runs[0].text).toBe('Hello World');
    });

    it('parses multiple paragraphs', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p><r><t>First</t></r></p>
          <p><r><t>Second</t></r></p>
          <p><r><t>Third</t></r></p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs).toHaveLength(3);
      expect(result.paragraphs[0].runs[0].text).toBe('First');
      expect(result.paragraphs[1].runs[0].text).toBe('Second');
      expect(result.paragraphs[2].runs[0].text).toBe('Third');
    });

    it('parses vertical alignment - center', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr anchor="ctr"/>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.verticalAlign).toBe('middle');
    });

    it('parses vertical alignment - bottom', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr anchor="b"/>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.verticalAlign).toBe('bottom');
    });

    it('parses body padding/insets', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr tIns="91440" rIns="182880" bIns="91440" lIns="182880"/>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      // 91440 EMU = 0.1 inch = 9.6 pixels at 96 DPI
      // 182880 EMU = 0.2 inch = 19.2 pixels at 96 DPI
      expect(result.padding?.top).toBeCloseTo(9.6);
      expect(result.padding?.right).toBeCloseTo(19.2);
      expect(result.padding?.bottom).toBeCloseTo(9.6);
      expect(result.padding?.left).toBeCloseTo(19.2);
    });

    it('parses paragraph alignment', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p><pPr algn="ctr"/><r><t>Centered</t></r></p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].align).toBe('center');
    });

    it('parses bold text', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r><rPr b="1"/><t>Bold</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].runs[0].bold).toBe(true);
    });

    it('parses italic text', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r><rPr i="1"/><t>Italic</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].runs[0].italic).toBe(true);
    });

    it('parses underline text', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r><rPr u="sng"/><t>Underlined</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].runs[0].underline).toBe(true);
    });

    it('parses strikethrough text', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r><rPr strike="sngStrike"/><t>Struck</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].runs[0].strikethrough).toBe(true);
    });

    it('parses font size', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r><rPr sz="2400"/><t>24pt text</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      // 2400 centipoints = 24 points = 32 pixels at 96 DPI
      expect(result.paragraphs[0].runs[0].fontSize).toBeCloseTo(32);
    });

    it('parses font family', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r><rPr><latin typeface="Arial"/></rPr><t>Arial text</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].runs[0].fontFamily).toBe('Arial');
    });

    it('parses bullet points', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <pPr><buChar char="•"/></pPr>
            <r><t>Bullet item</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].bullet).toBeDefined();
      expect(result.paragraphs[0].bullet?.type).toBe('bullet');
      expect(result.paragraphs[0].bullet?.char).toBe('•');
    });

    it('parses numbered lists', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <pPr><buAutoNum type="arabicPeriod" startAt="1"/></pPr>
            <r><t>Numbered item</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].bullet).toBeDefined();
      expect(result.paragraphs[0].bullet?.type).toBe('number');
      expect(result.paragraphs[0].bullet?.startAt).toBe(1);
    });

    it('parses indentation level', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <pPr lvl="2"/>
            <r><t>Nested item</t></r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].level).toBe(2);
    });

    it('handles field elements', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <fld type="slidenum"><t>1</t></fld>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].runs).toHaveLength(1);
      expect(result.paragraphs[0].runs[0].text).toBe('1');
    });

    it('parses text highlight color', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r>
              <rPr>
                <highlight><srgbClr val="FFFF00"/></highlight>
              </rPr>
              <t>Highlighted text</t>
            </r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].runs[0].highlight).toBeDefined();
      expect(result.paragraphs[0].runs[0].highlight?.hex).toBe('#FFFF00');
    });

    it('parses text glow effect', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r>
              <rPr>
                <effectLst>
                  <glow rad="101600">
                    <srgbClr val="FF0000"/>
                  </glow>
                </effectLst>
              </rPr>
              <t>Glowing text</t>
            </r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].runs[0].glow).toBeDefined();
      expect(result.paragraphs[0].runs[0].glow?.color.hex).toBe('#FF0000');
      expect(result.paragraphs[0].runs[0].glow?.radius).toBeGreaterThan(0);
    });

    it('parses text reflection effect', () => {
      const xml = parseXml(`
        <txBody>
          <bodyPr/>
          <p>
            <r>
              <rPr>
                <effectLst>
                  <reflection blurRad="12700" stA="50000" endA="0" dist="25400" dir="5400000"/>
                </effectLst>
              </rPr>
              <t>Reflected text</t>
            </r>
          </p>
        </txBody>
      `);
      const result = parseTextBody(xml.documentElement, mockTheme);

      expect(result.paragraphs[0].runs[0].reflection).toBeDefined();
      expect(result.paragraphs[0].runs[0].reflection?.startOpacity).toBe(0.5);
      expect(result.paragraphs[0].runs[0].reflection?.endOpacity).toBe(0);
      expect(result.paragraphs[0].runs[0].reflection?.direction).toBe(90); // 5400000 / 60000 = 90
    });
  });

  describe('parseColorElement', () => {
    it('parses srgbClr (RGB hex)', () => {
      const xml = parseXml(`
        <solidFill>
          <srgbClr val="FF5500"/>
        </solidFill>
      `);
      const color = parseColorElement(xml.documentElement, mockTheme);

      expect(color).not.toBeNull();
      expect(color?.hex).toBe('#FF5500');
      expect(color?.alpha).toBe(1);
    });

    it('parses srgbClr with alpha', () => {
      const xml = parseXml(`
        <solidFill>
          <srgbClr val="FF0000">
            <alpha val="50000"/>
          </srgbClr>
        </solidFill>
      `);
      const color = parseColorElement(xml.documentElement, mockTheme);

      expect(color?.hex).toBe('#FF0000');
      expect(color?.alpha).toBe(0.5);
    });

    it('parses schemeClr (theme color)', () => {
      const xml = parseXml(`
        <solidFill>
          <schemeClr val="accent1"/>
        </solidFill>
      `);
      const color = parseColorElement(xml.documentElement, mockTheme);

      expect(color?.hex).toBe('#FF0000'); // accent1 from mockTheme
    });

    it('parses prstClr (preset color)', () => {
      const xml = parseXml(`
        <solidFill>
          <prstClr val="blue"/>
        </solidFill>
      `);
      const color = parseColorElement(xml.documentElement, mockTheme);

      expect(color?.hex).toBe('#0000FF');
    });

    it('returns null for empty fill', () => {
      const xml = parseXml('<solidFill/>');
      const color = parseColorElement(xml.documentElement, mockTheme);

      expect(color).toBeNull();
    });
  });
});
