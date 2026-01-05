import { describe, it, expect } from 'vitest';
import { parseChart, getSeriesColor, DEFAULT_CHART_COLORS } from '../src/parser/ChartParser';
import type { ThemeColors, ChartSeries } from '../src/core/types';
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

// Helper to create a mock archive with chart XML
function createMockArchive(chartXml: string): PPTXArchive {
  return {
    files: new Map(),
    getText: (path: string) => {
      if (path === 'ppt/charts/chart1.xml') {
        return chartXml;
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

describe('ChartParser', () => {
  describe('parseChart', () => {
    it('returns null for missing chart file', () => {
      const archive = createMockArchive('');
      archive.getText = () => null;

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);
      expect(result).toBeNull();
    });

    it('returns null for invalid XML', () => {
      const archive = createMockArchive('<invalid');

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);
      expect(result).toBeNull();
    });

    it('parses a basic column chart', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:barDir val="col"/>
                <c:ser>
                  <c:tx>
                    <c:strRef>
                      <c:strCache>
                        <c:pt idx="0"><c:v>Sales</c:v></c:pt>
                      </c:strCache>
                    </c:strRef>
                  </c:tx>
                  <c:cat>
                    <c:strRef>
                      <c:strCache>
                        <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                        <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                        <c:pt idx="2"><c:v>Q3</c:v></c:pt>
                      </c:strCache>
                    </c:strRef>
                  </c:cat>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>100</c:v></c:pt>
                        <c:pt idx="1"><c:v>200</c:v></c:pt>
                        <c:pt idx="2"><c:v>150</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.chartType).toBe('column');
      expect(result?.data.categories).toEqual(['Q1', 'Q2', 'Q3']);
      expect(result?.data.series).toHaveLength(1);
      expect(result?.data.series[0].name).toBe('Sales');
      expect(result?.data.series[0].values).toEqual([100, 200, 150]);
    });

    it('parses a horizontal bar chart', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:barDir val="bar"/>
                <c:ser>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>50</c:v></c:pt>
                        <c:pt idx="1"><c:v>75</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.chartType).toBe('bar');
    });

    it('parses a stacked column chart', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:barDir val="col"/>
                <c:grouping val="stacked"/>
                <c:ser>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>10</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.chartType).toBe('stackedColumn');
    });

    it('parses a pie chart', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:pieChart>
                <c:ser>
                  <c:cat>
                    <c:strRef>
                      <c:strCache>
                        <c:pt idx="0"><c:v>Apples</c:v></c:pt>
                        <c:pt idx="1"><c:v>Oranges</c:v></c:pt>
                        <c:pt idx="2"><c:v>Bananas</c:v></c:pt>
                      </c:strCache>
                    </c:strRef>
                  </c:cat>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>30</c:v></c:pt>
                        <c:pt idx="1"><c:v>45</c:v></c:pt>
                        <c:pt idx="2"><c:v>25</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:pieChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.chartType).toBe('pie');
      expect(result?.data.categories).toEqual(['Apples', 'Oranges', 'Bananas']);
      expect(result?.data.series[0].values).toEqual([30, 45, 25]);
    });

    it('parses a doughnut chart', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:doughnutChart>
                <c:ser>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>60</c:v></c:pt>
                        <c:pt idx="1"><c:v>40</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:doughnutChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.chartType).toBe('doughnut');
    });

    it('parses a line chart', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:lineChart>
                <c:ser>
                  <c:tx>
                    <c:strRef>
                      <c:strCache>
                        <c:pt idx="0"><c:v>Revenue</c:v></c:pt>
                      </c:strCache>
                    </c:strRef>
                  </c:tx>
                  <c:cat>
                    <c:strRef>
                      <c:strCache>
                        <c:pt idx="0"><c:v>Jan</c:v></c:pt>
                        <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                        <c:pt idx="2"><c:v>Mar</c:v></c:pt>
                        <c:pt idx="3"><c:v>Apr</c:v></c:pt>
                      </c:strCache>
                    </c:strRef>
                  </c:cat>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>1000</c:v></c:pt>
                        <c:pt idx="1"><c:v>1200</c:v></c:pt>
                        <c:pt idx="2"><c:v>1100</c:v></c:pt>
                        <c:pt idx="3"><c:v>1500</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:lineChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.chartType).toBe('line');
      expect(result?.data.categories).toEqual(['Jan', 'Feb', 'Mar', 'Apr']);
      expect(result?.data.series[0].name).toBe('Revenue');
      expect(result?.data.series[0].values).toEqual([1000, 1200, 1100, 1500]);
    });

    it('parses an area chart', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:areaChart>
                <c:ser>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>5</c:v></c:pt>
                        <c:pt idx="1"><c:v>10</c:v></c:pt>
                        <c:pt idx="2"><c:v>8</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:areaChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.chartType).toBe('area');
    });

    it('parses a scatter chart', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:scatterChart>
                <c:ser>
                  <c:yVal>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>2</c:v></c:pt>
                        <c:pt idx="1"><c:v>4</c:v></c:pt>
                        <c:pt idx="2"><c:v>3</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:yVal>
                </c:ser>
              </c:scatterChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.chartType).toBe('scatter');
      expect(result?.data.series[0].values).toEqual([2, 4, 3]);
    });

    it('parses chart title', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <c:chart>
            <c:title>
              <c:tx>
                <c:rich>
                  <a:p>
                    <a:r>
                      <a:t>Quarterly Sales</a:t>
                    </a:r>
                  </a:p>
                </c:rich>
              </c:tx>
            </c:title>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>100</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.title).toBe('Quarterly Sales');
    });

    it('parses multiple series', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:tx><c:v>2023</c:v></c:tx>
                  <c:cat>
                    <c:strRef>
                      <c:strCache>
                        <c:pt idx="0"><c:v>Q1</c:v></c:pt>
                        <c:pt idx="1"><c:v>Q2</c:v></c:pt>
                      </c:strCache>
                    </c:strRef>
                  </c:cat>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>100</c:v></c:pt>
                        <c:pt idx="1"><c:v>150</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
                <c:ser>
                  <c:tx><c:v>2024</c:v></c:tx>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>120</c:v></c:pt>
                        <c:pt idx="1"><c:v>180</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.data.series).toHaveLength(2);
      expect(result?.data.series[0].name).toBe('2023');
      expect(result?.data.series[0].values).toEqual([100, 150]);
      expect(result?.data.series[1].name).toBe('2024');
      expect(result?.data.series[1].values).toEqual([120, 180]);
    });

    it('parses legend settings', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:legend>
              <c:legendPos val="b"/>
            </c:legend>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>100</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.style?.showLegend).toBe(true);
      expect(result?.style?.legendPosition).toBe('bottom');
    });

    it('generates numeric categories when none provided', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="0"><c:v>10</c:v></c:pt>
                        <c:pt idx="1"><c:v>20</c:v></c:pt>
                        <c:pt idx="2"><c:v>30</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      expect(result?.data.categories).toEqual(['1', '2', '3']);
    });

    it('handles values out of order by idx', () => {
      const chartXml = `
        <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart>
            <c:plotArea>
              <c:barChart>
                <c:ser>
                  <c:val>
                    <c:numRef>
                      <c:numCache>
                        <c:pt idx="2"><c:v>30</c:v></c:pt>
                        <c:pt idx="0"><c:v>10</c:v></c:pt>
                        <c:pt idx="1"><c:v>20</c:v></c:pt>
                      </c:numCache>
                    </c:numRef>
                  </c:val>
                </c:ser>
              </c:barChart>
            </c:plotArea>
          </c:chart>
        </c:chartSpace>
      `;
      const archive = createMockArchive(chartXml);

      const result = parseChart('ppt/charts/chart1.xml', archive, mockTheme);

      expect(result).not.toBeNull();
      // Values should be sorted by idx
      expect(result?.data.series[0].values).toEqual([10, 20, 30]);
    });
  });

  describe('getSeriesColor', () => {
    it('uses series color when defined', () => {
      const series: ChartSeries = {
        name: 'Test',
        values: [1, 2, 3],
        color: { hex: '#123456', alpha: 1 },
      };

      const color = getSeriesColor(0, series, mockTheme);

      expect(color.hex).toBe('#123456');
    });

    it('uses theme accent colors by index', () => {
      const series: ChartSeries = {
        name: 'Test',
        values: [1, 2, 3],
      };

      const color0 = getSeriesColor(0, series, mockTheme);
      const color1 = getSeriesColor(1, series, mockTheme);
      const color2 = getSeriesColor(2, series, mockTheme);

      expect(color0.hex).toBe('#FF0000'); // accent1
      expect(color1.hex).toBe('#00FF00'); // accent2
      expect(color2.hex).toBe('#0000FF'); // accent3
    });

    it('falls back to default colors when beyond accent range', () => {
      const series: ChartSeries = {
        name: 'Test',
        values: [1, 2, 3],
      };

      const color = getSeriesColor(10, series, mockTheme);

      // Should use DEFAULT_CHART_COLORS[10 % 10] = DEFAULT_CHART_COLORS[0]
      expect(color.hex).toBe(DEFAULT_CHART_COLORS[0].hex);
    });
  });
});
