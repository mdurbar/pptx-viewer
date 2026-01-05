/**
 * Parser for PPTX chart XML files.
 *
 * Charts in PPTX are stored in separate XML files (ppt/charts/chartN.xml)
 * and referenced from slides via graphicFrame elements.
 */

import type { ChartType, ChartData, ChartSeries, ChartStyle, Color, ThemeColors } from '../core/types';
import type { PPTXArchive } from '../core/unzip';
import { parseColorElement } from './TextParser';
import {
  parseXml,
  findFirstByName,
  findChildByName,
  findChildrenByName,
  getAttribute,
  getTextContent,
  getBooleanAttribute,
} from '../utils/xml';

/**
 * Parsed chart result.
 */
export interface ParsedChart {
  chartType: ChartType;
  data: ChartData;
  title?: string;
  style?: ChartStyle;
}

/**
 * Parses a chart XML file.
 *
 * @param chartPath - Path to the chart XML file in the archive
 * @param archive - PPTX archive
 * @param themeColors - Theme colors for resolving color references
 * @returns Parsed chart data or null if parsing fails
 */
export function parseChart(
  chartPath: string,
  archive: PPTXArchive,
  themeColors: ThemeColors
): ParsedChart | null {
  const chartXml = archive.getText(chartPath);
  if (!chartXml) return null;

  try {
    const doc = parseXml(chartXml);
    const chartSpace = doc.documentElement;

    // Find the chart element
    const chart = findFirstByName(chartSpace, 'chart');
    if (!chart) return null;

    // Parse title
    const title = parseChartTitle(chart);

    // Find plot area which contains the chart type
    const plotArea = findFirstByName(chart, 'plotArea');
    if (!plotArea) return null;

    // Determine chart type and parse data
    const chartTypeInfo = detectChartType(plotArea);
    if (!chartTypeInfo) return null;

    const { chartType, chartElement } = chartTypeInfo;
    const data = parseChartData(chartElement, chartType, themeColors);

    // Parse style options
    const style = parseChartStyle(chart);

    return {
      chartType,
      data,
      title,
      style,
    };
  } catch {
    return null;
  }
}

/**
 * Parses the chart title.
 */
function parseChartTitle(chart: Element): string | undefined {
  const title = findFirstByName(chart, 'title');
  if (!title) return undefined;

  // Try to find text in tx/rich/p/r/t path
  const tx = findChildByName(title, 'tx');
  if (!tx) return undefined;

  const rich = findChildByName(tx, 'rich');
  if (!rich) return undefined;

  const p = findFirstByName(rich, 'p');
  if (!p) return undefined;

  const r = findFirstByName(p, 'r');
  if (!r) return undefined;

  const t = findChildByName(r, 't');
  if (!t) return undefined;

  return getTextContent(t) || undefined;
}

/**
 * Detects the chart type from the plot area.
 */
function detectChartType(plotArea: Element): { chartType: ChartType; chartElement: Element } | null {
  // Check for different chart types in order of commonality
  const chartTypes: Array<{ tagName: string; type: ChartType }> = [
    { tagName: 'barChart', type: 'column' },
    { tagName: 'bar3DChart', type: 'column' },
    { tagName: 'pieChart', type: 'pie' },
    { tagName: 'pie3DChart', type: 'pie' },
    { tagName: 'doughnutChart', type: 'doughnut' },
    { tagName: 'lineChart', type: 'line' },
    { tagName: 'line3DChart', type: 'line' },
    { tagName: 'areaChart', type: 'area' },
    { tagName: 'area3DChart', type: 'area' },
    { tagName: 'scatterChart', type: 'scatter' },
  ];

  for (const { tagName, type } of chartTypes) {
    const element = findFirstByName(plotArea, tagName);
    if (element) {
      // Check bar direction for bar vs column
      let actualType = type;
      if (tagName === 'barChart' || tagName === 'bar3DChart') {
        const barDir = getAttribute(element, 'barDir') ||
          findChildByName(element, 'barDir')?.getAttribute('val');
        if (barDir === 'bar') {
          actualType = 'bar';
        }
      }

      // Check for stacked bar/column
      if (actualType === 'column' || actualType === 'bar') {
        const grouping = getAttribute(element, 'grouping') ||
          findChildByName(element, 'grouping')?.getAttribute('val');
        if (grouping === 'stacked' || grouping === 'percentStacked') {
          actualType = 'stackedColumn';
        }
      }

      return { chartType: actualType, chartElement: element };
    }
  }

  return null;
}

/**
 * Parses chart data from a chart type element.
 */
function parseChartData(
  chartElement: Element,
  chartType: ChartType,
  themeColors: ThemeColors
): ChartData {
  const categories: string[] = [];
  const series: ChartSeries[] = [];

  // Find all series elements
  const serElements = findChildrenByName(chartElement, 'ser');

  for (const ser of serElements) {
    const seriesData = parseSeries(ser, chartType, themeColors);
    series.push(seriesData);

    // Extract categories from first series (they should be the same across series)
    if (categories.length === 0) {
      const cat = findChildByName(ser, 'cat');
      if (cat) {
        categories.push(...parseCategories(cat));
      }
    }
  }

  // If no categories found, generate numeric labels
  if (categories.length === 0 && series.length > 0) {
    const maxLen = Math.max(...series.map(s => s.values.length));
    for (let i = 0; i < maxLen; i++) {
      categories.push(String(i + 1));
    }
  }

  return { categories, series };
}

/**
 * Parses a single series element.
 */
function parseSeries(
  ser: Element,
  chartType: ChartType,
  themeColors: ThemeColors
): ChartSeries {
  // Parse series name
  const tx = findChildByName(ser, 'tx');
  let name: string | undefined;
  if (tx) {
    // Try strRef/strCache/pt/v path
    const strRef = findChildByName(tx, 'strRef');
    if (strRef) {
      const strCache = findChildByName(strRef, 'strCache');
      if (strCache) {
        const pt = findFirstByName(strCache, 'pt');
        if (pt) {
          const v = findChildByName(pt, 'v');
          if (v) {
            name = getTextContent(v);
          }
        }
      }
    }
    // Try direct v element
    if (!name) {
      const v = findFirstByName(tx, 'v');
      if (v) {
        name = getTextContent(v);
      }
    }
  }

  // Parse values
  const values: number[] = [];
  const val = findChildByName(ser, 'val');
  if (val) {
    values.push(...parseValues(val));
  }

  // For scatter charts, also check yVal
  if (chartType === 'scatter' && values.length === 0) {
    const yVal = findChildByName(ser, 'yVal');
    if (yVal) {
      values.push(...parseValues(yVal));
    }
  }

  // Parse series color
  const color = parseSeriesColor(ser, themeColors);

  return { name, values, color };
}

/**
 * Parses category labels.
 */
function parseCategories(cat: Element): string[] {
  const categories: string[] = [];

  // Try strRef/strCache path
  const strRef = findChildByName(cat, 'strRef');
  if (strRef) {
    const strCache = findChildByName(strRef, 'strCache');
    if (strCache) {
      const pts = findChildrenByName(strCache, 'pt');
      for (const pt of pts) {
        const v = findChildByName(pt, 'v');
        if (v) {
          categories.push(getTextContent(v));
        }
      }
      return categories;
    }
  }

  // Try numRef/numCache path (for numeric categories)
  const numRef = findChildByName(cat, 'numRef');
  if (numRef) {
    const numCache = findChildByName(numRef, 'numCache');
    if (numCache) {
      const pts = findChildrenByName(numCache, 'pt');
      for (const pt of pts) {
        const v = findChildByName(pt, 'v');
        if (v) {
          categories.push(getTextContent(v));
        }
      }
      return categories;
    }
  }

  // Try strLit (literal string data)
  const strLit = findChildByName(cat, 'strLit');
  if (strLit) {
    const pts = findChildrenByName(strLit, 'pt');
    for (const pt of pts) {
      const v = findChildByName(pt, 'v');
      if (v) {
        categories.push(getTextContent(v));
      }
    }
  }

  return categories;
}

/**
 * Parses numeric values.
 */
function parseValues(val: Element): number[] {
  const values: number[] = [];

  // Try numRef/numCache path
  const numRef = findChildByName(val, 'numRef');
  if (numRef) {
    const numCache = findChildByName(numRef, 'numCache');
    if (numCache) {
      const pts = findChildrenByName(numCache, 'pt');
      // Sort by idx attribute to ensure correct order
      const sortedPts = [...pts].sort((a, b) => {
        const idxA = parseInt(getAttribute(a, 'idx') || '0', 10);
        const idxB = parseInt(getAttribute(b, 'idx') || '0', 10);
        return idxA - idxB;
      });
      for (const pt of sortedPts) {
        const v = findChildByName(pt, 'v');
        if (v) {
          const num = parseFloat(getTextContent(v));
          values.push(isNaN(num) ? 0 : num);
        }
      }
      return values;
    }
  }

  // Try numLit (literal numeric data)
  const numLit = findChildByName(val, 'numLit');
  if (numLit) {
    const pts = findChildrenByName(numLit, 'pt');
    const sortedPts = [...pts].sort((a, b) => {
      const idxA = parseInt(getAttribute(a, 'idx') || '0', 10);
      const idxB = parseInt(getAttribute(b, 'idx') || '0', 10);
      return idxA - idxB;
    });
    for (const pt of sortedPts) {
      const v = findChildByName(pt, 'v');
      if (v) {
        const num = parseFloat(getTextContent(v));
        values.push(isNaN(num) ? 0 : num);
      }
    }
  }

  return values;
}

/**
 * Parses series color from shape properties.
 */
function parseSeriesColor(ser: Element, themeColors: ThemeColors): Color | undefined {
  // Look for spPr (shape properties) with solidFill
  const spPr = findChildByName(ser, 'spPr');
  if (spPr) {
    const solidFill = findChildByName(spPr, 'solidFill');
    if (solidFill) {
      return parseColorElement(solidFill, themeColors) || undefined;
    }
  }

  return undefined;
}

/**
 * Parses chart style options.
 */
function parseChartStyle(chart: Element): ChartStyle {
  const style: ChartStyle = {};

  // Check for legend
  const legend = findFirstByName(chart, 'legend');
  if (legend) {
    style.showLegend = true;

    // Parse legend position
    const legendPos = findChildByName(legend, 'legendPos');
    if (legendPos) {
      const pos = getAttribute(legendPos, 'val');
      switch (pos) {
        case 't': style.legendPosition = 'top'; break;
        case 'b': style.legendPosition = 'bottom'; break;
        case 'l': style.legendPosition = 'left'; break;
        case 'r': style.legendPosition = 'right'; break;
        default: style.legendPosition = 'right';
      }
    }
  }

  // Check for data labels
  const dLbls = findFirstByName(chart, 'dLbls');
  if (dLbls) {
    const showVal = findChildByName(dLbls, 'showVal');
    const showCatName = findChildByName(dLbls, 'showCatName');
    const showPercent = findChildByName(dLbls, 'showPercent');

    style.showDataLabels =
      getBooleanAttribute(showVal || dLbls, 'val', false) ||
      getBooleanAttribute(showCatName || dLbls, 'val', false) ||
      getBooleanAttribute(showPercent || dLbls, 'val', false);
  }

  return style;
}

/**
 * Default chart colors to use when no colors are specified.
 */
export const DEFAULT_CHART_COLORS: Color[] = [
  { hex: '#4472C4', alpha: 1 }, // Blue
  { hex: '#ED7D31', alpha: 1 }, // Orange
  { hex: '#A5A5A5', alpha: 1 }, // Gray
  { hex: '#FFC000', alpha: 1 }, // Yellow
  { hex: '#5B9BD5', alpha: 1 }, // Light Blue
  { hex: '#70AD47', alpha: 1 }, // Green
  { hex: '#264478', alpha: 1 }, // Dark Blue
  { hex: '#9E480E', alpha: 1 }, // Dark Orange
  { hex: '#636363', alpha: 1 }, // Dark Gray
  { hex: '#997300', alpha: 1 }, // Dark Yellow
];

/**
 * Gets a color for a series by index.
 */
export function getSeriesColor(index: number, series: ChartSeries, themeColors: ThemeColors): Color {
  // Use series color if defined
  if (series.color) {
    return series.color;
  }

  // Use theme accent colors if available
  const accentColors = [
    themeColors.accent1,
    themeColors.accent2,
    themeColors.accent3,
    themeColors.accent4,
    themeColors.accent5,
    themeColors.accent6,
  ];

  if (index < accentColors.length) {
    return { hex: accentColors[index], alpha: 1 };
  }

  // Fall back to default colors
  return DEFAULT_CHART_COLORS[index % DEFAULT_CHART_COLORS.length];
}
