/**
 * Renderer for chart elements.
 *
 * Renders parsed chart data as SVG graphics.
 */

import type { ChartElement, ChartData, ChartSeries, Color, ThemeColors } from '../core/types';
import { getSeriesColor, DEFAULT_CHART_COLORS } from '../parser/ChartParser';
import { colorToCss } from '../utils/color';

/**
 * Chart rendering options.
 */
export interface ChartRenderOptions {
  /** Theme colors for series coloring */
  themeColors: ThemeColors;
  /** Whether to show the legend */
  showLegend?: boolean;
  /** Whether to show data labels */
  showDataLabels?: boolean;
  /** Whether to show gridlines */
  showGridlines?: boolean;
}

/**
 * Default chart render options.
 */
const DEFAULT_OPTIONS: ChartRenderOptions = {
  themeColors: {
    dark1: '#000000',
    dark2: '#444444',
    light1: '#FFFFFF',
    light2: '#EEEEEE',
    accent1: '#4472C4',
    accent2: '#ED7D31',
    accent3: '#A5A5A5',
    accent4: '#FFC000',
    accent5: '#5B9BD5',
    accent6: '#70AD47',
    hlink: '#0563C1',
    folHlink: '#954F72',
  },
  showLegend: true,
  showDataLabels: false,
  showGridlines: true,
};

/**
 * Renders a chart element to SVG.
 */
export function renderChart(chart: ChartElement, options: Partial<ChartRenderOptions> = {}): SVGGElement {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const group = document.createElementNS('http://www.w3.org/2000/svg', 'g');

  // If chart type is unknown or has fallback image, render the fallback
  if (chart.chartType === 'unknown' && chart.fallbackImage) {
    return renderFallbackImage(chart, group);
  }

  // Calculate chart area dimensions (leaving room for title and legend)
  const padding = 20;
  const titleHeight = chart.title ? 30 : 0;
  const legendHeight = (opts.showLegend || chart.style?.showLegend) ? 30 : 0;

  const chartArea = {
    x: padding,
    y: padding + titleHeight,
    width: chart.bounds.width - padding * 2,
    height: chart.bounds.height - padding * 2 - titleHeight - legendHeight,
  };

  // Render title if present
  if (chart.title) {
    renderChartTitle(group, chart.title, chart.bounds.width, padding);
  }

  // Render chart based on type
  switch (chart.chartType) {
    case 'bar':
      renderBarChart(group, chart.data, chartArea, opts, true);
      break;
    case 'column':
    case 'stackedColumn':
      renderBarChart(group, chart.data, chartArea, opts, false, chart.chartType === 'stackedColumn');
      break;
    case 'pie':
    case 'doughnut':
      renderPieChart(group, chart.data, chartArea, opts, chart.chartType === 'doughnut');
      break;
    case 'line':
      renderLineChart(group, chart.data, chartArea, opts);
      break;
    case 'area':
      renderAreaChart(group, chart.data, chartArea, opts);
      break;
    case 'scatter':
      renderScatterChart(group, chart.data, chartArea, opts);
      break;
    default:
      // Unknown chart type - render placeholder
      renderPlaceholder(group, chart.bounds, 'Chart');
  }

  // Render legend if enabled
  if (opts.showLegend || chart.style?.showLegend) {
    const legendY = chart.bounds.height - legendHeight;
    renderLegend(group, chart.data.series, chart.bounds.width, legendY, opts);
  }

  return group;
}

/**
 * Renders a fallback image.
 */
function renderFallbackImage(chart: ChartElement, group: SVGGElement): SVGGElement {
  const image = document.createElementNS('http://www.w3.org/2000/svg', 'image');
  image.setAttribute('x', '0');
  image.setAttribute('y', '0');
  image.setAttribute('width', String(chart.bounds.width));
  image.setAttribute('height', String(chart.bounds.height));
  image.setAttribute('href', chart.fallbackImage!);
  image.setAttribute('preserveAspectRatio', 'xMidYMid meet');
  group.appendChild(image);
  return group;
}

/**
 * Renders a chart title.
 */
function renderChartTitle(group: SVGGElement, title: string, width: number, y: number): void {
  const text = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  text.setAttribute('x', String(width / 2));
  text.setAttribute('y', String(y));
  text.setAttribute('text-anchor', 'middle');
  text.setAttribute('font-family', 'Arial, sans-serif');
  text.setAttribute('font-size', '14');
  text.setAttribute('font-weight', 'bold');
  text.setAttribute('fill', '#333333');
  text.textContent = title;
  group.appendChild(text);
}

/**
 * Renders a bar/column chart.
 */
function renderBarChart(
  group: SVGGElement,
  data: ChartData,
  area: { x: number; y: number; width: number; height: number },
  options: ChartRenderOptions,
  horizontal: boolean,
  stacked: boolean = false
): void {
  const { categories, series } = data;
  if (categories.length === 0 || series.length === 0) return;

  // Calculate max value for scaling
  let maxValue: number;
  if (stacked) {
    maxValue = Math.max(...categories.map((_, i) =>
      series.reduce((sum, s) => sum + (s.values[i] || 0), 0)
    ));
  } else {
    maxValue = Math.max(...series.flatMap(s => s.values));
  }
  if (maxValue <= 0) maxValue = 1;

  // Render gridlines
  if (options.showGridlines) {
    renderGridlines(group, area, maxValue, horizontal);
  }

  const numCategories = categories.length;
  const numSeries = series.length;
  const gap = 0.2; // Gap between category groups
  const innerGap = 0.05; // Gap between bars in a group

  if (horizontal) {
    // Horizontal bars
    const barHeight = (area.height / numCategories) * (1 - gap);
    const categorySpacing = area.height / numCategories;

    if (stacked) {
      // Stacked bars
      for (let i = 0; i < numCategories; i++) {
        let xOffset = 0;
        for (let j = 0; j < numSeries; j++) {
          const value = series[j].values[i] || 0;
          const barWidth = (value / maxValue) * area.width;
          const color = getSeriesColor(j, series[j], options.themeColors);

          const rect = createRect(
            area.x + xOffset,
            area.y + i * categorySpacing + (categorySpacing - barHeight) / 2,
            barWidth,
            barHeight,
            color
          );
          group.appendChild(rect);
          xOffset += barWidth;
        }
      }
    } else {
      // Grouped bars
      const singleBarHeight = barHeight / numSeries;
      for (let i = 0; i < numCategories; i++) {
        for (let j = 0; j < numSeries; j++) {
          const value = series[j].values[i] || 0;
          const barWidth = (value / maxValue) * area.width;
          const color = getSeriesColor(j, series[j], options.themeColors);

          const rect = createRect(
            area.x,
            area.y + i * categorySpacing + (categorySpacing - barHeight) / 2 + j * singleBarHeight,
            barWidth,
            singleBarHeight * (1 - innerGap),
            color
          );
          group.appendChild(rect);
        }
      }
    }

    // Render category labels
    for (let i = 0; i < numCategories; i++) {
      const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      label.setAttribute('x', String(area.x - 5));
      label.setAttribute('y', String(area.y + i * categorySpacing + categorySpacing / 2));
      label.setAttribute('text-anchor', 'end');
      label.setAttribute('dominant-baseline', 'middle');
      label.setAttribute('font-family', 'Arial, sans-serif');
      label.setAttribute('font-size', '10');
      label.setAttribute('fill', '#666666');
      label.textContent = truncateLabel(categories[i], 15);
      group.appendChild(label);
    }
  } else {
    // Vertical columns
    const barWidth = (area.width / numCategories) * (1 - gap);
    const categorySpacing = area.width / numCategories;

    if (stacked) {
      // Stacked columns
      for (let i = 0; i < numCategories; i++) {
        let yOffset = 0;
        for (let j = 0; j < numSeries; j++) {
          const value = series[j].values[i] || 0;
          const barHeight = (value / maxValue) * area.height;
          const color = getSeriesColor(j, series[j], options.themeColors);

          const rect = createRect(
            area.x + i * categorySpacing + (categorySpacing - barWidth) / 2,
            area.y + area.height - yOffset - barHeight,
            barWidth,
            barHeight,
            color
          );
          group.appendChild(rect);
          yOffset += barHeight;
        }
      }
    } else {
      // Grouped columns
      const singleBarWidth = barWidth / numSeries;
      for (let i = 0; i < numCategories; i++) {
        for (let j = 0; j < numSeries; j++) {
          const value = series[j].values[i] || 0;
          const barHeight = (value / maxValue) * area.height;
          const color = getSeriesColor(j, series[j], options.themeColors);

          const rect = createRect(
            area.x + i * categorySpacing + (categorySpacing - barWidth) / 2 + j * singleBarWidth,
            area.y + area.height - barHeight,
            singleBarWidth * (1 - innerGap),
            barHeight,
            color
          );
          group.appendChild(rect);
        }
      }
    }

    // Render category labels
    for (let i = 0; i < numCategories; i++) {
      const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      label.setAttribute('x', String(area.x + i * categorySpacing + categorySpacing / 2));
      label.setAttribute('y', String(area.y + area.height + 15));
      label.setAttribute('text-anchor', 'middle');
      label.setAttribute('font-family', 'Arial, sans-serif');
      label.setAttribute('font-size', '10');
      label.setAttribute('fill', '#666666');
      label.textContent = truncateLabel(categories[i], 10);
      group.appendChild(label);
    }
  }
}

/**
 * Renders a pie or doughnut chart.
 */
function renderPieChart(
  group: SVGGElement,
  data: ChartData,
  area: { x: number; y: number; width: number; height: number },
  options: ChartRenderOptions,
  isDoughnut: boolean = false
): void {
  // For pie charts, we typically use the first series
  const series = data.series[0];
  if (!series || series.values.length === 0) return;

  const values = series.values;
  const total = values.reduce((sum, v) => sum + Math.abs(v), 0);
  if (total <= 0) return;

  const centerX = area.x + area.width / 2;
  const centerY = area.y + area.height / 2;
  const radius = Math.min(area.width, area.height) / 2 - 10;
  const innerRadius = isDoughnut ? radius * 0.5 : 0;

  let startAngle = -Math.PI / 2; // Start at top

  for (let i = 0; i < values.length; i++) {
    const value = Math.abs(values[i]);
    const sweepAngle = (value / total) * 2 * Math.PI;
    const endAngle = startAngle + sweepAngle;

    const color = getSeriesColor(i, { values: [value] }, options.themeColors);
    const path = createPieSlice(centerX, centerY, radius, innerRadius, startAngle, endAngle, color);
    group.appendChild(path);

    // Add label if there's room
    if (sweepAngle > 0.15 && data.categories[i]) {
      const labelAngle = startAngle + sweepAngle / 2;
      const labelRadius = (radius + innerRadius) / 2;
      const labelX = centerX + Math.cos(labelAngle) * labelRadius;
      const labelY = centerY + Math.sin(labelAngle) * labelRadius;

      const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
      label.setAttribute('x', String(labelX));
      label.setAttribute('y', String(labelY));
      label.setAttribute('text-anchor', 'middle');
      label.setAttribute('dominant-baseline', 'middle');
      label.setAttribute('font-family', 'Arial, sans-serif');
      label.setAttribute('font-size', '10');
      label.setAttribute('fill', '#FFFFFF');
      label.textContent = truncateLabel(data.categories[i], 8);
      group.appendChild(label);
    }

    startAngle = endAngle;
  }
}

/**
 * Creates a pie slice path.
 */
function createPieSlice(
  cx: number,
  cy: number,
  outerRadius: number,
  innerRadius: number,
  startAngle: number,
  endAngle: number,
  color: Color
): SVGPathElement {
  const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');

  const x1 = cx + Math.cos(startAngle) * outerRadius;
  const y1 = cy + Math.sin(startAngle) * outerRadius;
  const x2 = cx + Math.cos(endAngle) * outerRadius;
  const y2 = cy + Math.sin(endAngle) * outerRadius;

  const largeArcFlag = endAngle - startAngle > Math.PI ? 1 : 0;

  let d: string;
  if (innerRadius > 0) {
    // Doughnut slice
    const x3 = cx + Math.cos(endAngle) * innerRadius;
    const y3 = cy + Math.sin(endAngle) * innerRadius;
    const x4 = cx + Math.cos(startAngle) * innerRadius;
    const y4 = cy + Math.sin(startAngle) * innerRadius;

    d = `M ${x1} ${y1} A ${outerRadius} ${outerRadius} 0 ${largeArcFlag} 1 ${x2} ${y2}
         L ${x3} ${y3} A ${innerRadius} ${innerRadius} 0 ${largeArcFlag} 0 ${x4} ${y4} Z`;
  } else {
    // Full pie slice
    d = `M ${cx} ${cy} L ${x1} ${y1} A ${outerRadius} ${outerRadius} 0 ${largeArcFlag} 1 ${x2} ${y2} Z`;
  }

  path.setAttribute('d', d);
  path.setAttribute('fill', colorToCss(color));
  path.setAttribute('stroke', '#FFFFFF');
  path.setAttribute('stroke-width', '1');

  return path;
}

/**
 * Renders a line chart.
 */
function renderLineChart(
  group: SVGGElement,
  data: ChartData,
  area: { x: number; y: number; width: number; height: number },
  options: ChartRenderOptions
): void {
  const { categories, series } = data;
  if (categories.length === 0 || series.length === 0) return;

  const maxValue = Math.max(...series.flatMap(s => s.values));
  const minValue = Math.min(0, Math.min(...series.flatMap(s => s.values)));
  const range = maxValue - minValue || 1;

  // Render gridlines
  if (options.showGridlines) {
    renderGridlines(group, area, maxValue, false);
  }

  const numPoints = Math.max(...series.map(s => s.values.length));
  const xStep = area.width / (numPoints - 1 || 1);

  // Render each series
  for (let j = 0; j < series.length; j++) {
    const s = series[j];
    const color = getSeriesColor(j, s, options.themeColors);
    const points: string[] = [];

    for (let i = 0; i < s.values.length; i++) {
      const x = area.x + i * xStep;
      const y = area.y + area.height - ((s.values[i] - minValue) / range) * area.height;
      points.push(`${x},${y}`);
    }

    // Draw line
    const polyline = document.createElementNS('http://www.w3.org/2000/svg', 'polyline');
    polyline.setAttribute('points', points.join(' '));
    polyline.setAttribute('fill', 'none');
    polyline.setAttribute('stroke', colorToCss(color));
    polyline.setAttribute('stroke-width', '2');
    polyline.setAttribute('stroke-linecap', 'round');
    polyline.setAttribute('stroke-linejoin', 'round');
    group.appendChild(polyline);

    // Draw data points
    for (let i = 0; i < s.values.length; i++) {
      const x = area.x + i * xStep;
      const y = area.y + area.height - ((s.values[i] - minValue) / range) * area.height;

      const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circle.setAttribute('cx', String(x));
      circle.setAttribute('cy', String(y));
      circle.setAttribute('r', '4');
      circle.setAttribute('fill', colorToCss(color));
      circle.setAttribute('stroke', '#FFFFFF');
      circle.setAttribute('stroke-width', '1');
      group.appendChild(circle);
    }
  }

  // Render category labels
  for (let i = 0; i < categories.length; i++) {
    const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    label.setAttribute('x', String(area.x + i * xStep));
    label.setAttribute('y', String(area.y + area.height + 15));
    label.setAttribute('text-anchor', 'middle');
    label.setAttribute('font-family', 'Arial, sans-serif');
    label.setAttribute('font-size', '10');
    label.setAttribute('fill', '#666666');
    label.textContent = truncateLabel(categories[i], 10);
    group.appendChild(label);
  }
}

/**
 * Renders an area chart.
 */
function renderAreaChart(
  group: SVGGElement,
  data: ChartData,
  area: { x: number; y: number; width: number; height: number },
  options: ChartRenderOptions
): void {
  const { categories, series } = data;
  if (categories.length === 0 || series.length === 0) return;

  const maxValue = Math.max(...series.flatMap(s => s.values));
  const minValue = Math.min(0, Math.min(...series.flatMap(s => s.values)));
  const range = maxValue - minValue || 1;

  // Render gridlines
  if (options.showGridlines) {
    renderGridlines(group, area, maxValue, false);
  }

  const numPoints = Math.max(...series.map(s => s.values.length));
  const xStep = area.width / (numPoints - 1 || 1);
  const baseline = area.y + area.height;

  // Render each series as filled area
  for (let j = 0; j < series.length; j++) {
    const s = series[j];
    const color = getSeriesColor(j, s, options.themeColors);

    let pathD = `M ${area.x} ${baseline}`;
    for (let i = 0; i < s.values.length; i++) {
      const x = area.x + i * xStep;
      const y = area.y + area.height - ((s.values[i] - minValue) / range) * area.height;
      pathD += ` L ${x} ${y}`;
    }
    pathD += ` L ${area.x + (s.values.length - 1) * xStep} ${baseline} Z`;

    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', pathD);
    path.setAttribute('fill', colorToCss({ ...color, alpha: 0.5 }));
    path.setAttribute('stroke', colorToCss(color));
    path.setAttribute('stroke-width', '2');
    group.appendChild(path);
  }

  // Render category labels
  for (let i = 0; i < categories.length; i++) {
    const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    label.setAttribute('x', String(area.x + i * xStep));
    label.setAttribute('y', String(area.y + area.height + 15));
    label.setAttribute('text-anchor', 'middle');
    label.setAttribute('font-family', 'Arial, sans-serif');
    label.setAttribute('font-size', '10');
    label.setAttribute('fill', '#666666');
    label.textContent = truncateLabel(categories[i], 10);
    group.appendChild(label);
  }
}

/**
 * Renders a scatter chart.
 */
function renderScatterChart(
  group: SVGGElement,
  data: ChartData,
  area: { x: number; y: number; width: number; height: number },
  options: ChartRenderOptions
): void {
  const { series } = data;
  if (series.length === 0) return;

  // For scatter charts, we need X and Y values
  // Typically first series has X values, second has Y values
  // Or each series has its own X/Y values
  const allValues = series.flatMap(s => s.values);
  const maxValue = Math.max(...allValues);
  const minValue = Math.min(...allValues);
  const range = maxValue - minValue || 1;

  // Render gridlines
  if (options.showGridlines) {
    renderGridlines(group, area, maxValue, false);
  }

  // Render each series
  for (let j = 0; j < series.length; j++) {
    const s = series[j];
    const color = getSeriesColor(j, s, options.themeColors);

    // Plot points assuming values are Y coordinates, index is X
    for (let i = 0; i < s.values.length; i++) {
      const x = area.x + (i / (s.values.length - 1 || 1)) * area.width;
      const y = area.y + area.height - ((s.values[i] - minValue) / range) * area.height;

      const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
      circle.setAttribute('cx', String(x));
      circle.setAttribute('cy', String(y));
      circle.setAttribute('r', '5');
      circle.setAttribute('fill', colorToCss(color));
      circle.setAttribute('stroke', '#FFFFFF');
      circle.setAttribute('stroke-width', '1');
      group.appendChild(circle);
    }
  }
}

/**
 * Renders gridlines for a chart.
 */
function renderGridlines(
  group: SVGGElement,
  area: { x: number; y: number; width: number; height: number },
  maxValue: number,
  horizontal: boolean
): void {
  const numGridlines = 5;
  const gridlineColor = '#E0E0E0';

  for (let i = 0; i <= numGridlines; i++) {
    const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');

    if (horizontal) {
      const x = area.x + (i / numGridlines) * area.width;
      line.setAttribute('x1', String(x));
      line.setAttribute('y1', String(area.y));
      line.setAttribute('x2', String(x));
      line.setAttribute('y2', String(area.y + area.height));
    } else {
      const y = area.y + (i / numGridlines) * area.height;
      line.setAttribute('x1', String(area.x));
      line.setAttribute('y1', String(y));
      line.setAttribute('x2', String(area.x + area.width));
      line.setAttribute('y2', String(y));
    }

    line.setAttribute('stroke', gridlineColor);
    line.setAttribute('stroke-width', '1');
    line.setAttribute('stroke-dasharray', '2,2');
    group.appendChild(line);

    // Add value labels
    const value = horizontal
      ? (i / numGridlines) * maxValue
      : maxValue - (i / numGridlines) * maxValue;

    const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    if (horizontal) {
      label.setAttribute('x', String(area.x + (i / numGridlines) * area.width));
      label.setAttribute('y', String(area.y + area.height + 12));
      label.setAttribute('text-anchor', 'middle');
    } else {
      label.setAttribute('x', String(area.x - 5));
      label.setAttribute('y', String(area.y + (i / numGridlines) * area.height));
      label.setAttribute('text-anchor', 'end');
      label.setAttribute('dominant-baseline', 'middle');
    }
    label.setAttribute('font-family', 'Arial, sans-serif');
    label.setAttribute('font-size', '9');
    label.setAttribute('fill', '#999999');
    label.textContent = formatValue(value);
    group.appendChild(label);
  }
}

/**
 * Renders a legend.
 */
function renderLegend(
  group: SVGGElement,
  series: ChartSeries[],
  width: number,
  y: number,
  options: ChartRenderOptions
): void {
  const itemWidth = 100;
  const totalWidth = series.length * itemWidth;
  const startX = (width - totalWidth) / 2;

  for (let i = 0; i < series.length; i++) {
    const x = startX + i * itemWidth;
    const color = getSeriesColor(i, series[i], options.themeColors);

    // Color swatch
    const rect = createRect(x, y + 5, 12, 12, color);
    group.appendChild(rect);

    // Label
    const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    label.setAttribute('x', String(x + 16));
    label.setAttribute('y', String(y + 14));
    label.setAttribute('font-family', 'Arial, sans-serif');
    label.setAttribute('font-size', '10');
    label.setAttribute('fill', '#666666');
    label.textContent = truncateLabel(series[i].name || `Series ${i + 1}`, 12);
    group.appendChild(label);
  }
}

/**
 * Renders a placeholder for unknown chart types.
 */
function renderPlaceholder(
  group: SVGGElement,
  bounds: { width: number; height: number },
  text: string
): void {
  // Background
  const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  rect.setAttribute('x', '0');
  rect.setAttribute('y', '0');
  rect.setAttribute('width', String(bounds.width));
  rect.setAttribute('height', String(bounds.height));
  rect.setAttribute('fill', '#F0F0F0');
  rect.setAttribute('stroke', '#CCCCCC');
  rect.setAttribute('stroke-width', '1');
  group.appendChild(rect);

  // Text
  const label = document.createElementNS('http://www.w3.org/2000/svg', 'text');
  label.setAttribute('x', String(bounds.width / 2));
  label.setAttribute('y', String(bounds.height / 2));
  label.setAttribute('text-anchor', 'middle');
  label.setAttribute('dominant-baseline', 'middle');
  label.setAttribute('font-family', 'Arial, sans-serif');
  label.setAttribute('font-size', '14');
  label.setAttribute('fill', '#999999');
  label.textContent = text;
  group.appendChild(label);
}

/**
 * Creates a rectangle element.
 */
function createRect(x: number, y: number, width: number, height: number, color: Color): SVGRectElement {
  const rect = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
  rect.setAttribute('x', String(x));
  rect.setAttribute('y', String(y));
  rect.setAttribute('width', String(Math.max(0, width)));
  rect.setAttribute('height', String(Math.max(0, height)));
  rect.setAttribute('fill', colorToCss(color));
  return rect;
}

/**
 * Truncates a label to a maximum length.
 */
function truncateLabel(text: string, maxLength: number): string {
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength - 1) + '…';
}

/**
 * Formats a numeric value for display.
 */
function formatValue(value: number): string {
  if (Math.abs(value) >= 1000000) {
    return (value / 1000000).toFixed(1) + 'M';
  }
  if (Math.abs(value) >= 1000) {
    return (value / 1000).toFixed(1) + 'K';
  }
  if (Number.isInteger(value)) {
    return String(value);
  }
  return value.toFixed(1);
}
