/**
 * Renderer for tables.
 *
 * Tables are rendered as HTML tables inside SVG foreignObject elements
 * for proper text rendering and cell layout.
 */

import type { TableElement, TableRow, TableCell, CellBorders, Stroke } from '../core/types';
import { colorToCss } from '../utils/color';
import { SVG_NS } from '../utils/svg';
import { sanitizeFontFamily } from '../utils/css';
import { SINGLE_LINE_SPACING, MIN_LINE_HEIGHT, MIN_LINE_HEIGHT_PX } from './TextRenderer';

/**
 * Renders a table element to an SVG foreignObject containing an HTML table.
 *
 * @param table - The table element to render
 * @returns SVG foreignObject element containing the HTML table
 */
export function renderTable(table: TableElement): SVGForeignObjectElement {
  const fo = document.createElementNS(SVG_NS, 'foreignObject');
  fo.setAttribute('width', String(table.bounds.width));
  fo.setAttribute('height', String(table.bounds.height));

  // Create HTML table
  const htmlTable = document.createElement('table');
  htmlTable.style.cssText = `
    width: 100%;
    height: 100%;
    border-collapse: collapse;
    table-layout: fixed;
    font-family: Calibri, Arial, sans-serif;
    font-size: 14px;
  `;

  // Create colgroup for column widths
  const colgroup = document.createElement('colgroup');
  const totalWidth = table.columnWidths.reduce((sum, w) => sum + w, 0);

  for (const width of table.columnWidths) {
    const col = document.createElement('col');
    // Use percentage widths based on the proportions
    const percentage = totalWidth > 0 ? (width / totalWidth) * 100 : 100 / table.columnWidths.length;
    col.style.width = `${percentage}%`;
    colgroup.appendChild(col);
  }
  htmlTable.appendChild(colgroup);

  // Create tbody
  const tbody = document.createElement('tbody');

  for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex++) {
    const row = table.rows[rowIndex];
    const tr = document.createElement('tr');

    // Set row height if specified
    if (row.height > 0) {
      tr.style.height = `${row.height}px`;
    }

    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex++) {
      const cell = row.cells[cellIndex];
      const td = renderTableCell(cell, table, rowIndex, cellIndex);
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  htmlTable.appendChild(tbody);

  // Wrap in a div for proper sizing
  const wrapper = document.createElement('div');
  wrapper.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
  wrapper.style.cssText = `
    width: 100%;
    height: 100%;
    overflow: hidden;
  `;
  wrapper.appendChild(htmlTable);

  fo.appendChild(wrapper);

  return fo;
}

/**
 * Renders a table cell.
 */
function renderTableCell(
  cell: TableCell,
  table: TableElement,
  rowIndex: number,
  cellIndex: number
): HTMLTableCellElement {
  const td = document.createElement('td');

  // Apply colspan/rowspan
  if (cell.colSpan && cell.colSpan > 1) {
    td.colSpan = cell.colSpan;
  }
  if (cell.rowSpan && cell.rowSpan > 1) {
    td.rowSpan = cell.rowSpan;
  }

  // Build cell styles
  const styles: string[] = [];

  // Vertical alignment
  switch (cell.verticalAlign) {
    case 'middle':
      styles.push('vertical-align: middle');
      break;
    case 'bottom':
      styles.push('vertical-align: bottom');
      break;
    default:
      styles.push('vertical-align: top');
  }

  // Cell fill/background
  if (cell.fill && cell.fill.type === 'solid') {
    const bgColor = colorToCss(cell.fill.color);
    styles.push(`background-color: ${bgColor}`);
  }

  // Cell borders
  if (cell.borders) {
    if (cell.borders.top) {
      styles.push(`border-top: ${formatBorder(cell.borders.top)}`);
    }
    if (cell.borders.right) {
      styles.push(`border-right: ${formatBorder(cell.borders.right)}`);
    }
    if (cell.borders.bottom) {
      styles.push(`border-bottom: ${formatBorder(cell.borders.bottom)}`);
    }
    if (cell.borders.left) {
      styles.push(`border-left: ${formatBorder(cell.borders.left)}`);
    }
  } else {
    // Default subtle border
    styles.push('border: 1px solid #d0d0d0');
  }

  // Padding
  styles.push('padding: 4px 8px');

  // Word wrap
  styles.push('word-wrap: break-word');
  styles.push('overflow: hidden');

  td.style.cssText = styles.join('; ');

  // Render text content
  if (cell.text && cell.text.paragraphs.length > 0) {
    for (const p of renderCellText(cell)) {
      td.appendChild(p);
    }
  }

  return td;
}

/**
 * Renders cell text content as DOM elements.
 *
 * Built with createElement/textContent/style assignment (never innerHTML):
 * run text, font names, and colors come from untrusted PPTX content.
 */
function renderCellText(cell: TableCell): HTMLParagraphElement[] {
  if (!cell.text) return [];

  const paragraphs: HTMLParagraphElement[] = [];

  for (const para of cell.text.paragraphs) {
    const p = document.createElement('p');
    p.style.margin = '0';

    if (para.align) {
      p.style.textAlign = para.align;
    }
    if (para.lineSpacing) {
      // Floor above zero — an untrusted spcPct/spcPts of 0 would otherwise
      // collapse cell text to zero line height (hidden-text spoofing),
      // matching the guard in TextRenderer.
      p.style.lineHeight =
        para.lineSpacing.type === 'multiple'
          ? String(Math.max(para.lineSpacing.value * SINGLE_LINE_SPACING, MIN_LINE_HEIGHT))
          : `${Math.max(para.lineSpacing.px, MIN_LINE_HEIGHT_PX)}px`;
    }

    if (para.runs.length === 0) {
      p.style.minHeight = '1em';
      p.appendChild(document.createTextNode(' '));
      paragraphs.push(p);
      continue;
    }

    for (const run of para.runs) {
      if (run.breakBefore) {
        p.appendChild(document.createElement('br'));
      }

      let node: HTMLElement = document.createElement('span');
      node.textContent = run.text;

      if (run.fontFamily) {
        node.style.fontFamily = `"${sanitizeFontFamily(run.fontFamily)}", sans-serif`;
      }
      if (run.fontSize) {
        node.style.fontSize = `${run.fontSize}px`;
      }
      if (run.color) {
        node.style.color = colorToCss(run.color);
      }

      // Wrap in formatting elements (innermost = text span)
      if (run.bold) node = wrap(node, 'strong');
      if (run.italic) node = wrap(node, 'em');
      if (run.underline) node = wrap(node, 'u');
      if (run.strikethrough) node = wrap(node, 's');

      p.appendChild(node);
    }

    paragraphs.push(p);
  }

  return paragraphs;
}

/**
 * Wraps an element in a new element of the given tag.
 */
function wrap(node: HTMLElement, tag: string): HTMLElement {
  const outer = document.createElement(tag);
  outer.appendChild(node);
  return outer;
}

/**
 * Formats a border stroke as CSS border value.
 */
function formatBorder(stroke: Stroke): string {
  const width = Math.max(1, Math.round(stroke.width));
  const color = colorToCss(stroke.color);
  return `${width}px solid ${color}`;
}
