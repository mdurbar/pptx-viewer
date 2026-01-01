/**
 * Renderer for tables.
 *
 * Tables are rendered as HTML tables inside SVG foreignObject elements
 * for proper text rendering and cell layout.
 */

import type { TableElement, TableRow, TableCell, CellBorders, Stroke } from '../core/types';
import { colorToCss } from '../utils/color';

/**
 * Renders a table element to an SVG foreignObject containing an HTML table.
 *
 * @param table - The table element to render
 * @returns SVG foreignObject element containing the HTML table
 */
export function renderTable(table: TableElement): SVGForeignObjectElement {
  const fo = document.createElementNS('http://www.w3.org/2000/svg', 'foreignObject');
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
    const textHtml = renderCellText(cell);
    td.innerHTML = textHtml;
  }

  return td;
}

/**
 * Renders cell text content to HTML.
 */
function renderCellText(cell: TableCell): string {
  if (!cell.text) return '';

  const paragraphs: string[] = [];

  for (const para of cell.text.paragraphs) {
    if (para.runs.length === 0) {
      paragraphs.push('<p style="margin: 0; min-height: 1em;">&nbsp;</p>');
      continue;
    }

    const runHtml: string[] = [];
    for (const run of para.runs) {
      let html = escapeHtml(run.text);

      // Apply formatting
      if (run.bold) {
        html = `<strong>${html}</strong>`;
      }
      if (run.italic) {
        html = `<em>${html}</em>`;
      }
      if (run.underline) {
        html = `<u>${html}</u>`;
      }
      if (run.strikethrough) {
        html = `<s>${html}</s>`;
      }

      // Apply inline styles
      const styles: string[] = [];
      if (run.fontFamily) {
        styles.push(`font-family: "${run.fontFamily}", sans-serif`);
      }
      if (run.fontSize) {
        styles.push(`font-size: ${run.fontSize}px`);
      }
      if (run.color) {
        styles.push(`color: ${colorToCss(run.color)}`);
      }

      if (styles.length > 0) {
        html = `<span style="${styles.join('; ')}">${html}</span>`;
      }

      runHtml.push(html);
    }

    // Paragraph styles
    const pStyles: string[] = ['margin: 0'];

    if (para.align) {
      pStyles.push(`text-align: ${para.align}`);
    }
    if (para.lineSpacing) {
      pStyles.push(`line-height: ${para.lineSpacing}`);
    }

    paragraphs.push(`<p style="${pStyles.join('; ')}">${runHtml.join('')}</p>`);
  }

  return paragraphs.join('');
}

/**
 * Formats a border stroke as CSS border value.
 */
function formatBorder(stroke: Stroke): string {
  const width = Math.max(1, Math.round(stroke.width));
  const color = colorToCss(stroke.color);
  return `${width}px solid ${color}`;
}

/**
 * Escapes HTML special characters.
 */
function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
