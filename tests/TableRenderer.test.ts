import { describe, it, expect } from 'vitest';
import { renderTable } from '../src/renderer/TableRenderer';
import type { TableElement, TableCell } from '../src/core/types';

function makeTable(cell: TableCell): TableElement {
  return {
    id: 't1',
    type: 'table',
    bounds: { x: 0, y: 0, width: 200, height: 100 },
    columnWidths: [200],
    rows: [{ height: 50, cells: [cell] }],
  };
}

function cellWithRun(run: Record<string, unknown>): TableCell {
  return {
    text: {
      paragraphs: [{ runs: [run as never] }],
    },
  };
}

describe('TableRenderer', () => {
  it('renders run text literally — markup in text does not become elements', () => {
    const fo = renderTable(makeTable(cellWithRun({ text: '<img src=x onerror=alert(1)>' })));
    expect(fo.querySelector('img')).toBeNull();
    expect(fo.textContent).toContain('<img src=x onerror=alert(1)>');
  });

  it('does not let a malicious font name break out of the style attribute', () => {
    const fo = renderTable(
      makeTable(
        cellWithRun({
          text: 'hi',
          fontFamily: '"></style><img src=x onerror=alert(1)><span style="',
        })
      )
    );
    expect(fo.querySelector('img')).toBeNull();
    expect(fo.textContent).toContain('hi');
  });

  it('does not let a malicious color value inject markup', () => {
    const fo = renderTable(
      makeTable(
        cellWithRun({
          text: 'hi',
          color: { hex: '"><script>alert(1)</script>', alpha: 1 },
        })
      )
    );
    expect(fo.querySelector('script')).toBeNull();
  });

  it('applies run formatting through nested elements', () => {
    const fo = renderTable(
      makeTable(cellWithRun({ text: 'styled', bold: true, italic: true, fontSize: 18 }))
    );
    const strong = fo.querySelector('strong');
    const em = fo.querySelector('em');
    const span = fo.querySelector('span');
    expect(strong).not.toBeNull();
    expect(em).not.toBeNull();
    expect(span!.style.fontSize).toBe('18px');
    expect(fo.textContent).toBe('styled');
  });

  it('renders hard line breaks between runs', () => {
    const fo = renderTable(
      makeTable({
        text: {
          paragraphs: [{ runs: [{ text: 'a' }, { text: 'b', breakBefore: true }] }],
        },
      })
    );
    expect(fo.querySelector('br')).not.toBeNull();
  });

  it('applies cell paragraph line spacing', () => {
    const fo = renderTable(
      makeTable({
        text: {
          paragraphs: [
            { runs: [{ text: 'a' }], lineSpacing: { type: 'multiple', value: 1.5 } },
            { runs: [{ text: 'b' }], lineSpacing: { type: 'exact', px: 24 } },
          ],
        },
      })
    );
    const ps = fo.querySelectorAll('p');
    expect(parseFloat(ps[0].style.lineHeight)).toBeCloseTo(1.8);
    expect(ps[1].style.lineHeight).toBe('24px');
  });

  it('renders empty cell paragraphs with preserved height and line spacing', () => {
    const fo = renderTable(
      makeTable({
        text: {
          paragraphs: [{ runs: [], lineSpacing: { type: 'exact', px: 32 } }],
        },
      })
    );
    const p = fo.querySelector('p')!;
    expect(p.style.minHeight).toBe('1em');
    expect(p.style.lineHeight).toBe('32px');
    expect(p.textContent!.length).toBeGreaterThan(0);
  });

  it('wraps underline and strikethrough runs with the text span innermost', () => {
    const fo = renderTable(
      makeTable(cellWithRun({ text: 'x', underline: true, strikethrough: true }))
    );
    expect(fo.querySelector('u')).not.toBeNull();
    expect(fo.querySelector('s')).not.toBeNull();
    expect(fo.querySelector('span')!.textContent).toBe('x');
  });
});
