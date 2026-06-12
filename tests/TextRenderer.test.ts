import { describe, it, expect } from 'vitest';
import { renderTextBody } from '../src/renderer/TextRenderer';
import type { TextBody, Paragraph } from '../src/core/types';

function render(paragraphs: Paragraph[], autofit?: TextBody['autofit']): HTMLElement {
  const container = document.createElement('div');
  renderTextBody({ paragraphs, autofit }, container);
  return container;
}

describe('TextRenderer', () => {
  describe('alignment', () => {
    it('centers text via text-align', () => {
      const container = render([{ runs: [{ text: 'Title' }], align: 'center' }]);
      const p = container.querySelector('p')!;
      expect(p.style.textAlign).toBe('center');
      expect(p.style.display).not.toBe('flex');
    });

    it('right-aligns and justifies text', () => {
      const container = render([
        { runs: [{ text: 'a' }], align: 'right' },
        { runs: [{ text: 'b' }], align: 'justify' },
      ]);
      const ps = container.querySelectorAll('p');
      expect(ps[0].style.textAlign).toBe('right');
      expect(ps[1].style.textAlign).toBe('justify');
    });
  });

  describe('hard line breaks', () => {
    it('renders breakBefore runs after a <br>', () => {
      const container = render([
        { runs: [{ text: 'line one' }, { text: 'line two', breakBefore: true }] },
      ]);
      const p = container.querySelector('p')!;
      const childTags = Array.from(p.childNodes).map((n) => n.nodeName.toLowerCase());
      expect(childTags).toEqual(['span', 'br', 'span']);
    });
  });

  describe('line spacing', () => {
    it('maps 100% spacing to ~1.2 line-height (single spacing)', () => {
      const container = render([
        { runs: [{ text: 'x' }], lineSpacing: { type: 'multiple', value: 1 } },
      ]);
      expect(container.querySelector('p')!.style.lineHeight).toBe('1.2');
    });

    it('applies exact spacing as absolute pixels regardless of font size', () => {
      const container = render([
        {
          runs: [{ text: 'x', fontSize: 32 }],
          lineSpacing: { type: 'exact', px: 24 },
        },
      ]);
      expect(container.querySelector('p')!.style.lineHeight).toBe('24px');
    });

    it('reduces spacing by autofit lnSpcReduction even without explicit lnSpc', () => {
      const container = render(
        [{ runs: [{ text: 'x' }] }],
        { type: 'normal', fontScale: 1, lineSpacingReduction: 0.2 }
      );
      // 1.2 single spacing * (1 - 0.2)
      expect(parseFloat(container.querySelector('p')!.style.lineHeight)).toBeCloseTo(0.96);
    });
  });

  describe('paragraph spacing', () => {
    it('applies percent space-before relative to the largest run size', () => {
      const container = render([
        {
          runs: [{ text: 'x', fontSize: 20 }, { text: 'y', fontSize: 40 }],
          spaceBefore: { type: 'percent', value: 0.5 },
        },
      ]);
      expect(container.querySelector('p')!.style.marginTop).toBe('20px');
    });

    it('applies exact space-after in pixels', () => {
      const container = render([
        { runs: [{ text: 'x' }], spaceAfter: { type: 'exact', px: 12 } },
      ]);
      expect(container.querySelector('p')!.style.marginBottom).toBe('12px');
    });
  });

  describe('indentation', () => {
    it('applies hanging indent via text-indent with a fixed-width bullet', () => {
      const container = render([
        {
          runs: [{ text: 'item' }],
          bullet: { type: 'bullet', char: '•' },
          marginLeft: 36,
          indent: -18,
        },
      ]);
      const p = container.querySelector('p')!;
      expect(p.style.marginLeft).toBe('36px');
      expect(p.style.textIndent).toBe('-18px');
      const bullet = p.querySelector('span')!;
      expect(bullet.style.width).toBe('18px');
      expect(bullet.style.display).toBe('inline-block');
    });
  });

  describe('whitespace', () => {
    it('preserves leading/trailing run spaces', () => {
      const container = render([{ runs: [{ text: 'Hello ' }, { text: 'world' }] }]);
      expect(container.textContent).toBe('Hello world');
    });
  });

  describe('hyperlinks', () => {
    it('renders safe links as anchors', () => {
      const container = render([{ runs: [{ text: 'go', link: 'https://example.com/' }] }]);
      const a = container.querySelector('a')!;
      expect(a).not.toBeNull();
      expect(a.href).toBe('https://example.com/');
    });

    it('does not render javascript: links as anchors', () => {
      const container = render([{ runs: [{ text: 'go', link: 'javascript:alert(1)' }] }]);
      expect(container.querySelector('a')).toBeNull();
      expect(container.textContent).toBe('go');
    });
  });

  describe('untrusted values', () => {
    it('clamps autofit so text cannot collapse to zero line height', () => {
      const container = render(
        [{ runs: [{ text: 'x' }], lineSpacing: { type: 'multiple', value: 1 } }],
        { type: 'normal', fontScale: 1, lineSpacingReduction: 1 }
      );
      // reduction is clamped to 0.4 → 1.2 * 0.6 = 0.72, never 0
      expect(parseFloat(container.querySelector('p')!.style.lineHeight)).toBeCloseTo(0.72);
    });

    it('floors explicit zero line spacing above zero (no hidden-text collapse)', () => {
      const multiple = render([
        { runs: [{ text: 'x' }], lineSpacing: { type: 'multiple', value: 0 } },
      ]);
      expect(parseFloat(multiple.querySelector('p')!.style.lineHeight)).toBeGreaterThan(0);

      const exact = render([
        { runs: [{ text: 'x' }], lineSpacing: { type: 'exact', px: 0 } },
      ]);
      expect(parseFloat(exact.querySelector('p')!.style.lineHeight)).toBeGreaterThan(0);
    });

    it('clamps negative space-before to zero', () => {
      const container = render([
        { runs: [{ text: 'x' }], spaceBefore: { type: 'exact', px: -40 } },
      ]);
      expect(container.querySelector('p')!.style.marginTop).toBe('0px');
    });
  });

  describe('empty paragraphs', () => {
    it('keeps empty paragraphs from collapsing', () => {
      const container = render([{ runs: [] }]);
      const p = container.querySelector('p')!;
      expect(p.querySelector('span')).not.toBeNull();
      expect(p.textContent!.length).toBeGreaterThan(0);
    });
  });

  describe('percent spacing details', () => {
    it('falls back to 16px and applies fontScale for percent spacing', () => {
      const container = document.createElement('div');
      const paragraphs = [
        { runs: [{ text: 'x' }], spaceBefore: { type: 'percent' as const, value: 0.5 } },
      ];
      renderTextBody(
        { paragraphs, autofit: { type: 'normal', fontScale: 0.5, lineSpacingReduction: 0 } },
        container
      );
      // (16 fallback) * 0.5 fontScale * 0.5 percent = 4px
      expect(container.querySelector('p')!.style.marginTop).toBe('4px');
    });
  });
});
