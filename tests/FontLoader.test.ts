import { describe, it, expect, afterEach } from 'vitest';
import { injectFontStyles, cleanupFontStyles } from '../src/renderer/FontLoader';
import type { EmbeddedFont } from '../src/core/types';

function fontMap(name: string): Map<string, EmbeddedFont> {
  return new Map([
    [name, { name, url: 'blob:fake', format: 'truetype' } as EmbeddedFont],
  ]);
}

afterEach(() => cleanupFontStyles());

describe('injectFontStyles', () => {
  it('does not let a malicious font name escape the @font-face block', () => {
    injectFontStyles(fontMap('X"; } body { display:none } @font-face { font-family:"Y'));
    const css = document.getElementById('pptx-embedded-fonts')!.textContent!;

    // The emitted family must stay a plain double-quoted string: no
    // quotes, braces, semicolons, or backslashes can survive into it.
    const family = css.match(/font-family: "([^"\n]*)";/);
    expect(family).not.toBeNull();
    expect(family![1]).not.toMatch(/["{};\\]/);
    expect(css).not.toContain('body {');
  });

  it('keeps Unicode font names intact', () => {
    injectFontStyles(fontMap('メイリオ'));
    const css = document.getElementById('pptx-embedded-fonts')!.textContent!;

    expect(css).toContain('font-family: "メイリオ"');
  });

  it('cleanupFontStyles removes the style element', () => {
    injectFontStyles(fontMap('Test Font'));
    expect(document.getElementById('pptx-embedded-fonts')).not.toBeNull();
    cleanupFontStyles();
    expect(document.getElementById('pptx-embedded-fonts')).toBeNull();
  });
});
