import { describe, it, expect } from 'vitest';
import {
  parseHexColor,
  parseOoxmlAlpha,
  resolveThemeColor,
  resolvePresetColor,
  applyLuminance,
  colorToCss,
  DEFAULT_THEME_COLORS,
  PRESET_COLORS,
} from '../src/utils/color';
import type { ThemeColors } from '../src/core/types';

describe('Color Utilities', () => {
  describe('parseHexColor', () => {
    it('parses 6-digit hex color', () => {
      const color = parseHexColor('FF0000');
      expect(color.hex).toBe('#FF0000');
      expect(color.alpha).toBe(1);
    });

    it('removes # prefix if present', () => {
      const color = parseHexColor('#00FF00');
      expect(color.hex).toBe('#00FF00');
    });

    it('converts to uppercase', () => {
      const color = parseHexColor('aabbcc');
      expect(color.hex).toBe('#AABBCC');
    });

    it('accepts custom alpha', () => {
      const color = parseHexColor('0000FF', 0.5);
      expect(color.hex).toBe('#0000FF');
      expect(color.alpha).toBe(0.5);
    });

    it('clamps alpha to 0-1 range', () => {
      expect(parseHexColor('000000', -0.5).alpha).toBe(0);
      expect(parseHexColor('000000', 1.5).alpha).toBe(1);
    });
  });

  describe('parseOoxmlAlpha', () => {
    it('converts 100000 to 1 (fully opaque)', () => {
      expect(parseOoxmlAlpha(100000)).toBe(1);
    });

    it('converts 50000 to 0.5 (50% opacity)', () => {
      expect(parseOoxmlAlpha(50000)).toBe(0.5);
    });

    it('converts 0 to 0 (fully transparent)', () => {
      expect(parseOoxmlAlpha(0)).toBe(0);
    });
  });

  describe('resolveThemeColor', () => {
    const theme: ThemeColors = {
      dark1: '#111111',
      dark2: '#222222',
      light1: '#EEEEEE',
      light2: '#DDDDDD',
      accent1: '#FF0000',
      accent2: '#00FF00',
      accent3: '#0000FF',
      accent4: '#FFFF00',
      accent5: '#FF00FF',
      accent6: '#00FFFF',
      hlink: '#0000CC',
      folHlink: '#660066',
    };

    it('resolves dk1 to dark1', () => {
      expect(resolveThemeColor('dk1', theme)).toBe('#111111');
    });

    it('resolves lt1 to light1', () => {
      expect(resolveThemeColor('lt1', theme)).toBe('#EEEEEE');
    });

    it('resolves accent colors', () => {
      expect(resolveThemeColor('accent1', theme)).toBe('#FF0000');
      expect(resolveThemeColor('accent2', theme)).toBe('#00FF00');
    });

    it('resolves tx1 (text) to dark1', () => {
      expect(resolveThemeColor('tx1', theme)).toBe('#111111');
    });

    it('resolves bg1 (background) to light1', () => {
      expect(resolveThemeColor('bg1', theme)).toBe('#EEEEEE');
    });

    it('resolves hlink', () => {
      expect(resolveThemeColor('hlink', theme)).toBe('#0000CC');
    });

    it('returns black for unknown color', () => {
      expect(resolveThemeColor('unknownColor', theme)).toBe('#000000');
    });
  });

  describe('resolvePresetColor', () => {
    it('resolves red', () => {
      expect(resolvePresetColor('red')).toBe('#FF0000');
    });

    it('resolves blue', () => {
      expect(resolvePresetColor('blue')).toBe('#0000FF');
    });

    it('resolves white', () => {
      expect(resolvePresetColor('white')).toBe('#FFFFFF');
    });

    it('resolves black', () => {
      expect(resolvePresetColor('black')).toBe('#000000');
    });

    it('resolves camelCase colors', () => {
      expect(resolvePresetColor('darkBlue')).toBe('#00008B');
      expect(resolvePresetColor('lightGray')).toBe('#D3D3D3');
    });

    it('returns black for unknown color', () => {
      expect(resolvePresetColor('unknownColor')).toBe('#000000');
    });
  });

  describe('applyLuminance', () => {
    it('returns same color with no modification', () => {
      expect(applyLuminance('#FF0000', 100000, 0)).toBe('#FF0000');
    });

    it('darkens color with lumMod < 100000', () => {
      // 50% luminance mod on white should give gray
      const result = applyLuminance('#FFFFFF', 50000, 0);
      expect(result).toBe('#808080');
    });

    it('applies luminance offset', () => {
      // Adding 50% offset to black
      const result = applyLuminance('#000000', 100000, 50000);
      expect(result).toBe('#808080');
    });

    it('clamps values to 0-255 range', () => {
      // Try to exceed 255
      const result = applyLuminance('#FFFFFF', 200000, 0);
      expect(result).toBe('#FFFFFF');
    });
  });

  describe('colorToCss', () => {
    it('returns hex for fully opaque color', () => {
      const color = { hex: '#FF0000', alpha: 1 };
      expect(colorToCss(color)).toBe('#FF0000');
    });

    it('returns rgba for semi-transparent color', () => {
      const color = { hex: '#FF0000', alpha: 0.5 };
      expect(colorToCss(color)).toBe('rgba(255, 0, 0, 0.5)');
    });

    it('parses hex correctly for rgba', () => {
      const color = { hex: '#1A2B3C', alpha: 0.75 };
      expect(colorToCss(color)).toBe('rgba(26, 43, 60, 0.75)');
    });
  });

  describe('DEFAULT_THEME_COLORS', () => {
    it('has all required theme color keys', () => {
      expect(DEFAULT_THEME_COLORS).toHaveProperty('dark1');
      expect(DEFAULT_THEME_COLORS).toHaveProperty('light1');
      expect(DEFAULT_THEME_COLORS).toHaveProperty('accent1');
      expect(DEFAULT_THEME_COLORS).toHaveProperty('hlink');
    });
  });

  describe('PRESET_COLORS', () => {
    it('contains standard web colors', () => {
      expect(PRESET_COLORS.red).toBe('#FF0000');
      expect(PRESET_COLORS.green).toBe('#008000');
      expect(PRESET_COLORS.blue).toBe('#0000FF');
    });

    it('contains extended color palette', () => {
      expect(PRESET_COLORS.coral).toBeDefined();
      expect(PRESET_COLORS.salmon).toBeDefined();
      expect(PRESET_COLORS.turquoise).toBeDefined();
    });
  });
});
