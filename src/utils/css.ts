/**
 * Sanitization helpers for values that flow from untrusted PPTX content
 * into CSS or DOM attributes.
 */

/**
 * Sanitizes a font family name from PPTX content for use in CSS.
 *
 * One policy, used everywhere a font name is emitted (`@font-face`
 * definitions and `font-family` references) — if definition and reference
 * sites sanitized differently, embedded fonts would silently stop
 * matching. Keeps Unicode names ("メイリオ", "Café") intact; strips only
 * characters that can affect CSS structure inside a double-quoted string
 * (quotes, backslashes, braces, semicolons, control characters).
 *
 * @param name - Font family name from PPTX content
 * @returns Name safe to embed inside a double-quoted CSS string
 */
export function sanitizeFontFamily(name: string): string {
  // eslint-disable-next-line no-control-regex
  return name.replace(/[\x00-\x1f\x7f"'\\{};]/g, '');
}

/** URL schemes allowed for hyperlinks in rendered slides. */
const SAFE_LINK_SCHEMES = ['http:', 'https:', 'mailto:', 'tel:'];

/**
 * Checks whether a hyperlink target from PPTX content is safe to assign
 * to an anchor's href. Blocks `javascript:` and other script-capable
 * schemes — a malicious deck can put anything in a relationship target.
 *
 * @param url - Hyperlink target from the relationships part
 * @returns true if the URL parses and uses an allowed scheme
 */
export function isSafeLinkUrl(url: string): boolean {
  try {
    const parsed = new URL(url);
    return SAFE_LINK_SCHEMES.includes(parsed.protocol);
  } catch {
    // Not an absolute URL — internal slide jumps etc. are not rendered
    // as links, so reject anything unparseable rather than guessing.
    return false;
  }
}
