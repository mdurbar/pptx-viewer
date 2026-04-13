# Changelog

## 0.2.1

### Fixed

- Republish with the actual 0.2.0 code. The 0.2.0 npm tarball was published with a stale `dist/` from before the placeholder inheritance work, so installs of 0.2.0 did not include the fix. 0.2.1 ships the correct bundle.

## 0.2.0

### Added

- **Placeholder inheritance** - Slide shapes that rely on layout/master inheritance for bounds and text styling now render correctly. Previously, placeholder shapes with empty `<p:spPr>` (no `<a:xfrm>`) were silently dropped, producing blank slides for agent-generated PPTX files.
- **Text style inheritance from master `txStyles`** - Parse `titleStyle`, `bodyStyle`, and `otherStyle` from the slide master and apply per-level defaults (font size, color, weight, etc.) to slide placeholder runs that omit their own `rPr`.
- **Layout-to-master placeholder inheritance** - Layout-level placeholder shapes with no bounds now inherit from the master, enabling the full slide-layout-master chain.
- **`defPPr` fallback in list style parsing** - Masters that store text defaults in `<a:defPPr>` instead of numbered `<a:lvlNpPr>` elements are now handled correctly.

### Changed

- **`slideLayouts` and `slideMasters` maps re-keyed by file path** - Previously keyed by scoped relationship IDs that didn't match across the slide/layout/master boundary, causing lookups to return `undefined`. Now keyed by canonical file path (e.g., `ppt/slideLayouts/slideLayout1.xml`). `slide.layoutId` and `layout.masterId` are also paths. This is a breaking change for consumers that relied on the old rId-based keys, though those lookups were already silently failing.
- **`PPTXViewer` uses `renderSlideWithInheritance`** - The viewer now renders master backgrounds, layout shapes, and master decorations instead of only drawing slide-level elements on a white background.
- **`ph@type` default corrected from `body` to `obj`** - Per ECMA-376 19.3.1.36, placeholder elements with no explicit `type` attribute default to `obj` (generic object), not `body`.

### Fixed

- Placeholder shapes with empty `<p:spPr>` no longer silently dropped.
- Slides from agent-generated PPTX files (which lean heavily on layout/master inheritance) now render with visible text and correct backgrounds.
- Layout/master lookup no longer fails due to mismatched relationship ID scoping.

## 0.1.0

Initial release.
