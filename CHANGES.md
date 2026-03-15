# Changelog

All notable changes to CertWizard are documented here.

---

## v2.1 (2026-03-15)

### Package restructure
- Monolithic `certgen.py` split into a maintainable `app/` package.
- Entry point changed to `main.py`.
- Module layout:
  - `app/constants.py` — theme palette and app-wide literals
  - `app/helpers.py` — pure utility functions, safe to unit-test
  - `app/font_manager.py` — font scanning with `lru_cache` loader
  - `app/image_renderer.py` — all PIL drawing and in-memory PNG export
  - `app/excel_loader.py` — openpyxl ingestion returning plain dicts
  - `app/generator.py` — background thread worker with no direct UI calls
  - `app/project_io.py` — `.certwiz` JSON serialise/save/load
  - `app/core.py` — top-level controller wiring all components
  - `app/ui/` — navbar, status bar, control panel, field list, canvas, dialogs

### UI overhaul
- Dark navy navbar replacing the old teal header.
- Flat bordered card panels instead of raised `LabelFrame` widgets.
- Alternating row tints in the field list for easier scanning.
- Visibility checkboxes inline in each field row header.
- Status bar driven by `textvariable`, no `.config()` on every update.
- Fonts changed to Segoe UI / Consolas throughout.

### Bug fixes
- Fixed double Color-button creation in field rows (dead widget loop removed).
- Fixed `field_vars` visibility check in renderer: now uses `.get()` on
  `BooleanVar` instead of incorrectly calling the var as a function.

### Performance
- Font loading now cached via `functools.lru_cache(maxsize=64)` in
  `font_manager.py`. No more repeated `ImageFont.truetype()` disk calls
  during preview or batch generation.
- Certificate generation writes PNG to a `BytesIO` buffer instead of a
  temp file on disk, removing the `try/finally os.remove` pattern.

### Canvas
- Horizontal and vertical scrollbars added to the canvas area.
- Mousewheel scroll support for Windows, macOS, and Linux.
- Scroll region updates automatically when a template is loaded.

### Build
- `build.spec` entry point updated to `main.py`.
- Removed `email`, `html`, `http`, `urllib` from excludes — these are
  required internally by fpdf2 and excluding them caused silent build failures.
- Added `xmlrpc`, `pydoc`, `distutils`, `lib2to3` to excludes (genuinely unused).

### Dependencies
- Pinned upper bounds: `pillow<13`, `openpyxl<4`, `fpdf2<3` to prevent
  silent breaking upgrades.

---

## v2.0

### Changes
- Theme colors extracted into a single `C` dict, reducing repeated string literals.
- `_btn` and `_label` helper functions to consolidate repetitive widget creation.
- `os.makedirs(exist_ok=True)` replaces try/except makedirs pattern.
- `_get_scaled_positions` rewritten as a single dict comprehension.
- Legacy aliases preserved for `.certwiz` project file compatibility.

---

## v1.0

- Initial release.
- Tkinter GUI for certificate generation from PNG template + Excel data.
- Draggable text placeholders on canvas.
- CMYK and RGB color mode selection.
- PDF output via fpdf2.
- Project save/load as `.certwiz` JSON.
