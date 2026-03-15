<![CDATA[<div align="center">

# CertWizard

**Bulk certificate generation, done right.**

[![Python](https://img.shields.io/badge/Python-3.7%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green?style=flat-square)](LICENSE)
[![Stars](https://img.shields.io/github/stars/shahil-sk/certwizard?style=flat-square&color=yellow)](https://github.com/shahil-sk/certwizard/stargazers)
[![Forks](https://img.shields.io/github/forks/shahil-sk/certwizard?style=flat-square&color=blue)](https://github.com/shahil-sk/certwizard/network/members)
[![Issues](https://img.shields.io/github/issues/shahil-sk/certwizard?style=flat-square)](https://github.com/shahil-sk/certwizard/issues)
[![Last Commit](https://img.shields.io/github/last-commit/shahil-sk/certwizard?style=flat-square)](https://github.com/shahil-sk/certwizard/commits/master)
[![Made with Tkinter](https://img.shields.io/badge/GUI-Tkinter-informational?style=flat-square)](https://docs.python.org/3/library/tkinter.html)

A desktop application that generates professional PDF certificates in bulk from a PNG template and an Excel data file.

</div>

---

## What it does

You bring a PNG certificate template and an `.xlsx` file with recipient data. CertWizard lets you drag each field (name, ID, date, etc.) onto the template, style the text, preview the result, and then generate a PDF for every row in the sheet — in one click.

---

## Features

- **Drag-and-drop field placement** — position text fields directly on the canvas, no coordinate guessing
- **Live preview** — see a real sample certificate before committing to a full batch run
- **RGB and CMYK color spaces** — switch modes depending on whether you are printing or going digital
- **Custom font support** — drop any `.ttf` or `.otf` file into `fonts/` and it appears in the selector
- **Batch PDF generation** — one output PDF per row, named from the first two fields (e.g. `Alice_101_certificate.pdf`)
- **Project save/load** — serialize the entire session (template path, field styles, Excel path) to a `.certwiz` file
- **Scrollable canvas** — horizontal and vertical scrollbars with mousewheel support on Windows, macOS, and Linux
- **Background generation thread** — the UI stays responsive with a real-time progress bar during batch runs

---

## Tech stack

| Library | Role |
|---|---|
| [Pillow](https://python-pillow.org/) `>=10,<13` | Image rendering and PNG export |
| [openpyxl](https://openpyxl.readthedocs.io/) `>=3.1,<4` | Reading `.xlsx` data files |
| [fpdf2](https://py-pdf.github.io/fpdf2/) `>=2.7,<3` | Writing output PDFs |
| Tkinter (stdlib) | Desktop GUI |

---

## Project structure

```
certwizard/
├── main.py                  # entry point
├── requirements.txt
├── certgen.ico / certgen.png
├── fonts/                   # drop .ttf / .otf files here
├── testfiles/               # sample template and dummy data
└── app/
    ├── constants.py         # theme palette and app-wide literals
    ├── helpers.py           # pure utility functions
    ├── font_manager.py      # font scanning with lru_cache loader
    ├── image_renderer.py    # all PIL drawing and in-memory PNG export
    ├── excel_loader.py      # openpyxl ingestion, returns plain dicts
    ├── generator.py         # background thread worker, no direct UI calls
    ├── project_io.py        # .certwiz JSON serialise / save / load
    ├── core.py              # top-level controller wiring all components
    └── ui/
        ├── navbar.py
        ├── status_bar.py
        ├── control_panel.py
        ├── field_list.py
        ├── canvas.py
        └── dialogs.py
```

---

## Getting started

### Prerequisites

- Python 3.7 or higher
- pip

### Install

```bash
git clone https://github.com/shahil-sk/certwizard.git
cd certwizard
pip install -r requirements.txt
```

### Add fonts (optional)

Place any `.ttf` or `.otf` font files in the `fonts/` directory. The app detects them automatically at startup.

### Run

```bash
python main.py
```

---

## Usage

1. **Load Template** — select a PNG certificate background (landscape or portrait)
2. **Load Excel** — select an `.xlsx` file; the first row must contain field names
3. **Position fields** — drag each field label onto the canvas where the text should appear
4. **Style fields** — set font, size, and color per field; toggle visibility with the inline checkbox
5. **Preview** — click Preview to render one certificate with sample data
6. **Generate** — click Generate, pick a color space (RGB or CMYK) and output folder, then wait for the progress bar to finish
7. **Save project** — use the Project menu to save a `.certwiz` file and reload the session later

---

## File formats

| Type | Format | Notes |
|---|---|---|
| Template | `.png` | Transparent or solid backgrounds both work |
| Data | `.xlsx` | First row = field names, data from row 2 onward |
| Project | `.certwiz` | JSON file storing all settings and field positions |
| Output | `.pdf` | One file per row, named from the first two fields |

---

## Troubleshooting

**Fonts not showing up** — confirm the files are `.ttf` or `.otf` and are placed directly inside the `fonts/` directory (not in a subfolder).

**Excel not loading** — the first row must contain header names; data rows start from row 2.

**Template not displaying** — only PNG files are supported. Convert other formats with an image editor first.

**Output PDFs missing** — check that the output directory exists and that the current user has write permission.

---

## Changelog

See [CHANGES.md](CHANGES.md) for the full version history.

---

## License

[MIT](LICENSE)

---

## Credits

Developed by [Shahil SK](https://github.com/shahil-sk).
Uses Pillow, openpyxl, fpdf2, and Tkinter.
]]>