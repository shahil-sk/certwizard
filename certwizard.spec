# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for CertWizard.
# Build with:  pyinstaller certwizard.spec
#
# The spec approach (vs bare CLI flags) gives us explicit control over
# hidden imports and collected data so the binary is truly self-contained.

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# ---------------------------------------------------------------------------
# Hidden imports
# fpdf2 loads font metrics and svg/html sub-modules at runtime via importlib.
# openpyxl reads its shared-strings and styles XML via pkg_resources.
# Pillow registers format plugins lazily; list them explicitly so they are
# frozen into the bundle.
# ---------------------------------------------------------------------------
hidden_imports = [
    # fpdf2 internals
    "fpdf",
    "fpdf.enums",
    "fpdf.errors",
    "fpdf.fonts",
    "fpdf.image_datastructures",
    "fpdf.image_parsing",
    "fpdf.line_break",
    "fpdf.output",
    "fpdf.page_break",
    "fpdf.recorder",
    "fpdf.svg",
    "fpdf.table",
    "fpdf.transitions",
    "fpdf.util",

    # openpyxl internals loaded at runtime
    "openpyxl",
    "openpyxl.cell",
    "openpyxl.styles",
    "openpyxl.utils",
    "openpyxl.reader.excel",
    "openpyxl.reader.strings",
    "openpyxl.workbook",
    "openpyxl.worksheet",
    "openpyxl.worksheet._reader",
    "openpyxl.packaging.manifest",

    # Pillow format plugins (registered via importlib in Pillow >= 10)
    "PIL._imaging",
    "PIL.Image",
    "PIL.ImageDraw",
    "PIL.ImageFont",
    "PIL.PngImagePlugin",
    "PIL.JpegImagePlugin",
    "PIL.BmpImagePlugin",
    "PIL.TiffImagePlugin",
    "PIL.GifImagePlugin",
    "PIL.WebPImagePlugin",

    # Tkinter — frozen Python builds sometimes miss these
    "tkinter",
    "tkinter.ttk",
    "tkinter.filedialog",
    "tkinter.messagebox",
    "tkinter.colorchooser",
    "tkinter.scrolledtext",

    # stdlib used at runtime
    "io",
    "json",
    "threading",
    "collections",
    "functools",
    "pathlib",
]

# collect_submodules pulls every sub-module of a package — catches anything
# the explicit list above may miss across library patch versions.
hidden_imports += collect_submodules("fpdf")
hidden_imports += collect_submodules("openpyxl")

# ---------------------------------------------------------------------------
# Data files bundled into the binary
# ---------------------------------------------------------------------------
datas = [
    # fonts/ and testfiles/ from the repo root
    ("fonts",     "fonts"),
    ("testfiles", "testfiles"),
    # app icons
    ("icon.ico", ".")
]

# fpdf2 ships its core font files as package data; collect them so the
# binary can render text without any external files.
datas += collect_data_files("fpdf", include_py_files=False)

# openpyxl bundles XML schemas and template files.
datas += collect_data_files("openpyxl", include_py_files=False)

# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------
a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # things fpdf2 optionally uses but we do not need
        "matplotlib",
        "numpy",
        "scipy",
        "pandas",
        # test frameworks
        "pytest",
        "unittest",
        # dev tooling
        "IPython",
        "pydoc",
        "distutils",
        "lib2to3",
        "xmlrpc",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="CertWizard",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,           # stripping can break Tkinter on some Linux distros
    upx=True,              # compress with UPX if available; runner installs it
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,         # no console window — windowed GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Windows only; ignored on other platforms
    icon="icon.ico",
    version=None,
)
