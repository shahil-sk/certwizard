"""
Pure utility functions — no tkinter, no PIL side-effects.
Safe to import anywhere including tests.
"""
import os
import sys
import re


def resource_path(relative_path: str) -> str:
    """Return absolute path whether running from source or a PyInstaller bundle."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    # __file__ is app/helpers.py so go one level up to reach project root
    base = os.path.dirname(base) if not hasattr(sys, "_MEIPASS") else base
    return os.path.join(base, relative_path)


def px_to_mm(px: float) -> float:
    """Convert pixels (96 dpi) to millimetres."""
    from app.constants import PX_TO_MM
    return px * PX_TO_MM


def safe_filename(*parts: str) -> str:
    """Build a filesystem-safe name from one or more string parts."""
    cleaned = [
        re.sub(r"[^\w\-_. ]", "", str(p)).strip()
        for p in parts
    ]
    return "_".join(p for p in cleaned if p)


def hex_to_rgb(color) -> tuple:
    """Convert a hex string or CMYK string to an (R, G, B) tuple."""
    import tkinter as tk
    if isinstance(color, tk.StringVar):
        color = color.get()
    if color.startswith("cmyk("):
        c, m, y, k = map(float, color[5:-1].split(","))
        return (
            int(255 * (1 - c) * (1 - k)),
            int(255 * (1 - m) * (1 - k)),
            int(255 * (1 - y) * (1 - k)),
        )
    h = color.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def rgb_to_cmyk(hex_color: str) -> str:
    r, g, b = [x / 255 for x in hex_to_rgb(hex_color)]
    k = 1 - max(r, g, b)
    if k == 1:
        return "cmyk(0.00,0.00,0.00,1.00)"
    c = (1 - r - k) / (1 - k)
    m = (1 - g - k) / (1 - k)
    y = (1 - b - k) / (1 - k)
    return f"cmyk({c:.2f},{m:.2f},{y:.2f},{k:.2f})"


def cmyk_to_hex(cmyk: str) -> str:
    c, m, y, k = map(float, cmyk[5:-1].split(","))
    return "#{:02x}{:02x}{:02x}".format(
        int(255 * (1 - c) * (1 - k)),
        int(255 * (1 - m) * (1 - k)),
        int(255 * (1 - y) * (1 - k)),
    )
