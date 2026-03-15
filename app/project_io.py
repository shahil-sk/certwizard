"""
Project file save / load (the .certwiz JSON format).
No UI calls here — caller is responsible for showing dialogs.
"""
import json
from datetime import datetime
from typing import Any


def serialise(
    template_path: str | None,
    excel_path: str | None,
    color_space: str,
    positions: dict,
    fields: list,
    font_settings: dict,
    field_vars: dict,
) -> dict:
    """Return a plain dict ready for json.dump."""
    return {
        "version":        "2.1",
        "last_modified":  datetime.now().isoformat(),
        "template_path":  template_path,
        "excel_path":     excel_path,
        "color_space":    color_space,
        "positions":      positions,
        "field_settings": {
            f: {
                "size":      font_settings[f]["size"].get(),
                "color":     font_settings[f]["color"].get(),
                "visible":   field_vars[f].get(),
                "font_name": font_settings[f]["font_name"].get(),
            }
            for f in fields
        },
    }


def save(path: str, data: dict) -> None:
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2)


def load(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)
