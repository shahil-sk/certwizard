"""
Certificate generation worker.
Runs on a background thread; communicates back via callbacks.
Features:
  - PDF output via fpdf
  - Configurable filename pattern with {field} and {serial} tokens
  - Duplicate detection (warns before generating)
"""
import io
import os
import re
import threading
from collections import Counter

from fpdf import FPDF
from PIL import Image

from app.helpers import px_to_mm, safe_filename
from app.image_renderer import draw_text_on_image

_TOKEN_RE = re.compile(r"\{(\w+)\}")


def _build_filename(pattern: str, student: dict, idx: int, fields: list) -> str:
    if pattern and pattern.strip():
        def _repl(m):
            key = m.group(1).lower()
            if key == "serial":
                return str(idx + 1).zfill(3)
            val = student.get(key, "")
            return safe_filename(val) if val else m.group(0)
        name = _TOKEN_RE.sub(_repl, pattern.strip())
        return name or f"cert_{idx + 1}"
    parts = [student.get(fields[i], "") for i in range(min(2, len(fields)))]
    return safe_filename(*parts) or f"cert_{idx + 1}"


def _find_duplicates(excel_data: list, fields: list) -> list:
    if not fields:
        return []
    key    = fields[0]
    counts = Counter(r.get(key, "").strip() for r in excel_data)
    return [(v, c, key) for v, c in counts.items() if c > 1 and v]


def inject_serial(excel_data: list) -> list:
    """
    Return a copy of excel_data with a 'serial' key added to each row.
    Does not modify the original list.
    """
    return [{**row, "serial": str(i + 1).zfill(3)}
            for i, row in enumerate(excel_data)]


def run(
    excel_data: list,
    fields: list,
    field_vars: dict,
    font_settings: dict,
    available_fonts: dict,
    original_image: Image.Image,
    positions: dict,
    out_dir: str,
    color_mode: str,
    on_progress,
    on_log,
    on_done,
    lock: threading.Lock,
    filename_pattern: str = "",
) -> None:
    sub     = color_mode
    total   = len(excel_data)
    iw, ih  = original_image.size
    pdf_w   = px_to_mm(iw)
    pdf_h   = px_to_mm(ih)
    out_sub = os.path.join(out_dir, sub)
    os.makedirs(out_sub, exist_ok=True)

    # Inject serial numbers so {serial} works in both renderer and filenames
    enriched = inject_serial(excel_data)
    all_fields = fields + (["serial"] if "serial" not in fields else [])

    def _worker():
        count = 0
        on_log("Starting generation...", True)
        on_log(f"Records: {total}   Mode: {sub}   Output: {out_dir}", False)
        if filename_pattern:
            on_log(f"Filename pattern: {filename_pattern}", False)

        dupes = _find_duplicates(excel_data, fields)
        if dupes:
            on_log(f"[warn] {len(dupes)} duplicate(s) in '{fields[0]}':", False)
            for val, cnt, _ in dupes[:5]:
                on_log(f"  • '{val}'  x{cnt}", False)
            if len(dupes) > 5:
                on_log(f"  ... and {len(dupes)-5} more", False)

        on_log("-" * 44, False)

        for idx, student in enumerate(enriched):
            try:
                img = draw_text_on_image(
                    original_image.copy().convert("RGB"),
                    all_fields, field_vars, font_settings,
                    available_fonts, student, positions,
                )
                buf = io.BytesIO()
                img.save(buf, format="PNG", optimize=True)
                buf.seek(0)

                pdf = FPDF(unit="mm", format=(pdf_w, pdf_h))
                pdf.add_page()
                pdf.image(buf, x=0, y=0, w=pdf_w, h=pdf_h)

                name = _build_filename(filename_pattern, student, idx, fields)
                dest = os.path.join(out_sub, f"{name}_certificate.pdf")
                pdf.output(dest)
                count += 1
                on_log(f"[{idx+1}/{total}]  {name}_certificate.pdf", False)
            except Exception as exc:
                on_log(f"[error]  cert {idx+1}: {exc}", False)

            on_progress((idx + 1) / total * 100)

        on_log("-" * 44, False)
        on_log(f"Done  {count}/{total} certificates saved.", False)
        on_done(count, total)
        lock.release()

    threading.Thread(target=_worker, daemon=True).start()
