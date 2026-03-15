"""
Certificate generation worker.
Runs on a background thread; communicates back via callbacks.
"""
import io
import os
import re
import threading

from fpdf import FPDF
from PIL import Image

from app.helpers import px_to_mm, safe_filename
from app.image_renderer import draw_text_on_image


def run(
    excel_data: list,
    fields: list,
    field_vars: dict,
    font_settings: dict,
    available_fonts: dict,
    original_image: Image.Image,
    positions: dict,
    out_dir: str,
    color_mode: str,          # "RGB" or "CMYK"
    on_progress,              # callable(pct: float)
    on_log,                   # callable(msg: str, clear: bool)
    on_done,                  # callable(count: int, total: int)
    lock: threading.Lock,
) -> None:
    """Start generation on a daemon thread and return immediately."""

    sub     = color_mode
    total   = len(excel_data)
    iw, ih  = original_image.size
    pdf_w   = px_to_mm(iw)
    pdf_h   = px_to_mm(ih)
    out_sub = os.path.join(out_dir, sub)
    os.makedirs(out_sub, exist_ok=True)

    def _worker():
        count = 0
        on_log("Starting generation...", True)
        on_log(f"Records: {total}   Mode: {sub}   Output: {out_dir}", False)
        on_log("-" * 44, False)

        for idx, student in enumerate(excel_data):
            try:
                img = draw_text_on_image(
                    original_image.copy().convert("RGB"),
                    fields, field_vars, font_settings,
                    available_fonts, student, positions,
                )

                buf = io.BytesIO()
                img.save(buf, format="PNG", optimize=True)
                buf.seek(0)

                pdf = FPDF(unit="mm", format=(pdf_w, pdf_h))
                pdf.add_page()
                pdf.image(buf, x=0, y=0, w=pdf_w, h=pdf_h)

                parts = [
                    student.get(fields[i], "")
                    for i in range(min(2, len(fields)))
                ]
                name = safe_filename(*parts) or f"cert_{idx + 1}"
                dest = os.path.join(out_sub, f"{name}_certificate.pdf")
                pdf.output(dest)
                count += 1
                on_log(f"[{idx + 1}/{total}]  {name}_certificate.pdf", False)
            except Exception as exc:
                on_log(f"[error]  cert {idx + 1}: {exc}", False)

            on_progress((idx + 1) / total * 100)

        on_log("-" * 44, False)
        on_log(f"Done  {count}/{total} certificates saved.", False)
        on_done(count, total)
        lock.release()

    threading.Thread(target=_worker, daemon=True).start()
