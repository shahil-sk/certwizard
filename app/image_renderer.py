"""
All PIL drawing operations live here.
No tkinter widgets are created in this module.
"""
import io

from PIL import Image, ImageDraw

from app.helpers import hex_to_rgb
from app.font_manager import resolve_font


def draw_text_on_image(
    img: Image.Image,
    fields: list,
    field_vars: dict,
    font_settings: dict,
    available_fonts: dict,
    student: dict,
    positions: dict,
) -> Image.Image:
    """Render all visible field text onto a copy of img and return it."""
    draw = ImageDraw.Draw(img)
    for field in fields:
        if not field_vars.get(field, lambda: False)():
            continue
        if field not in positions:
            continue
        try:
            x, y   = positions[field]
            s      = font_settings[field]
            size   = s["size"].get()
            color  = s["color"].get()
            fname  = s["font_name"].get()
            font   = resolve_font(available_fonts, fname, size)
            text   = student.get(field, "")
            tw     = draw.textlength(text, font=font)
            try:
                bbox = font.getbbox(text)
                th   = bbox[3] - bbox[1]
                yo   = (size - th) // 2
            except Exception:
                yo = 0
            draw.text(
                (x - tw / 2, y - size / 2 + yo),
                text, font=font, fill=hex_to_rgb(color),
            )
        except Exception as exc:
            print(f"[renderer] {field}: {exc}")
    return img


def render_placeholder(
    field: str,
    font_settings: dict,
    available_fonts: dict,
    excel_data: list,
    scale_x: float,
    scale_y: float,
):
    """Return a PIL Image of the sample text at canvas scale."""
    s     = font_settings[field]
    size  = s["size"].get()
    color = s["color"].get()
    fname = s["font_name"].get()
    font  = resolve_font(available_fonts, fname, size)
    text  = (excel_data[0].get(field, field) if excel_data else field) or field

    tmp  = Image.new("RGBA", (1, 1))
    draw = ImageDraw.Draw(tmp)
    tw   = max(int(draw.textlength(text, font=font)), 1)
    try:
        asc, desc = font.getmetrics()
        th = max(asc + desc, 1)
    except Exception:
        th = max(size, 1)

    img  = Image.new("RGBA", (tw, th), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    try:
        bbox = font.getbbox(text)
        yo   = (th - (bbox[3] - bbox[1])) // 2
    except Exception:
        yo = 0
    draw.text((0, yo), text, font=font, fill=hex_to_rgb(color))

    sw = max(int(tw / scale_x), 1)
    sh = max(int(th / scale_y), 1)
    return img.resize((sw, sh), Image.LANCZOS)


def image_to_pdf_bytes(img: Image.Image) -> bytes:
    """Save a PIL image to an in-memory PNG buffer and return the raw bytes."""
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PNG", optimize=True)
    buf.seek(0)
    return buf.read()
