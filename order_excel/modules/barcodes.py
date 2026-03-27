import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import io
import barcode
from barcode.writer import ImageWriter
from PIL import Image as PILImage, ImageDraw, ImageFont
import config

FONT_SIZE   = 28
LINE_SPACING = 6   # extra pixels between lines


def _to_clean_str(value) -> str:
    """
    Convert value to clean string, preserving leading zeros.
    e.g. "0407267" stays "0407267", "12345.0" becomes "12345"
    """
    if value is None:
        return ""
    s = str(value).strip()
    # Only remove .0 suffix if there are no leading zeros
    if s.endswith(".0") and not s.startswith("0"):
        try:
            return str(int(float(s)))
        except (ValueError, TypeError):
            pass
    return s


def make_barcode(sku: str, item_name: str = "") -> PILImage.Image | None:
    """
    Generate a barcode image.
    - Bars encode the SKU (or item name if no SKU)
    - Item name shown below the bars, wrapping to multiple lines if needed
    """
    sku       = _to_clean_str(sku)
    item_name = _to_clean_str(item_name)

    code_value = sku       if sku.lower()       not in ("nan", "", "none") else item_name
    label_text = item_name if item_name.lower() not in ("nan", "", "none") else sku

    if not code_value:
        return None

    try:
        return _render(code_value, label_text)
    except Exception as e:
        print(f"  [WARN] Barcode generation failed for '{code_value}': {e}")
        return None


def _load_font(size=FONT_SIZE):
    for path in ["arial.ttf", "Arial.ttf",
                 "C:/Windows/Fonts/arial.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                 "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"]:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _wrap_text(text: str, font, max_width: int, draw: ImageDraw.ImageDraw) -> list[str]:
    """Split text into lines that fit within max_width pixels."""
    words = text.split()
    lines = []
    current = ""

    for word in words:
        test = (current + " " + word).strip()
        w = draw.textbbox((0, 0), test, font=font)[2]
        if w <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    return lines if lines else [text]


def _render(code_value: str, label_text: str) -> PILImage.Image:
    # ── 1. Generate barcode bars (no built-in text) ──────────────────────
    writer = ImageWriter()
    code   = barcode.get("code128", code_value, writer=writer)
    buf    = io.BytesIO()
    code.write(buf, options={
        "module_width":  0.4,
        "module_height": 22,
        "quiet_zone":    4,
        "font_size":     0,
        "text_distance": 0,
        "write_text":    False,
        "background":    "white",
        "foreground":    "black",
    })
    buf.seek(0)
    barcode_img = PILImage.open(buf).convert("RGB")
    bar_w, bar_h = barcode_img.size

    # ── 2. Calculate how many lines the label needs ──────────────────────
    font    = _load_font(FONT_SIZE)
    # Use a temporary draw surface to measure text
    tmp     = PILImage.new("RGB", (bar_w, 1), "white")
    tmp_draw = ImageDraw.Draw(tmp)

    lines     = _wrap_text(label_text, font, bar_w - 8, tmp_draw)
    line_h    = tmp_draw.textbbox((0, 0), "Ag", font=font)[3]  # height of one line
    label_h   = len(lines) * (line_h + LINE_SPACING) + 6       # total label area height

    # ── 3. Compose final image: barcode + label area ─────────────────────
    total_h   = bar_h + label_h
    final_img = PILImage.new("RGB", (bar_w, total_h), "white")
    final_img.paste(barcode_img, (0, 0))

    draw = ImageDraw.Draw(final_img)
    y    = bar_h + 4

    for line in lines:
        bbox   = draw.textbbox((0, 0), line, font=font)
        text_w = bbox[2] - bbox[0]
        text_x = max(0, (bar_w - text_w) // 2)   # center each line
        draw.text((text_x, y), line, fill="black", font=font)
        y += line_h + LINE_SPACING

    # ── 4. Resize to configured Excel dimensions ─────────────────────────
    # Keep aspect ratio — make it tall enough for the text
    target_w = config.BARCODE_WIDTH_PX
    scale    = target_w / bar_w
    target_h = int(total_h * scale)

    final_img = final_img.resize((target_w, target_h), PILImage.LANCZOS)
    return final_img