import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import io
import barcode
from barcode.writer import ImageWriter
from PIL import Image as PILImage, ImageDraw, ImageFont
import config

FONT_SIZE    = 22   # text size in pixels — NOT scaled down
LINE_SPACING = 5


def _to_clean_str(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s.endswith(".0") and not s.startswith("0"):
        try:
            return str(int(float(s)))
        except (ValueError, TypeError):
            pass
    return s


def make_barcode(sku: str, item_name: str = "") -> PILImage.Image | None:
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


def _load_font(size):
    for path in [
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/Arial.ttf",
        "arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def _wrap_text(text, font, max_width, draw):
    words = text.split()
    lines, current = [], ""
    for word in words:
        test = (current + " " + word).strip()
        if draw.textbbox((0, 0), test, font=font)[2] <= max_width:
            current = test
        else:
            if current:
                lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines if lines else [text]


def _render(code_value: str, label_text: str) -> PILImage.Image:
    target_w = config.BARCODE_WIDTH_PX  # e.g. 280px — final output width

    # ── 1. Generate barcode at native resolution ─────────────────────────
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

    # ── 2. Scale ONLY the barcode bars to target width ───────────────────
    # Text is drawn AFTER scaling so it stays at full FONT_SIZE
    scale          = target_w / bar_w
    scaled_bar_h   = int(bar_h * scale)
    barcode_scaled = barcode_img.resize((target_w, scaled_bar_h), PILImage.LANCZOS)

    # ── 3. Measure label text at full font size ───────────────────────────
    font     = _load_font(FONT_SIZE)
    tmp      = PILImage.new("RGB", (target_w, 1), "white")
    tmp_draw = ImageDraw.Draw(tmp)
    lines    = _wrap_text(label_text, font, target_w - 10, tmp_draw)
    line_h   = tmp_draw.textbbox((0, 0), "Ag", font=font)[3]
    label_h  = len(lines) * (line_h + LINE_SPACING) + 8

    # ── 4. Compose: scaled barcode + full-size label ──────────────────────
    total_h   = scaled_bar_h + label_h
    final_img = PILImage.new("RGB", (target_w, total_h), "white")
    final_img.paste(barcode_scaled, (0, 0))

    draw = ImageDraw.Draw(final_img)
    y    = scaled_bar_h + 4

    for line in lines:
        bbox   = draw.textbbox((0, 0), line, font=font)
        text_w = bbox[2] - bbox[0]
        text_x = max(0, (target_w - text_w) // 2)
        draw.text((text_x, y), line, fill="black", font=font)
        y += line_h + LINE_SPACING

    return final_img   # already at target_w — NO final resize