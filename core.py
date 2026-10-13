import base64
import io

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font
from PIL import Image, ImageDraw, ImageFont, ImageOps

CHUNK = 45000  # Google Sheets allows 50,000 characters per cell


def load_image(file):
    img = Image.open(file)
    return ImageOps.exif_transpose(img).convert("RGB")


def prep_photo(img, rotate=0, max_side=1000, quality=75):
    """Full photo (not cropped), resized so numbers stay readable. Returns base64 text."""
    if rotate:
        img = img.rotate(rotate, expand=True)
    img = img.copy()
    img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=quality, optimize=True)
    return base64.b64encode(buf.getvalue()).decode()


def stamp_photo(img, meta):
    """Print date/time, GPS and place in the bottom-left corner of the photo."""
    meta = meta or {}
    img = img.convert("RGB")
    lines = []
    if meta.get("captured_at"):
        lines.append(str(meta["captured_at"]))
    if meta.get("lat") is not None and meta.get("lon") is not None:
        t = f"Lat {meta['lat']:.5f}  Lon {meta['lon']:.5f}"
        if meta.get("acc") is not None:
            t += f"  (+/-{meta['acc']} m)"
        lines.append(t)
    else:
        lines.append("Location not available")
    if meta.get("place"):
        lines.append(str(meta["place"]))
    size = max(14, int(img.width * 0.026))
    font = ImageFont.load_default(size=size)
    pad = max(6, size // 3)
    overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
    d = ImageDraw.Draw(overlay)
    maxw = img.width - 4 * pad
    fitted = []
    for ln in lines:
        while d.textlength(ln, font=font) > maxw and len(ln) > 4:
            ln = ln[:-4].rstrip() + "..."
        fitted.append(ln)
    line_h = int(size * 1.3)
    box_w = int(max(d.textlength(ln, font=font) for ln in fitted)) + 2 * pad
    box_h = line_h * len(fitted) + 2 * pad
    x0, y0 = pad, img.height - box_h - pad
    d.rectangle([x0, y0, x0 + box_w, y0 + box_h], fill=(0, 0, 0, 150))
    for i, ln in enumerate(fitted):
        d.text((x0 + pad, y0 + pad + i * line_h), ln, font=font, fill=(255, 255, 255, 255))
    return Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")


def split_chunks(s):
    return [s[i:i + CHUNK] for i in range(0, len(s), CHUNK)]


def from_b64(s):
    return Image.open(io.BytesIO(base64.b64decode(s)))


def build_excel(rows):
    """rows: dicts with text fields and images {'VIN','HRS','ODO','FRONT'} as base64 (full photos)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Vehicles"
    ws.append(["Submitted At", "Vehicle Ref", "VIN (typed)", "Hours (typed)", "Odometer km (typed)",
               "VIN photo", "Hours photo", "Odometer photo", "Front photo"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for col, wdt in zip("ABCDEFGHI", (19, 16, 22, 12, 16, 44, 44, 44, 36)):
        ws.column_dimensions[col].width = wdt
    img_cols = {"VIN": "F", "HRS": "G", "ODO": "H", "FRONT": "I"}
    for r, row in enumerate(rows, start=2):
        ws.append([row.get("submitted", ""), row.get("ref", ""), row.get("vin", ""),
                   row.get("hours", ""), row.get("odo", "")])
        ws.row_dimensions[r].height = 165
        for key, col in img_cols.items():
            b64 = row.get("images", {}).get(key)
            if not b64:
                continue
            im = from_b64(b64)
            buf = io.BytesIO(base64.b64decode(b64))
            xl = XLImage(buf)  # original photo is embedded; only the display size is reduced
            scale = min(215 / im.height, 300 / im.width)
            xl.width, xl.height = int(im.width * scale), int(im.height * scale)
            ws.add_image(xl, f"{col}{r}")
    ws.freeze_panes = "A2"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
