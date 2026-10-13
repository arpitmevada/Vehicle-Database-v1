import base64
import io

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font
from PIL import Image, ImageOps

# Default crop areas (percent of the photo). Adjust with sliders in the app.
# (x_from, x_to), (y_from, y_to), rotate_degrees
DEFAULTS = {
    "VIN": ((24, 35), (28, 81), 270),
    "ODO": ((55, 78), (53, 64), 0),
    "HRS": ((44, 72), (57, 72), 0),
}


def load_image(file):
    img = Image.open(file)
    return ImageOps.exif_transpose(img).convert("RGB")


def crop_percent(img, xr, yr, rotate=0):
    w, h = img.size
    c = img.crop((int(w * xr[0] / 100), int(h * yr[0] / 100),
                  int(w * xr[1] / 100), int(h * yr[1] / 100)))
    if rotate:
        c = c.rotate(rotate, expand=True)
    return c


def to_b64(img, max_w=600, limit=45000):
    """JPEG -> base64 text that fits in one Google Sheets cell (50,000 chars max)."""
    img = img.copy()
    q = 80
    while True:
        im = img.copy()
        im.thumbnail((max_w, max_w * 3))
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=q, optimize=True)
        s = base64.b64encode(buf.getvalue()).decode()
        if len(s) <= limit:
            return s
        if q > 40:
            q -= 10
        else:
            max_w = int(max_w * 0.8)


def from_b64(s):
    return Image.open(io.BytesIO(base64.b64decode(s)))


def build_excel(rows):
    """rows: list of dicts with text fields + image dict {'VIN','HRS','ODO','FRONT'} of base64."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Vehicles"
    heads = ["Submitted At", "Vehicle Ref", "VIN (typed)", "Hours (typed)", "Odometer km (typed)",
             "VIN photo", "Hours photo", "Odometer photo", "Front photo"]
    ws.append(heads)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for col, wdt in zip("ABCDEFGHI", (19, 16, 22, 12, 16, 42, 34, 34, 34)):
        ws.column_dimensions[col].width = wdt
    img_cols = {"VIN": "F", "HRS": "G", "ODO": "H", "FRONT": "I"}
    for r, row in enumerate(rows, start=2):
        ws.append([row.get("submitted", ""), row.get("ref", ""), row.get("vin", ""),
                   row.get("hours", ""), row.get("odo", "")])
        ws.row_dimensions[r].height = 62
        for key, col in img_cols.items():
            b64 = row.get("images", {}).get(key)
            if not b64:
                continue
            im = from_b64(b64)
            buf = io.BytesIO()
            im.save(buf, "PNG")
            buf.seek(0)
            xl = XLImage(buf)
            scale = min(78 / im.height, 280 / im.width)
            xl.width, xl.height = int(im.width * scale), int(im.height * scale)
            ws.add_image(xl, f"{col}{r}")
    ws.freeze_panes = "A2"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
