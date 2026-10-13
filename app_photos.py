import base64
import io
import json
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image

from core import build_excel, load_image, prep_photo, split_chunks

st.set_page_config(page_title="Vehicle Photo Capture", page_icon="🚚", layout="centered")

TEXT_HEADS = ["ID", "Submitted At", "Vehicle Ref", "VIN (typed)", "Hours (typed)", "Odometer km (typed)"]
IMG_HEADS = ["ID", "KEY", "SEQ", "DATA"]

# Shrinks each photo on the phone BEFORE upload (5 MB -> ~100-250 KB)
_picker = components.declare_component(
    "photo_picker", path=str(Path(__file__).parent / "photo_picker"))


def get_photo(label, key, fallback):
    if fallback:
        f = st.file_uploader(label, type=["jpg", "jpeg", "png"], key=key)
        return load_image(f) if f else None
    v = _picker(label=label, max_side=1000, quality=0.75, key=key, default=None)
    if v and v.get("b64"):
        return Image.open(io.BytesIO(base64.b64decode(v["b64"]))).convert("RGB")
    return None


def check_password():
    pw = st.secrets.get("APP_PASSWORD")
    if not pw or st.session_state.get("auth"):
        return True
    entered = st.text_input("Password", type="password")
    if entered == pw:
        st.session_state["auth"] = True
        st.rerun()
    elif entered:
        st.error("Wrong password")
    return False


@st.cache_resource
def get_sheets():
    gc = gspread.service_account_from_dict(json.loads(st.secrets["GCP_JSON"]))
    sh = gc.open_by_key(st.secrets["SHEET_ID"])
    data = sh.sheet1
    if data.row_values(1) != TEXT_HEADS:
        data.clear()
        data.append_row(TEXT_HEADS)
    try:
        imgs = sh.worksheet("Images")
    except gspread.WorksheetNotFound:
        imgs = sh.add_worksheet("Images", rows=1000, cols=len(IMG_HEADS))
    if imgs.row_values(1) != IMG_HEADS:
        imgs.clear()
        imgs.append_row(IMG_HEADS)
    return data, imgs


@st.cache_data(ttl=30, show_spinner="Loading records...")
def load_rows():
    data, imgs = get_sheets()
    parts = {}
    for r in imgs.get_all_values()[1:]:
        if len(r) >= 4:
            parts.setdefault((r[0], r[1]), []).append((int(r[2]), r[3]))
    rows = []
    for r in data.get_all_values()[1:]:
        r = r + [""] * (6 - len(r))
        images = {k: "".join(c for _, c in sorted(parts.get((r[0], k), [])))
                  for k in ("VIN", "HRS", "ODO", "FRONT")}
        rows.append({"submitted": r[1], "ref": r[2], "vin": r[3], "hours": r[4], "odo": r[5],
                     "images": images})
    return rows


st.title("🚚 Vehicle Photo Capture")
if not check_password():
    st.stop()

tab_add, tab_db = st.tabs(["Add new set", "Database"])

with tab_add:
    st.caption("Upload the 4 photos and tap Save. Photos go into Excel as they are (no cropping, no AI).")
    n = st.session_state.setdefault("nonce", 0)
    fallback = st.checkbox("Photo picker not working? Use standard uploader (no phone-side resize)")
    up = {
        "VIN": get_photo("1. VIN plate photo", f"vin{n}", fallback),
        "HRS": get_photo("2. Hours (HRS) photo", f"hrs{n}", fallback),
        "ODO": get_photo("3. Odometer (km) photo", f"odo{n}", fallback),
        "FRONT": get_photo("4. Front of truck photo", f"front{n}", fallback),
    }
    rot = st.selectbox("Rotate VIN photo (use 270 if the plate appears sideways)", [0, 90, 180, 270],
                       index=3, key=f"rot{n}")

    with st.expander("Optional: type values and a vehicle reference (makes Excel searchable)"):
        ref = st.text_input("Vehicle ref / plate", key=f"ref{n}")
        t_vin = st.text_input("VIN", key=f"tv{n}")
        t_hrs = st.text_input("Hours", key=f"th{n}")
        t_odo = st.text_input("Odometer km", key=f"to{n}")

    if st.button("Save to database", type="primary", disabled=any(v is None for v in up.values())):
        with st.spinner("Saving photos..."):
            now = datetime.now(ZoneInfo("Asia/Kolkata"))
            rid = now.strftime("%Y%m%d%H%M%S%f")[:17]
            data, imgs = get_sheets()
            rows = []
            for key, f in up.items():
                b64 = prep_photo(f, rotate=rot if key == "VIN" else 0)
                for seq, chunk in enumerate(split_chunks(b64)):
                    rows.append([rid, key, seq, chunk])
            imgs.append_rows(rows, value_input_option="RAW")
            data.append_row([rid, now.strftime("%Y-%m-%d %H:%M:%S"), ref, t_vin.upper(), t_hrs, t_odo],
                            value_input_option="RAW")
            load_rows.clear()
        st.success("Saved. You can add the next set.")
        st.session_state["nonce"] += 1
        st.rerun()

with tab_db:
    if st.button("Refresh"):
        load_rows.clear()
        st.rerun()
    rows = load_rows()
    st.write(f"{len(rows)} records")
    if rows:
        st.dataframe(pd.DataFrame([{k: r[k] for k in ("submitted", "ref", "vin", "hours", "odo")}
                                   for r in rows]), use_container_width=True)
        st.download_button(
            "Download Excel (photos in cells)", build_excel(rows), "vehicle_database.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
