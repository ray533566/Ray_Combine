"""
Excel Mapping Combiner — Streamlit app
Function 1: Auto-combine 'Mapping' sheets of many Chart xlsx files (header kept once, A..AZ).
Function 2: Append new rows below an existing (previously combined) Excel file.
Run locally:  streamlit run app.py
"""
import os
import zipfile
from datetime import datetime, timedelta, timezone
from io import BytesIO

import pandas as pd
import streamlit as st

import combine_core as cc

TW = timezone(timedelta(hours=8))   # Taipei time for the default file name


class Src:
    """Uniform input item: uploaded xlsx or an xlsx extracted from a zip."""
    def __init__(self, name, data):
        self.name, self.data, self.size = name, data, len(data)


def expand(uploads):
    """Uploaded files -> list[Src]; .zip files are unpacked (xlsx/xlsm inside, any folder)."""
    out = []
    for u in uploads:
        if u.name.lower().endswith(".zip"):
            with zipfile.ZipFile(BytesIO(u.getvalue())) as z:
                for info in z.infolist():
                    base = os.path.basename(info.filename)
                    if (info.is_dir() or "__MACOSX" in info.filename or base.startswith(("~$", "."))
                            or not base.lower().endswith((".xlsx", ".xlsm"))):
                        continue
                    out.append(Src(base, z.read(info)))
        else:
            out.append(Src(u.name, u.getvalue()))
    return sorted(out, key=lambda s: s.name)


st.set_page_config(page_title="Excel Mapping Combiner", page_icon="📊", layout="wide")
st.title("📊 Excel Mapping Combiner")
st.caption("Combine the 'Mapping' sheet of Chart xlsx files (80-400mA, 100-400mA …) · "
           "Row 1 kept once · Columns A..AZ · values only · .zip upload supported")

# ---------------------------------------------------------------- settings
with st.sidebar:
    st.header("Settings")
    sheet = st.text_input("Source sheet name", "Mapping")
    cols = st.text_input("Column range", "A:AZ")
    skip_dup = st.checkbox("Skip files whose SN + TESTNUMBER already exist", value=True,
                           help="Prevents adding the same unit/test twice.")
    out_name = st.text_input("Output file name",
                             f"Combine_{datetime.now(TW):%Y%m%d_%H%M}.xlsx")

try:
    cc.parse_cols(cols)
except Exception:
    st.error(f"Invalid column range: '{cols}' (use e.g. A:AZ)")
    st.stop()

# ---------------------------------------------------------------- functions
c1, c2 = st.columns(2)
f1 = c1.checkbox("① Auto-combine Excel files", value=True,
                 help="Upload many Chart files → one combined sheet.")
f2 = c2.checkbox("② Append to previous Excel (new rows below old rows)", value=False,
                 help="Upload the previous combined Excel, then new files; new data goes below the old rows.")

if not (f1 or f2):
    st.info("Select at least one function.")
    st.stop()

old_file = None
if f2:
    old_file = st.file_uploader("Previous Excel file (old data)", type=["xlsx", "xlsm"], key="old")

uploads = st.file_uploader(
    "New Excel file(s) or a .zip of them" + (" — multiple allowed" if f1 else " — one file"),
    type=["xlsx", "xlsm", "zip"], accept_multiple_files=f1, key=f"new_{f1}")
if uploads and not isinstance(uploads, list):
    uploads = [uploads]
try:
    new_files = expand(uploads or [])
except zipfile.BadZipFile:
    st.error("The .zip file could not be opened.")
    st.stop()
if uploads and not new_files:
    st.warning("No .xlsx/.xlsm files found in the upload.")
elif new_files:
    st.caption(f"{len(new_files)} Excel file(s) ready")

ready = bool(new_files) and (not f2 or old_file is not None)
if not ready:
    st.info("Upload " + ("the previous Excel file and " if f2 and old_file is None else "") + "new file(s) to continue.")
    st.stop()

def data_check(header, rows):
    """Summary: units, repeated SN, rows with empty WaveLength per SN/TESTNUMBER."""
    names = [str(h).strip().upper() if h is not None else "" for h in header]
    def idx(*cands):
        for c in cands:
            for i, n in enumerate(names):
                if n.startswith(c):
                    return i
        return None
    i_sn, i_tn, i_wl, i_t = idx("SN"), idx("TESTNUMBER"), idx("WAVELENGTH"), idx("TEMP")
    if i_sn is None or i_tn is None:
        return None
    tests = {}
    for r in rows:
        tests.setdefault(r[i_sn], set()).add(r[i_tn])
    repeated = {sn: len(t) for sn, t in tests.items() if len(t) > 1}
    blanks = {}
    if i_wl is not None:
        for r in rows:
            if r[i_wl] in (None, ""):
                b = blanks.setdefault((r[i_sn], r[i_tn]), {"rows": 0, "temps": set()})
                b["rows"] += 1
                if i_t is not None:
                    b["temps"].add(r[i_t])
    return {
        "units": len(tests), "tests": sum(len(t) for t in tests.values()),
        "repeated": pd.DataFrame([{"SN": k, "Tests": v} for k, v in sorted(repeated.items())]),
        "blanks": pd.DataFrame([{"SN": k[0], "TESTNUMBER": k[1], "Empty WL rows": v["rows"],
                                 "Temps (°C)": ", ".join(str(t) for t in sorted(v["temps"], key=lambda x: (x is None, x)))}
                                for k, v in sorted(blanks.items(), key=lambda kv: -kv[1]["rows"])]),
    }


# signature of current inputs — hides an old result when inputs change
sig = (f1, f2, sheet, cols, skip_dup,
       (old_file.name, old_file.size) if old_file else None,
       tuple((f.name, f.size) for f in new_files))

# ---------------------------------------------------------------- run
if st.button("▶ Run", type="primary"):
    log, header, rows, old_count = [], None, [], 0

    if f2:
        try:
            header, rows, name = cc.read_table(old_file.getvalue(), cols)
            old_count = len(rows)
            log.append({"File": old_file.name, "Role": "previous", "Rows": old_count,
                        "Status": "OK", "Note": f"sheet '{name}'"})
        except Exception as e:
            st.error(f"Cannot read previous file: {e}")
            st.stop()

    kc = cc.key_columns(header) if header else None
    seen = cc.keys_of(rows, kc)

    bar = st.progress(0.0, text="Reading files…")
    with st.container():
        for idx, f in enumerate(new_files, 1):
            bar.progress(idx / len(new_files), text=f"Reading {idx}/{len(new_files)}: {f.name}")
            try:
                h, body, note = cc.read_any(f.data, sheet, cols)
            except Exception as e:
                log.append({"File": f.name, "Role": "new", "Rows": 0, "Status": "SKIP", "Note": str(e)})
                continue
            if header is None:
                header = h
                kc = cc.key_columns(header)
            elif [str(x) for x in h if x is not None] != [str(x) for x in header if x is not None]:
                log.append({"File": f.name, "Role": "new", "Rows": 0, "Status": "HEADER MISMATCH",
                            "Note": "columns differ from base file (e.g. 80-400mA vs 100-400mA) — skipped"})
                continue
            k = cc.keys_of(body, kc)
            if skip_dup and k and k <= seen:
                log.append({"File": f.name, "Role": "new", "Rows": 0, "Status": "DUPLICATE",
                            "Note": "SN+TESTNUMBER already present — skipped"})
                continue
            seen |= k
            rows.extend(body)
            log.append({"File": f.name, "Role": "new", "Rows": len(body), "Status": "OK", "Note": note})

    bar.progress(1.0, text="Writing Excel…")
    if header is None or not rows:
        bar.empty()
        st.error("No data combined.")
        st.dataframe(pd.DataFrame(log), use_container_width=True, hide_index=True)
        st.stop()

    w = cc.used_width(header, rows)
    names = [str(header[i]) if i < len(header) and header[i] is not None else f"col{i+1}" for i in range(w)]
    preview = pd.DataFrame([(list(r[:w]) + [None] * (w - len(r[:w]))) for r in rows[:old_count + 500]],
                           columns=names)
    xlsx = cc.to_xlsx(header, rows)
    bar.empty()
    st.session_state["result"] = {
        "sig": sig, "xlsx": xlsx, "check": data_check(header, rows),
        "log": log, "old": old_count, "total": len(rows), "width": w, "preview": preview,
        "name": out_name if out_name.lower().endswith(".xlsx") else out_name + ".xlsx",
    }

# ---------------------------------------------------------------- result
res = st.session_state.get("result")
if res and res["sig"] == sig:
    from openpyxl.utils import get_column_letter
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Old rows", res["old"])
    m2.metric("New rows", res["total"] - res["old"])
    m3.metric("Total rows", res["total"])
    m4.metric("Columns", f"A..{get_column_letter(res['width'])}")

    st.subheader("Files")
    st.dataframe(pd.DataFrame(res["log"]), use_container_width=True, hide_index=True)

    st.download_button("⬇ Download combined Excel", res["xlsx"], file_name=res["name"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")

    chk = res.get("check")
    if chk:
        st.subheader("Data check")
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Units (SN)", chk["units"])
        d2.metric("Tests (SN+TESTNUMBER)", chk["tests"])
        d3.metric("SN tested >1×", len(chk["repeated"]))
        d4.metric("Tests with empty WL", len(chk["blanks"]))
        if len(chk["repeated"]):
            with st.expander(f"SN tested more than once ({len(chk['repeated'])})"):
                st.dataframe(chk["repeated"], use_container_width=True, hide_index=True)
        if len(chk["blanks"]):
            with st.expander(f"Tests with empty WaveLength ({len(chk['blanks'])}) — possibly incomplete sweep"):
                st.dataframe(chk["blanks"], use_container_width=True, hide_index=True)

    st.subheader("Preview")
    df = res["preview"]
    if res["old"]:
        t1, t2 = st.tabs(["New rows (appended)", "All rows"])
        t1.dataframe(df.iloc[res["old"]:].head(500), use_container_width=True)
        t2.dataframe(df.head(500), use_container_width=True)
    else:
        st.dataframe(df.head(500), use_container_width=True)
    st.caption("Preview shows up to 500 rows; the download contains all rows.")
