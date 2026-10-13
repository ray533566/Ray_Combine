"""
Excel Mapping Combiner — Streamlit app
Function 1: Auto-combine 'Mapping' sheets of many Chart xlsx files (header kept once, A..AZ).
Function 2: Append new rows below an existing (previously combined) Excel file.
Run locally:  streamlit run app.py
"""
from datetime import datetime

import pandas as pd
import streamlit as st

import combine_core as cc

st.set_page_config(page_title="Excel Mapping Combiner", page_icon="📊", layout="wide")
st.title("📊 Excel Mapping Combiner")
st.caption("Combine the 'Mapping' sheet of *-Chart-80-400mA.xlsx files · Row 1 kept once · Columns A..AZ · values only")

# ---------------------------------------------------------------- settings
with st.sidebar:
    st.header("Settings")
    sheet = st.text_input("Source sheet name", "Mapping")
    cols = st.text_input("Column range", "A:AZ")
    skip_dup = st.checkbox("Skip files whose SN + TESTNUMBER already exist", value=True,
                           help="Prevents adding the same unit/test twice.")
    out_name = st.text_input("Output file name",
                             f"Combine80-400mA_{datetime.now():%Y%m%d_%H%M}.xlsx")

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

new_files = st.file_uploader(
    "New Excel file(s)" + (" — multiple allowed" if f1 else " — one file"),
    type=["xlsx", "xlsm"], accept_multiple_files=f1, key=f"new_{f1}")
if new_files and not isinstance(new_files, list):
    new_files = [new_files]
new_files = sorted(new_files or [], key=lambda f: f.name)

ready = bool(new_files) and (not f2 or old_file is not None)
if not ready:
    st.info("Upload " + ("the previous Excel file and " if f2 and old_file is None else "") + "new file(s) to continue.")
    st.stop()

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

    with st.spinner("Processing..."):
        for f in new_files:
            try:
                h, body, note = cc.read_any(f.getvalue(), sheet, cols)
            except Exception as e:
                log.append({"File": f.name, "Role": "new", "Rows": 0, "Status": "SKIP", "Note": str(e)})
                continue
            if header is None:
                header = h
                kc = cc.key_columns(header)
            elif [str(x) for x in h if x is not None] != [str(x) for x in header if x is not None]:
                note = (note + "; " if note else "") + "header differs — appended by column position"
            k = cc.keys_of(body, kc)
            if skip_dup and k and k <= seen:
                log.append({"File": f.name, "Role": "new", "Rows": 0, "Status": "DUPLICATE",
                            "Note": "SN+TESTNUMBER already present — skipped"})
                continue
            seen |= k
            rows.extend(body)
            log.append({"File": f.name, "Role": "new", "Rows": len(body), "Status": "OK", "Note": note})

    if header is None or not rows:
        st.error("No data combined.")
        st.dataframe(pd.DataFrame(log), use_container_width=True, hide_index=True)
        st.stop()

    w = cc.used_width(header, rows)
    names = [str(header[i]) if i < len(header) and header[i] is not None else f"col{i+1}" for i in range(w)]
    preview = pd.DataFrame([(list(r[:w]) + [None] * (w - len(r[:w]))) for r in rows[:old_count + 500]],
                           columns=names)
    st.session_state["result"] = {
        "sig": sig, "xlsx": cc.to_xlsx(header, rows),
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

    st.subheader("Preview")
    df = res["preview"]
    if res["old"]:
        t1, t2 = st.tabs(["New rows (appended)", "All rows"])
        t1.dataframe(df.iloc[res["old"]:].head(500), use_container_width=True)
        t2.dataframe(df.head(500), use_container_width=True)
    else:
        st.dataframe(df.head(500), use_container_width=True)
    st.caption("Preview shows up to 500 rows; the download contains all rows.")
