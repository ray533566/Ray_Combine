"""
Excel Sheet Combiner — Streamlit app
Function 1: Auto-combine any sheet(s) (picked by name) of many Excel files (header kept once).
Function 2: Append new rows below an existing (previously combined) Excel file.
Run locally:  streamlit run app.py
"""
import os
import zipfile
from collections import Counter
from datetime import datetime, timedelta, timezone
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

import combine_core as cc

TW = timezone(timedelta(hours=8))   # Taipei time for the default file name
SRC_COL = "Source file"


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


def data_check(header, rows):
    """Summary: units, repeated SN, rows with empty WaveLength per SN/TESTNUMBER (Mapping-type sheets)."""
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
        if len(r) > max(i_sn, i_tn):
            tests.setdefault(r[i_sn], set()).add(r[i_tn])
    repeated = {sn: len(t) for sn, t in tests.items() if len(t) > 1}
    blanks = {}
    if i_wl is not None:
        for r in rows:
            if len(r) > i_wl and r[i_wl] in (None, ""):
                b = blanks.setdefault((r[i_sn], r[i_tn]), {"rows": 0, "temps": set()})
                b["rows"] += 1
                if i_t is not None:
                    b["temps"].add(r[i_t])
    return {
        "units": len(tests), "tests": sum(len(t) for t in tests.values()),
        "repeated": pd.DataFrame([{"SN": k, "Tests": v} for k, v in sorted(repeated.items(), key=lambda kv: str(kv[0]))]),
        "blanks": pd.DataFrame([{"SN": k[0], "TESTNUMBER": k[1], "Empty WL rows": v["rows"],
                                 "Temps (°C)": ", ".join(str(t) for t in sorted(v["temps"], key=lambda x: (x is None, str(x))))}
                                for k, v in sorted(blanks.items(), key=lambda kv: -kv[1]["rows"])]),
    }


def split_src(header, rows):
    """Remove an existing 'Source file' column -> (header, rows, sources)."""
    if SRC_COL not in header:
        return header, rows, [None] * len(rows)
    i = header.index(SRC_COL)
    src = [r[i] if i < len(r) else None for r in rows]
    return header[:i] + header[i + 1:], [r[:i] + r[i + 1:] for r in rows], src


# ================================================================ UI
st.set_page_config(page_title="Excel Sheet Combiner", page_icon="📊", layout="wide")
st.title("📊 Excel Sheet Combiner")
st.caption("Combine any sheet (pick by name) of many Excel files into one · Row 1 kept once · "
           "values only · .zip upload supported")

# ---------------------------------------------------------------- settings
with st.sidebar:
    st.header("Settings")
    cols = st.text_input("Column range", "A:AZ", help="e.g. A:AZ. Leave empty (or ALL) for all used columns.")
    mismatch = st.radio(
        "If a file's header differs", ["Match by header name", "Keep column position", "Skip file"],
        help="Match by header name: same-position columns stay, moved columns are found by name, "
             "new names are added at the end.\nKeep column position: append column-by-column (A→A).\n"
             "Skip file: file is listed as HEADER MISMATCH.")
    key_col = st.text_input("Skip rows where this column is empty", "A",
                            help="Removes template/filler rows (e.g. rows that only contain helper formulas). "
                                 "Column letter of the source sheet. Leave empty to keep all rows. "
                                 "Not applied to a file where this column is empty in every row.")
    skip_dup = st.checkbox("Skip files whose SN + TESTNUMBER already exist", value=True,
                           help="Only for sheets that have SN and TESTNUMBER columns (e.g. Mapping).")
    dedup_rows = st.checkbox("Remove duplicate rows", value=False,
                             help="Drop rows that are exactly identical to a row already combined "
                                  "(e.g. the same test appearing in several daily files).")
    add_src = st.checkbox(f"Add '{SRC_COL}' column", value=False,
                          help="Last column = name of the file each row came from.")
    out_name = st.text_input("Output file name", f"Combine_{datetime.now(TW):%Y%m%d_%H%M}.xlsx")

try:
    cc.parse_cols(cols)
except Exception:
    st.error(f"Invalid column range: '{cols}' (use e.g. A:AZ, or leave empty for all columns)")
    st.stop()
key_idx = None
if key_col.strip():
    try:
        key_idx = cc.column_index_from_string(key_col.strip().upper()) - cc.parse_cols(cols)[0]
        assert key_idx >= 0
    except Exception:
        st.error(f"Invalid column '{key_col}' for 'Skip rows where this column is empty' (use e.g. A)")
        st.stop()

# ---------------------------------------------------------------- functions
c1, c2 = st.columns(2)
f1 = c1.checkbox("① Auto-combine Excel files", value=True,
                 help="Upload many Excel files → selected sheet(s) combined.")
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

if not new_files or (f2 and old_file is None):
    st.info("Upload " + ("the previous Excel file and " if f2 and old_file is None else "") + "new file(s) to continue.")
    st.stop()

# ---------------------------------------------------------------- sheet picker
sheets_of = {}
for f in new_files:
    try:
        sheets_of[f.name] = cc.list_sheets(f.data)
    except Exception as e:
        sheets_of[f.name] = []
        st.warning(f"{f.name}: cannot open ({e})")
count = Counter(s for v in sheets_of.values() for s in v)
order = sorted(count, key=lambda s: (-count[s], [n for v in sheets_of.values() for n in v].index(s)))
n_files = len(new_files)
default = ["Mapping"] if "Mapping" in count else order[:1]
picked = st.multiselect(
    f"Sheet(s) to combine — {n_files} file(s) uploaded", order, default=default,
    format_func=lambda s: f"{s}   ({count[s]}/{n_files} files)",
    help="Each selected sheet becomes one sheet in the output file.")
if not picked:
    st.info("Select at least one sheet.")
    st.stop()

old_sheets = cc.list_sheets(old_file.getvalue()) if old_file else []


def old_match(sheet):
    """Which sheet of the previous file continues `sheet`."""
    if sheet in old_sheets:
        return sheet
    if len(picked) == 1 and old_sheets:          # e.g. old 'Combine' sheet <- new 'Mapping'
        return "Combine" if "Combine" in old_sheets else old_sheets[0]
    return None


# signature of current inputs — hides an old result when inputs change
sig = (f1, f2, tuple(picked), cols, mismatch, skip_dup, dedup_rows, add_src,
       (old_file.name, old_file.size) if old_file else None,
       tuple((f.name, f.size) for f in new_files))


# ---------------------------------------------------------------- run
def combine_sheet(sheet, tick):
    log, header, rows, src, old_count = [], None, [], [], 0
    out_sheet = sheet
    om = old_match(sheet) if old_file else None
    if old_file and om:
        h, b, _ = cc.read_table(old_file.getvalue(), cols, om)
        header, rows, src = split_src(list(h), [list(r) for r in b])
        old_count, out_sheet = len(rows), om
        log.append({"File": old_file.name, "Role": "previous", "Rows": old_count,
                    "Status": "OK", "Note": f"sheet '{om}'"})
    elif old_file:
        log.append({"File": old_file.name, "Role": "previous", "Rows": 0, "Status": "—",
                    "Note": f"no matching sheet for '{sheet}' — starting new"})

    kc = cc.key_columns(header) if header else None
    seen_keys = cc.keys_of(rows, kc)
    seen_rows = {tuple(r) for r in rows} if dedup_rows else set()

    for f in new_files:
        tick(f"{sheet}: {f.name}")
        if sheet not in sheets_of.get(f.name, []):
            log.append({"File": f.name, "Role": "new", "Rows": 0, "Status": "NO SHEET",
                        "Note": f"sheet '{sheet}' not in this file"})
            continue
        try:
            h, body, note = cc.read_sheet(f.data, sheet, cols)
        except Exception as e:
            log.append({"File": f.name, "Role": "new", "Rows": 0, "Status": "SKIP", "Note": str(e)})
            continue
        h, body, _ = split_src(list(h), [list(r) for r in body])
        notes = [note] if note else []
        if key_idx is not None:
            def filled(r):
                return key_idx < len(r) and r[key_idx] not in (None, "")
            if any(filled(r) for r in body):
                n0 = len(body)
                body = [r for r in body if filled(r)]
                if n0 - len(body):
                    notes.append(f"{n0 - len(body)} rows with empty column {key_col.strip().upper()} skipped")
            elif body:
                notes.append(f"column {key_col.strip().upper()} empty in every row — row filter not applied")
        if header is None:
            header = h
            kc = cc.key_columns(header)
        elif not cc.same_header(h, header):
            if mismatch == "Skip file":
                log.append({"File": f.name, "Role": "new", "Rows": 0, "Status": "HEADER MISMATCH",
                            "Note": "columns differ from base file — skipped"})
                continue
            if mismatch == "Match by header name":
                header, body, added = cc.align_to(header, h, body)
                notes.append("header differs → matched by name" + (f"; new columns: {', '.join(added)}" if added else ""))
            else:
                notes.append("header differs → kept column position")
        k = cc.keys_of(body, kc)
        if skip_dup and k and k <= seen_keys:
            log.append({"File": f.name, "Role": "new", "Rows": 0, "Status": "DUPLICATE",
                        "Note": "SN+TESTNUMBER already present — skipped"})
            continue
        seen_keys |= k
        if dedup_rows:
            body, n_rm = cc.dedupe(body, seen_rows)
            if n_rm:
                notes.append(f"{n_rm} duplicate rows removed")
        rows.extend(body)
        src.extend([f.name] * len(body))
        log.append({"File": f.name, "Role": "new", "Rows": len(body), "Status": "OK", "Note": "; ".join(notes)})

    if header is None or not rows:
        return {"sheet": out_sheet, "src_sheet": sheet, "log": log, "empty": True}

    w = cc.used_width(header, rows)
    header = list(header[:w]) + [None] * (w - len(header[:w]))
    rows = [list(r[:w]) + [None] * (w - len(r[:w])) for r in rows]
    if add_src:
        header = header + [SRC_COL]
        rows = [r + [s] for r, s in zip(rows, src)]
        w += 1
    names = [str(header[i]) if header[i] not in (None, "") else f"col{i+1}" for i in range(w)]
    names = [n if names.index(n) == i else f"{n}_{i+1}" for i, n in enumerate(names)]   # unique for pandas
    return {
        "sheet": out_sheet, "src_sheet": sheet, "log": log, "empty": False,
        "header": header, "rows": rows, "old": old_count, "total": len(rows), "width": w,
        "check": data_check(header, rows),
        "preview_new": pd.DataFrame(rows[old_count:old_count + 500], columns=names),
        "preview_old": pd.DataFrame(rows[:200], columns=names) if old_count else None,
    }


if st.button("▶ Run", type="primary"):
    steps, done = len(picked) * len(new_files), [0]
    bar = st.progress(0.0, text="Reading files…")

    def tick(msg):
        done[0] += 1
        bar.progress(min(done[0] / steps, 1.0), text=f"Reading {done[0]}/{steps} — {msg}")

    results = [combine_sheet(s, tick) for s in picked]
    good = [r for r in results if not r["empty"]]
    xlsx = None
    if good:
        bar.progress(1.0, text="Writing Excel…")
        xlsx = cc.to_xlsx_multi([(r["sheet"], r["header"], r["rows"]) for r in good])
        for r in good:                         # free memory: keep only what the page shows
            del r["rows"], r["header"]
    bar.empty()
    st.session_state["result"] = {
        "sig": sig, "xlsx": xlsx, "results": results,
        "name": out_name if out_name.lower().endswith(".xlsx") else out_name + ".xlsx",
    }

# ---------------------------------------------------------------- result
res = st.session_state.get("result")
if res and res["sig"] == sig:
    if res["xlsx"] is None:
        st.error("No data combined.")
    else:
        st.download_button("⬇ Download combined Excel", res["xlsx"], file_name=res["name"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")
    tabs = st.tabs([f"📄 {r['sheet']}" for r in res["results"]])
    for tab, r in zip(tabs, res["results"]):
        with tab:
            if r["empty"]:
                st.warning(f"No data for sheet '{r['src_sheet']}'.")
                st.dataframe(pd.DataFrame(r["log"]), use_container_width=True, hide_index=True)
                continue
            if r["sheet"] != r["src_sheet"]:
                st.caption(f"Source sheet '{r['src_sheet']}' → appended to sheet '{r['sheet']}' of the previous file")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Old rows", f"{r['old']:,}")
            m2.metric("New rows", f"{r['total'] - r['old']:,}")
            m3.metric("Total rows", f"{r['total']:,}")
            m4.metric("Columns", f"A..{get_column_letter(r['width'])}")

            with st.expander("Files", expanded=len(r["log"]) <= 15):
                st.dataframe(pd.DataFrame(r["log"]), use_container_width=True, hide_index=True)

            chk = r.get("check")
            if chk:
                st.markdown("**Data check**")
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

            st.markdown("**Preview** (first 500 new rows; the download contains all rows)")
            st.dataframe(r["preview_new"], use_container_width=True)
            if r["preview_old"] is not None:
                with st.expander("First rows of the previous file"):
                    st.dataframe(r["preview_old"], use_container_width=True)
