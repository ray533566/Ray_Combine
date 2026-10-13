# 📊 Excel Mapping Combiner

Streamlit web app that combines the **`Mapping`** sheet of `*-Chart-80-400mA.xlsx` test files into one Excel sheet.

## Functions (checkboxes)

| | Function | What it does |
|---|---|---|
| ① | **Auto-combine Excel files** | Upload many Chart files → one sheet. Row 1 (header) kept once, columns **A..AZ**, values only. |
| ② | **Append to previous Excel** | Upload the previous combined Excel + new file(s) → new rows are added **below the old rows**. |

Tick both to combine several new files and append them to the old file in one step.

**Details**
- Values come from Excel's cached formula results. If a file was never saved in Excel (no cache), the Mapping values are recalculated from the `RawData` sheet.
- *Skip duplicates* (sidebar, on by default): a new file whose SN + TESTNUMBER already exists is skipped.
- New files can be Chart files (`Mapping` sheet) or already-combined files (`Combine` / first sheet).
- Sidebar: source sheet name, column range, output file name.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Push this folder to a GitHub repo (`app.py`, `combine_core.py`, `requirements.txt` in the repo root).
2. Go to <https://share.streamlit.io> → **Create app** → pick the repo, branch `main`, main file `app.py` → **Deploy**.

## Files
```
app.py            Streamlit UI
combine_core.py   combine / append logic (no UI, reusable)
requirements.txt
```
