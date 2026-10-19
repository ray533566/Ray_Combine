# 📊 Excel Mapping Combiner

Streamlit web app that combines the **`Mapping`** sheet of Chart test files (`*-Chart-80-400mA.xlsx`, `*-Chart-100-400mA.xlsx`, …) into one Excel sheet.
Upload the xlsx files directly **or one .zip** containing them (any folder inside the zip).

## Functions (checkboxes)

| | Function | What it does |
|---|---|---|
| ① | **Auto-combine Excel files** | Upload many Chart files → one sheet. Row 1 (header) kept once, columns **A..AZ**, values only. |
| ② | **Append to previous Excel** | Upload the previous combined Excel + new file(s) → new rows are added **below the old rows**. |

Tick both to combine several new files and append them to the old file in one step.

**Details**
- Values come from Excel's cached formula results. If a file was never saved in Excel (no cache), the Mapping values are recalculated from the `RawData` sheet.
- *Skip duplicates* (sidebar, on by default): a new file whose SN + TESTNUMBER already exists is skipped.
- Works with any Mapping template (80-400mA, 100-400mA, extra columns such as Case_Temp / DDMI_Temp): the fallback reads the Mapping formulas, nothing is hard-coded.
- Header cells like `=RawData!A2&"-Case_Temp"` are written as `Case_Temp` (no SN prefix).
- Files whose columns differ from the base file (e.g. 80-400mA mixed with 100-400mA) are skipped as HEADER MISMATCH — combine each template separately.
- New files can be Chart files (`Mapping` sheet) or already-combined files (`Combine` / first sheet).
- Sidebar: source sheet name, column range, output file name.
- **Data check** after each run: number of units / tests, SN tested more than once, tests with empty WaveLength (e.g. incomplete temperature sweep).
- Upload limit raised to 500 MB in `.streamlit/config.toml`.

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
.streamlit/config.toml   upload size limit
```
