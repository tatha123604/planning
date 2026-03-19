# CLI Matrix Overdue Updater

Lightweight Streamlit app to refresh the first two sheets of your CLI Matrix workbook using the latest CMS export.

## What it does
- Reads the new matrix file (e.g., `CLI Matrix-19-03-2026.xlsx`).
- Aggregates overdue counts by `CLI ID` + `CLI Name` + `Alloted Desig.`.
- Rebuilds:
  - Sheet 1 (`Summary position of FP OVERDUE`): only FP overdue rows.
  - Sheet 2 (`18.03.26` or whatever is the second sheet name in the template): all overdue metrics.
- Copies the remaining sheets from the template unchanged and saves a fresh workbook.

## Run locally
```bash
cd C:\Users\HP\Documents\Playground
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## How to use
1) Open the Streamlit page in your browser (auto-opens after the command).
2) Upload:
   - **Latest CLI Matrix (source)** — the fresh CMS export.
   - **Template / previous workbook** — the file whose other sheets you want to keep.
3) Click **Generate output file**.
4) Download `CLI_Matrix_updated.xlsx`.

## Notes
- Column detection is driven by the header row at index 2 of the source file (after the two title rows).
- Date parsing assumes `dayfirst=True` (e.g., `19-03-2026`).
- Total overdue is recalculated as `FP + Counsel + Grading` per row before aggregation.
