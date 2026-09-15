# HR Planning Prototype

Lightweight script to plan headcount, recruiting, and promotions for roles such as Motorman, LPP, LPM, ALP, and Shunter.

## Files
- `config/requirements.json` - target headcount per role.
- `data/employees.csv` - roster with retirement and promotion readiness dates.
- `src/hr.py` - planner CLI (console).
- `src/app.py` - FastAPI web app.
- `src/models.py`, `src/logic.py`, `src/db.py`, `src/seed.py` - shared data model, planning logic, DB setup, and seed loaders.
- `templates/index.html`, `static/style.css` - browser UI assets.

## Console planner
```bash
python src/hr.py --as-of 2026-03-17 --horizon-months 12 --lead-time-days 90 \
  --requirements config/requirements.json --employees data/employees.csv
```

## Web app (browser + local SQLite)
1) Install deps (once): `python -m pip install -r requirements.txt`
2) Run server: `python -m uvicorn src.app:app --host 127.0.0.1 --port 8000 --reload`
3) Open `http://localhost:8000`

Features: shows headcount vs requirement, recruiting steps (immediate and timed backfills), promotion schedule, upcoming retirements, roster table, forms to add employees/change requirements, and Excel uploads (employees + seniority). Planning window and lead time are query params on the home page. Roles are sorted by hierarchy: LPM > Motorman > LPP > LPG > LPS (Shunter) > ALP.

Authentication
- Login required. Default credentials: user `admin`, password `sdah1234`.

### Run on Google Colab
Colab can host the app and expose it publicly with ngrok.

1) Clone the online branch and enter it:
```bash
!git clone -b codex/online https://gitlab.com/tatha1234/promotion.git
%cd promotion
```
2) Install Colab-specific deps:
```bash
!pip install -r requirements-colab.txt
```
3) Set your ngrok token (from https://dashboard.ngrok.com/get-started/your-authtoken):
```python
import os
os.environ["NGROK_AUTHTOKEN"] = "<your-token>"
```
4) Start the server with a tunnel:
```bash
!python colab_run.py --port 8000
```
The cell prints a public URL like `https://xxxx.ngrok.io`; open it to use the app. (Region is optional; leave it off for ngrok v3 defaults.)

### Excel upload format
- Employees: headers `name`, `role`, `hire_date`, `retirement_date`; optional `promotion_role`, `promotion_ready_date`.
- Seniority: headers `name`, `role`, `seniority_rank` (or `seniority`); optional `promotion_role`, `promotion_ready_date`. Rows update matching employees.
- Dates can be ISO strings (YYYY-MM-DD) or Excel date cells. Role aliases like `Shunter`/`LPS(Shunter)` normalize to `LPS`.
- Promotion ordering: only superior roles per hierarchy; ordered by role then seniority rank then ready date.

## RTIS uploads

The sidebar RTIS tab (`/rtis`) accepts division-wise `.xlsx` exports for SDAH,
HWH, ASN and MLDT. Each division has its own labeled Excel file input and upload
button. Choose files under the matching division and upload up to 10 files
(25 MB each) per upload, then
filter H/J/K events by event date. Passenger Train Analysis is the default; Goods
Train Analysis is a separate option. Following the user's rule, digit-only train
numbers are passenger trains, while numbers containing both letters and digits
are goods trains. Blank and unmatched identifiers remain in Unclassified events.
The selection is retained across filters, pagination and uploads.
The date selector filters Event Time. Speed options are All speeds, 30 and above,
40 and above, and 50 and above (inclusive thresholds). Active speed filters omit
blank speeds and apply to the event counts and results in both analysis options.
All event types and original files are retained
in the application's configured SQLite database. The upload history offers the
original download. Existing app login is required for viewing and uploading.

Imports validate every row's Division Code and the standard RTIS export headers.
Blank speed remains blank, and train identifiers stored as text retain leading
zeros. Re-uploading identical files does not add events. Overlapping exports are
deduplicated by division, device, loco, station, event time/type and coordinates;
the first imported event is retained. J and K remain source codes pending a
confirmed definition. Event dates come from cell values, not the filename.

Restart the app after installing this change; its normal startup creates the
RTIS tables. Tests use an isolated database: install `httpx` and `pypdf` in the development
environment and run `python scripts/test_rtis.py`. Optionally pass the supplied
`RTIS_Events_SDAH_2026-09-14.xlsx` path to exercise the reference export too.

### Combined RTIS output

Click **Output** on RTIS to view `/rtis/output`: one table containing all saved
events from SDAH, HWH, ASN and MLDT, including all event types, train types and
dates. The 13 original Excel headers and source serial numbers are retained.
Existing import deduplication still applies. The browser shows 100 rows per page.
**Generate PDF** and **Download Excel** export the entire combined dataset,
including every page. Excel keeps train/device/loco identifiers as text and
provides a frozen header and filters. PDF uses landscape A3 pages with repeated
headers and page numbers. These routes require the same login as the rest of RTIS.

## Customize
- Edit `config/requirements.json` to change required numbers (used for seeding).
- Add or modify rows in `data/employees.csv` with ISO dates (YYYY-MM-DD) for seeding.
- Adjust `--horizon-months` and `--lead-time-days` (CLI) or form fields (web) per policy.
