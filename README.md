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

### Excel upload format
- Employees: headers `name`, `role`, `hire_date`, `retirement_date`; optional `promotion_role`, `promotion_ready_date`.
- Seniority: headers `name`, `role`, `seniority_rank` (or `seniority`); optional `promotion_role`, `promotion_ready_date`. Rows update matching employees.
- Dates can be ISO strings (YYYY-MM-DD) or Excel date cells. Role aliases like `Shunter`/`LPS(Shunter)` normalize to `LPS`.
- Promotion ordering: only superior roles per hierarchy; ordered by role then seniority rank then ready date.

## Customize
- Edit `config/requirements.json` to change required numbers (used for seeding).
- Add or modify rows in `data/employees.csv` with ISO dates (YYYY-MM-DD) for seeding.
- Adjust `--horizon-months` and `--lead-time-days` (CLI) or form fields (web) per policy.