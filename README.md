# WELL Ledger

A local web application for WELL ongoing-work planning, evidence storage and annual packages. Built around the supplied WELL v2 Q1 2021 19-part service guide and version-specific Excel scorecard.

## Run

Python 3.12 was tested.

```bash
cd WELL-project-manager
python -m venv .venv
.venv/bin/python -m pip install -r requirements.lock
.venv/bin/python -m uvicorn app:app --host 127.0.0.1 --port 8000
```

On your own computer, open `http://127.0.0.1:8000` in your browser. Keep the terminal running while you use the app. The cloud onboarding UI does not provide a localhost preview.

On Windows (PowerShell), use:

```powershell
cd WELL-project-manager
py -3.12 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.lock
.\.venv\Scripts\python.exe -m uvicorn app:app --host 127.0.0.1 --port 8000
```

Download this repository using GitHub's **Code → Download ZIP** button, extract it, and open a terminal inside the extracted folder (its name may be `WELL-project-manager-main`). Install Python 3.12 first if needed. The original uploaded PDF and scorecard are not included in the repository; import your own copies in the app.

Import the scorecard `.xlsx` and service guide `.pdf`. The provided files have already been imported into this instance's ignored `data/` directory. New clones do not contain user uploads. Review credit selection against the final WELL report, inspect source differences, set certification/occupancy dates, then manage work in the Gantt chart.

## Files and annual export

All project files default to **`data/` beside `app.py`**, regardless of the working directory. SQLite stores the task register; `data/projects/<id>/sources/` retains imports, `evidence/` holds uploaded evidence, and `exports/` holds annual ZIP archives. `WELL_DATA_DIR` optionally overrides that root. Back up the entire directory. Source evidence is not committed to Git.

Attach evidence to each task, enter factual activity notes and record actual submission references separately. Marking work complete requires an attachment; it does not validate its contents. Download annual packages using the export button. ZIPs contain `narrative.md`, `task-register.csv`, `manifest.json`, source documents and evidence. Narratives explicitly identify tracking gaps and require human review.

Automatic year-end archives are optional and disabled initially. Enter and confirm the certification date/reporting basis, then enable automatic export. The running server checks hourly; on restart it catches up on closed periods. It creates a historical archive once per period; use manual export for revised evidence. Live processes must restart after cloud environment restoration.

## Scope and limitations

- Eight parts are selected in the supplied scorecard; target selections are not certification awards.
- The workbook's explicit recurring verification requirements take precedence where the secondary service guide differs. Review all credit rules before relying on deadlines.
- Workbook import currently supports the supplied Matrix Summary / Feature Requirements layout and the known Q1 2021 guide. Other layouts/guides are rejected instead of silently guessed.
- Unknown layouts, feature versions, reviewer-specific alternatives and building-specific plan frequencies need a reviewed mapping or manual task additions.
- Task defaults are planning windows. Confirm deadlines in WELL Online, especially registration anchors, calibration history and recertification dates. T02.1 survey eligibility must be confirmed against occupancy date.
- Seasonal work is separate from annual reporting. Event-based and one-time parts are review placeholders; add real events manually and deactivate fulfilled one-time parts.
- Changing the reporting basis after activity has been recorded is blocked to preserve evidence. Edit task dates individually instead.
- No WELL Online submission integration, laboratory interpretation or automatic compliance determination.
- This is a **single-user local prototype** without account authentication, encryption or malware scanning. Keep it bound to loopback; deployment for team use requires those capabilities. Upload only anonymised survey summaries and use encrypted device storage where needed.
- Independent official web research was blocked by cloud network policy. See [research findings](docs/research.md) for evidence and open checks.

## Validation

```bash
.venv/bin/python -m pytest -q
node --check static/app.js
```

The tests exercise scorecard mapping, seasonal/reporting separation, evidence requirements, safe filenames, annual ZIP contents/hashes, schedule protection and repeatable automatic export. User-interface browser checks are separate from backend validation.
