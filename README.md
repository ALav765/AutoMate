# AutoMate

An automated supply chain planning tool. Upload five input spreadsheets, AutoMate builds an integrated supply plan, runs a 15-point checklist against it, automatically fixes calculation errors using Claude, and flags genuine business alerts for review.

## How it works

1. **Run a plan** — upload Forecast & FG SOH, Bill of Materials, Component SOH, Component Prices, and Vendor Master Excel files for a given month.
2. **Build** — `build_planner.py` constructs an Integrated Supply Plan Excel file from the five inputs.
3. **Check** — `checklist_engine.py` runs 15 checks against the plan (production bands, component consumption thresholds, cross-sheet totals, etc).
4. **Fix** — if any check fails with a genuine calculation error, Claude (`claude_agent.py`) diagnoses the issue, proposes a precise code patch, and the pipeline rebuilds and re-checks automatically (up to 5 fix loops).
5. **Flag** — failures that are not calculation errors (e.g. a forecast genuinely exceeding a threshold) are written to an Alerts & Flags sheet in the final output, not silently "fixed."
6. **Download** — the finished Integrated Supply Plan, with its Alerts & Flags sheet, is available to download.

You can also add new checks through the UI — describe the rule in plain language, and Claude writes the corresponding check logic.

## Tech stack

- **Frontend:** React + Vite, Tailwind CSS, react-router
- **Backend:** FastAPI (Python), Server-Sent Events for live pipeline progress
- **Database:** PostgreSQL — stores checks and job history
- **AI:** Anthropic Claude API — diagnoses calc errors, writes new checks

## Project structure

```
AutoMate/
├── src/                        # React frontend
│   ├── pages/                  # Run Plan, Checks, History
│   └── components/              # FileUploadZone, RunStatus, ChecklistPanel, DownloadCard
└── supply-chain-backend/        # FastAPI backend
    ├── main.py                  # API routes
    ├── runner.py                 # Pipeline orchestration (build → check → fix → alerts)
    ├── build_planner.py          # Builds the supply plan from input files
    ├── checklist_engine.py       # Runs the 15 checks
    ├── claude_agent.py           # Calls Claude to diagnose fixes / write new checks
    ├── db.py                     # Postgres connection
    └── schema.sql / seed_checks.sql
```

## Setup

### 1. Database
```bash
createdb supply_plan
cd supply-chain-backend
psql supply_plan -f schema.sql
psql supply_plan -f seed_checks.sql
```

### 2. Backend
```bash
cd supply-chain-backend
python3 -m venv venv
source venv/bin/activate
pip install fastapi uvicorn aiofiles anthropic pandas openpyxl numpy python-multipart asyncpg python-dotenv
```

Create `supply-chain-backend/.env`:
```
DATABASE_URL=postgresql://localhost:5432/supply_plan
ANTHROPIC_API_KEY=your-key-here
```

Run it:
```bash
python3 -m uvicorn main:app --reload --port 8000
```

### 3. Frontend
```bash
npm install
npm run dev
```

Visit `http://localhost:5173`.
