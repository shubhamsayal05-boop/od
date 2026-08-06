# DriveScope (ODRIV)

Python port of the **ODRIV v29.2.1 AT** Excel/VBA drivability workbook — project setup, TRIE acquisition import, SDV classification, rating scorecard, and report export.

## Quick start

### 1. MongoDB

```bash
# Docker
docker run -d --name odriv-mongo -p 27017:27017 mongo:7

# or use docker compose (starts mongo + API + UI)
docker compose up
```

### 2. Backend

```bash
cd backend
cp .env.example .env          # MONGO_URL / DB_NAME
python3 -m pip install -r requirements-app.txt
uvicorn server:app --reload --port 8000
```

API: `http://localhost:8000/api/state`

### 3. Frontend

```bash
cd frontend
cp .env.example .env          # REACT_APP_BACKEND_URL=http://localhost:8000
yarn install
yarn start
```

UI: `http://localhost:3000`

## Workflow

1. **NEW PROJECT** — set mode, fuel, gears, milestone, PREMIUM/MAINSTREAM
2. **ADD FILE TO DATABASE** — upload a TRIE `.xlsx` (or generate demo data)
3. **CALCULATE RATING** — scores every SDV sheet and builds the global verdict
4. Review **RATING** + per-SDV sheets (scatter, status dots, criteria)
5. **CREATE REPORT** — PPTX / PDF

Unlock config sheets with password `UNLOCK`.

## Tests

```bash
# Engine unit tests (no server required)
cd backend && python3 -m pytest tests/test_engine_unit.py -q

# Full HTTP suite (server + Mongo must be running)
export REACT_APP_BACKEND_URL=http://localhost:8000
python3 -m pytest tests/test_odriv_backend.py -q
```

## Layout

| Path | Role |
|------|------|
| `backend/server.py` | FastAPI routes |
| `backend/engine/` | Classifier, scoring, importer, reports |
| `backend/engine/data/` | Seed config extracted from the xlsm |
| `frontend/src/` | Excel-like workbook UI |
| `source_artifacts/` | Original workbook + extract script |

## Notes

- Targets seed currently contains **PREMIUM** rows (matches the source workbook extract). MAINSTREAM falls back gracefully.
- ~30 SDVs are marked UNAFFECTED in priorisation grids and correctly score as P3.
- `requirements.txt` is the legacy Emergent image pin set; use `requirements-app.txt` for local installs.
