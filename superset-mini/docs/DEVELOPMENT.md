# Development Guide

## Prerequisites
- Python 3.11+
- Node 18+ (built/tested on Node 22)

## Backend

```bash
cd backend
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

- On startup the app creates metadata tables and seeds the **Sample Sales**
  dataset (2000 rows) into `examples.db` (idempotent — runs once).
- OpenAPI docs: http://localhost:8000/docs
- Data files live under `backend/data/` and are git-ignored. Delete them to
  reset all state.

### Environment variables
| Var | Default | Purpose |
|-----|---------|---------|
| `ANTHROPIC_API_KEY` | (unset) | Enables the Claude text-to-chart path. Without it, the heuristic fallback is used. |
| `ANTHROPIC_MODEL` | `claude-opus-4-8` | Model for text-to-chart. |

## Frontend

```bash
cd frontend
npm install
npm run dev      # http://localhost:5173, proxies /api → :8000
npm run build    # production bundle in dist/
```

The Vite dev server proxies `/api` to the backend, so the SPA uses relative
URLs and there is no CORS friction in development.

## Tests

```bash
cd backend && source .venv/bin/activate
python -m pytest -q
```

The suite (`tests/test_api.py`) is end-to-end against a `TestClient`, pointed at
a throwaway data directory, with `ANTHROPIC_API_KEY` cleared so text-to-chart
exercises the deterministic fallback. It asserts **business outcomes**, e.g.:

- `SUM(sales)` grouped by region returns 4 ordered rows with correct totals.
- An unknown column is rejected (HTTP 400).
- SQL Lab rejects `DELETE`.
- Excel upload sanitizes headers and the result is immediately queryable
  (`A = 10 + 30 = 40`).
- Time grain buckets sample data into ≤24 `YYYY-MM` months.
- Chart CSV export emits `header + 4` rows.

## Project layout

```
superset-mini/
├── backend/app/{main,config,database,models,schemas,engine,query_builder}.py
│   ├── routers/{databases,datasets,sql,charts,dashboards,uploads,nl}.py
│   └── services/text_to_chart.py
├── backend/tests/test_api.py
├── frontend/src/{api.js, App.jsx, components/, pages/}
├── docs/                 # the reference docs you are reading
└── docs/site/index.html  # the documentation website
```

## Conventions
- Backend: `snake_case`, type hints, Pydantic for I/O schemas, thin routers.
- Frontend: functional React components, hooks, a single `api.js` client.
- The query builder is pure (no I/O) and must stay that way — it keeps query
  logic unit-testable and the executor (`engine.py`) the only I/O boundary.

## Extending it
- **New viz type**: add to `VizType` (schemas) + `VIZ_TYPES` (Explore) + a branch
  in `components/ChartView.jsx`.
- **New database**: just add a SQLAlchemy URI in the Databases page (install the
  driver if needed, e.g. `psycopg2-binary` for Postgres — already included).
- **New aggregate / filter op**: extend the enums in `schemas.py` and the
  templates in `query_builder.py`.
