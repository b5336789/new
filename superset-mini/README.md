# Superset-mini

A compact, working clone of the core [Apache Superset](https://superset.apache.org/)
workflow, plus a **Text → Chart** feature powered by the Claude API.

It mirrors Superset's real architecture in miniature: a **Python (FastAPI)** backend
with a **SQLAlchemy** metadata store, and a **React (Vite)** single-page frontend.

> Scope note: this is *not* a full reproduction of Apache Superset (which is
> hundreds of thousands of lines). It implements the core end-to-end loop —
> connect → dataset → query → chart → dashboard — as a solid, tested foundation.

## Features

| Area | What it does |
|------|--------------|
| **Databases** | Connect any SQLAlchemy-supported source (SQLite, Postgres, …) with a connection test. |
| **Datasets** | Register a physical table or a virtual (SQL-defined) source; cached column metadata. |
| **Excel/CSV upload** | Upload `.xlsx`/`.csv`; rows are loaded into a SQLite store and auto-registered as a dataset. |
| **SQL Lab** | Run ad-hoc read-only `SELECT`/`WITH` queries and view results. |
| **Explore (chart builder)** | Pick dimensions, metrics (SUM/AVG/COUNT/…), filters, and a viz type; live preview. |
| **Visualizations** | table, bar, line, area, pie, scatter, big number (Recharts). |
| **Dashboards** | Compose saved charts onto a grid and view them together. |
| **Text → Chart** | Describe a chart in plain language; Claude produces the query spec, which is executed against your data. |

## Architecture

```
superset-mini/
├── backend/                 # FastAPI + SQLAlchemy
│   ├── app/
│   │   ├── main.py          # app + routers + startup seed
│   │   ├── models.py        # Database / Dataset / Chart / Dashboard ORM
│   │   ├── engine.py        # connect to user data sources, introspect, run SQL
│   │   ├── query_builder.py # QuerySpec -> safe parameterized SQL
│   │   ├── routers/         # databases, datasets, sql, charts, dashboards, uploads, nl
│   │   └── services/text_to_chart.py  # Claude API + heuristic fallback
│   └── tests/test_api.py    # end-to-end business-logic tests
└── frontend/                # React + Vite + Recharts
    └── src/{api.js, App.jsx, components/, pages/}
```

The metadata store, sample data, and uploads each live in their own SQLite file
under `backend/data/` (created at runtime).

## Running it

### 1. Backend

```bash
cd backend
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

On first start it seeds a **Sample Sales** dataset (2000 rows) so charts work
immediately. API docs are at http://localhost:8000/docs.

### 2. Frontend

```bash
cd frontend
npm install
npm run dev          # http://localhost:5173 (proxies /api -> :8000)
```

### 3. Text → Chart (Claude API)

The Text → Chart feature calls the Claude API. Set a key before starting the backend:

```bash
export ANTHROPIC_API_KEY=sk-ant-...
# optional: export ANTHROPIC_MODEL=claude-opus-4-8
```

**Without a key**, the feature still works using a deterministic keyword parser,
and every response is tagged `source: "heuristic_fallback"` (also shown in the UI
and `/api/health`) so you always know whether the LLM actually ran.

## Tests

```bash
cd backend && source .venv/bin/activate
python -m pytest -q
```

The suite covers aggregation correctness, unknown-column rejection, the SQL-Lab
mutation guard, Excel upload → queryable dataset, dashboard validation, and the
text-to-chart fallback producing a valid, executed chart.
