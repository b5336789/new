# Architecture

Superset-mini mirrors Apache Superset's real architecture in miniature: a
**Python backend** that owns data-source connectivity, a query engine, and a
metadata store; and a **React frontend** that is a thin client over a JSON API.

```
┌──────────────────────────────┐         ┌───────────────────────────────────┐
│  Frontend (React + Vite)      │  HTTP   │  Backend (FastAPI)                  │
│                               │  JSON   │                                     │
│  pages/  ─ Databases          │ ──────► │  routers/  ─ /api/databases         │
│          ─ Datasets+Upload    │         │            ─ /api/datasets          │
│          ─ SQL Lab            │         │            ─ /api/sql               │
│          ─ Explore (builder)  │         │            ─ /api/charts            │
│          ─ Text → Chart       │         │            ─ /api/dashboards        │
│          ─ Charts             │         │            ─ /api/upload            │
│          ─ Dashboards         │         │            ─ /api/nl  (text→chart)  │
│  components/ ChartView        │         │                                     │
│      (Recharts)               │         │  query_builder ─ QuerySpec → SQL    │
└──────────────────────────────┘         │  engine        ─ connect/introspect │
                                          │  services/text_to_chart ─ Claude    │
                                          └───────────────┬─────────────────────┘
                                                          │ SQLAlchemy
                          ┌───────────────────────────────┼───────────────────┐
                          ▼                               ▼                   ▼
                   metadata.db                      examples.db          uploads.db
                (databases, datasets,            (seeded sample          (Excel/CSV
                 charts, dashboards)              sales data)            imports)
```

## Layers

1. **Routers** (`app/routers/*`) — HTTP surface. Validation via Pydantic,
   error mapping to HTTP status codes. No business logic beyond orchestration.
2. **Query builder** (`app/query_builder.py`) — pure function translating a
   `QuerySpec` into safe, parameterized SQL. No I/O.
3. **Engine** (`app/engine.py`) — the only place that talks to *user* data
   sources: connection pooling, schema introspection, SQL execution.
4. **Metadata store** (`app/models.py` + `app/database.py`) — SQLAlchemy ORM
   describing Superset-mini's own objects, persisted in `metadata.db`.
5. **Text-to-chart service** (`app/services/text_to_chart.py`) — converts NL to
   a `QuerySpec` via the Claude API, with a deterministic fallback.

## Two databases, one pattern

There is a deliberate separation between:

- **The metadata store** — Superset-mini's own bookkeeping (what databases,
  datasets, charts and dashboards exist). Always SQLite (`metadata.db`).
- **User data sources** — the actual data you analyze. Any SQLAlchemy URI.
  The seeded `examples.db` and the `uploads.db` are just *registered* user
  sources; there is nothing special about them beyond being created for you.

This is exactly how Apache Superset separates its metadata DB from analytics
databases.

## Request lifecycle (chart preview)

1. Frontend `Explore` builds a `QuerySpec` and POSTs to `/api/charts/explore`.
2. Router loads the `Dataset` (and its `Database`) from the metadata store.
3. `query_builder.build_query(dataset, spec, dialect)` validates every column
   reference and emits parameterized SQL.
4. `engine.run_sql(uri, sql, params)` executes it against the user data source.
5. Rows + the generated SQL are returned; `ChartView` renders them by viz type.

See [QUERY_ENGINE.md](./QUERY_ENGINE.md) for the builder details and
[DATA_MODEL.md](./DATA_MODEL.md) for the object graph.
