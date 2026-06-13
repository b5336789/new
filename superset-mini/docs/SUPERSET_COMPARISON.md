# Comparison to Apache Superset

Superset-mini deliberately reproduces Superset's **core concepts and workflow**,
not its full surface area. This document is honest about what maps directly,
what is simplified, and what is intentionally omitted.

## Concept mapping

| Apache Superset | Superset-mini | Notes |
|-----------------|---------------|-------|
| Metadata DB (Postgres/MySQL) | `metadata.db` (SQLite) | Same separation of metadata vs analytics DBs. |
| Database connection | `Database` | SQLAlchemy URI + connection test. |
| Dataset (physical / virtual) | `Dataset` | `table_name` or `sql`; cached columns. |
| SQL Lab | `/api/sql` + SQL Lab page | Read-only SELECT/WITH. |
| Explore / chart builder | `Explore` page + `QuerySpec` | Dimensions, metrics, filters, time grain. |
| Saved chart (slice) | `Chart` | viz_type + params. |
| Dashboard | `Dashboard` | Grid layout of charts. |
| Time grain | `time_grain` / `time_column` | `strftime` (SQLite) / `date_trunc` (Postgres). |
| CSV export | `/data.csv`, `/explore.csv` | |
| Viz plugins (~50+) | 7 viz types (Recharts) | table, bar, line, area, pie, scatter, big number. |
| — (no equivalent) | **Text → Chart** | NL → QuerySpec via Claude. |

## Intentionally simplified / omitted

These exist in Superset but are out of scope here; each is a clear extension
point rather than a hidden gap:

- **AuthN/AuthZ & RBAC** — no users, roles, or row-level security. Single-tenant.
- **Async query execution** (Celery workers, results backend) — queries run
  synchronously in-request.
- **Caching layer** (Redis) — every request hits the source; engines are pooled.
- **Dashboard native filters & cross-filtering** — layout is static positions.
- **Drag-and-drop dashboard editor** — layout is auto-arranged 2-up.
- **Jinja templating in SQL**, saved queries, query history.
- **The full viz plugin ecosystem** (maps, pivot tables, deck.gl, etc.).
- **Alerts & reports, annotations, CSS theming, i18n.**
- **Semantic layer** features: metrics defined on datasets, calculated columns.

## Where the fidelity is real

- The **metadata-vs-source split** is faithful and is what makes connecting
  arbitrary SQLAlchemy databases work the same way Superset does.
- The **declarative QuerySpec → SQL** compilation (group-bys, aggregate metrics,
  filters, ordering, **dialect-aware time grain**) reflects how Superset builds
  queries from Explore controls.
- The **virtual dataset** mechanism (a saved SELECT wrapped as a subquery) is
  the same idea as Superset's SQL-defined datasets.

## Why a clone instead of the real thing

The brief was to build a complete-feature Superset *from scratch*. Vendoring
real Superset would not demonstrate how the pieces fit. This reimplementation
keeps the architecture recognizable while remaining small enough to read end to
end in an afternoon — and adds the text-to-chart capability on top.
