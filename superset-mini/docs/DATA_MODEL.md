# Data Model

The metadata store holds four entities, defined in `app/models.py`.

```
Database 1───* Dataset 1───* Chart *───1 Dashboard (referenced by layout)
```

## Database
A connection to a physical data source.

| Field | Type | Notes |
|-------|------|-------|
| `id` | int (PK) | |
| `name` | str (unique) | Display name |
| `sqlalchemy_uri` | str | e.g. `sqlite:///…`, `postgresql+psycopg2://…` |
| `created_at` | datetime | |

Deleting a Database cascades to its Datasets.

## Dataset
A queryable source within a Database. Exactly one of `table_name` / `sql` is set.

| Field | Type | Notes |
|-------|------|-------|
| `id` | int (PK) | |
| `database_id` | FK → Database | |
| `name` | str | |
| `table_name` | str \| null | Physical table (mutually exclusive with `sql`) |
| `sql` | text \| null | Virtual dataset defined by a SELECT |
| `columns` | JSON | Cached `[{"name","type"}]`; `type ∈ {number, string, datetime, boolean}` |
| `created_at` | datetime | |

`columns` is cached at creation and refreshable via `POST /api/datasets/{id}/refresh`.
Column type normalization happens in `engine._normalize_type`.

## Chart
A saved visualization: a dataset + viz type + query spec.

| Field | Type | Notes |
|-------|------|-------|
| `id` | int (PK) | |
| `name` | str | |
| `dataset_id` | FK → Dataset | |
| `viz_type` | str | `table, bar, line, area, pie, scatter, big_number` |
| `params` | JSON | A serialized [`QuerySpec`](./QUERY_ENGINE.md) |
| `created_at` | datetime | |

## Dashboard
A grid arrangement of charts.

| Field | Type | Notes |
|-------|------|-------|
| `id` | int (PK) | |
| `name` | str | |
| `layout` | JSON | `[{"chart_id","x","y","w","h"}]` |
| `created_at` | datetime | |

`layout` references chart IDs; the API validates each referenced chart exists
on create/update. Charts are not deleted when removed from a layout.

## QuerySpec (embedded JSON)

Stored inside `Chart.params`; the contract between frontend, query builder, and
text-to-chart. See [QUERY_ENGINE.md](./QUERY_ENGINE.md) for full semantics.

```jsonc
{
  "dimensions": ["region"],
  "metrics": [{"column": "sales", "aggregate": "SUM", "label": "total_sales"}],
  "filters": [{"column": "year", "op": "=", "value": 2023}],
  "order_by": [{"field": "total_sales", "desc": true}],
  "row_limit": 1000,
  "time_grain": "month",       // optional temporal bucketing
  "time_column": "order_date"  // required when time_grain is set
}
```
