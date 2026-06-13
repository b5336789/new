# API Reference

Base URL `http://localhost:8000`. Interactive docs (OpenAPI/Swagger) at `/docs`.
All bodies are JSON unless noted. Errors return `{"detail": "..."}` with an
appropriate HTTP status.

## Health
| Method | Path | Description |
|--------|------|-------------|
| GET | `/api/health` | `{status, text_to_chart}` — `text_to_chart` is `claude` or `heuristic_fallback`. |

## Databases
| Method | Path | Body / Notes |
|--------|------|--------------|
| GET | `/api/databases` | List connections. |
| POST | `/api/databases` | `{name, sqlalchemy_uri}` — tests the connection before saving (400 on failure). |
| GET | `/api/databases/{id}/tables` | List table names in the source. |
| DELETE | `/api/databases/{id}` | Cascades to datasets. |

## Datasets
| Method | Path | Body / Notes |
|--------|------|--------------|
| GET | `/api/datasets` | List. |
| GET | `/api/datasets/{id}` | Single, with cached columns. |
| POST | `/api/datasets` | `{database_id, name, table_name?\|sql?}` — exactly one of `table_name`/`sql`. Introspects columns. |
| POST | `/api/datasets/{id}/refresh` | Re-introspect and cache columns. |
| DELETE | `/api/datasets/{id}` | |

## Upload
| Method | Path | Notes |
|--------|------|-------|
| POST | `/api/upload` | `multipart/form-data`: `file` (.csv/.xlsx/.xls) + `dataset_name`. Loads rows into `uploads.db`, sanitizes headers, registers a dataset. Returns the new `Dataset`. |

## SQL Lab
| Method | Path | Body |
|--------|------|------|
| POST | `/api/sql/run` | `{database_id, sql, row_limit}` — read-only `SELECT`/`WITH` only. Returns `{columns, rows, sql, row_count}`. |

## Charts
| Method | Path | Body / Notes |
|--------|------|--------------|
| GET | `/api/charts` | List. |
| GET | `/api/charts/{id}` | Single. |
| POST | `/api/charts` | `{name, dataset_id, viz_type, params}` (`params` = QuerySpec). |
| PUT | `/api/charts/{id}` | `{name?, viz_type?, params?}`. |
| DELETE | `/api/charts/{id}` | |
| GET | `/api/charts/{id}/data` | Run the saved query → `QueryResult`. |
| GET | `/api/charts/{id}/data.csv` | Same, as a CSV download. |
| POST | `/api/charts/explore` | Run an unsaved `{dataset_id, viz_type, params}` → `QueryResult` (live preview). |
| POST | `/api/charts/explore.csv` | Same, as CSV. |

## Dashboards
| Method | Path | Body / Notes |
|--------|------|--------------|
| GET | `/api/dashboards` | List. |
| GET | `/api/dashboards/{id}` | Single. |
| POST | `/api/dashboards` | `{name, layout:[{chart_id,x,y,w,h}]}` — validates chart IDs. |
| PUT | `/api/dashboards/{id}` | `{name?, layout?}`. |
| DELETE | `/api/dashboards/{id}` | |

## Text → Chart
| Method | Path | Body / Notes |
|--------|------|--------------|
| POST | `/api/nl/chart` | `{dataset_id, prompt}` → `{viz_type, params, explanation, source, result}`. Generates a spec (Claude or fallback), then validates + executes it. |

## Common shapes

**QueryResult**
```json
{ "columns": ["region","total"], "rows": [{"region":"East","total":123}],
  "sql": "SELECT ...", "row_count": 1 }
```

**QuerySpec** — see [DATA_MODEL.md](./DATA_MODEL.md) / [QUERY_ENGINE.md](./QUERY_ENGINE.md).
