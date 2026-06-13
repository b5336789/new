# Query Engine

The query engine turns a declarative `QuerySpec` into SQL and runs it. It is
split into a **pure builder** (`query_builder.py`, no I/O) and an **executor**
(`engine.py`, all I/O), which makes the builder trivially unit-testable.

## QuerySpec

| Field | Meaning |
|-------|---------|
| `dimensions` | Columns to `GROUP BY` (when metrics are present) and select. |
| `metrics` | `[{column, aggregate, label}]`; aggregate ∈ `SUM,AVG,MIN,MAX,COUNT,COUNT_DISTINCT`. `column=null` with `COUNT` ⇒ `COUNT(*)`. |
| `filters` | `[{column, op, value}]`; ops include `=,!=,<,<=,>,>=,IN,NOT IN,LIKE,IS NULL,IS NOT NULL`. |
| `order_by` | `[{field, desc}]`; `field` may be a metric label or a column. |
| `row_limit` | Capped by `MAX_ROW_LIMIT` (50,000). |
| `time_grain` | Optional: `day,week,month,quarter,year`. |
| `time_column` | The datetime column truncated to `time_grain`. |

## Safety model

Two distinct trust boundaries:

1. **Identifiers (column / table names)** are *validated against the dataset's
   known column list* before being placed into SQL, then double-quoted. A spec
   referencing an unknown column raises `QueryBuildError` → HTTP 400. This is
   the primary defense against SQL injection through column names.
2. **Literal values** (filter values) are **never** interpolated into the SQL
   string. They are passed as bound parameters (`:f0`, `:f1_0`, …) to the
   driver.

```python
# query_builder.build_query — abbreviated
require(col)                       # raise if col not in dataset.columns
where_parts.append(f'{quote(col)} {op} :{pk}')
params[pk] = filter.value          # bound, not inlined
```

## Time grain

Superset's signature temporal bucketing. The builder is dialect-aware:

- **SQLite** (default / examples / uploads): `strftime` formats —
  `month` → `strftime('%Y-%m', col)`, `quarter` derived from month number, etc.
- **PostgreSQL**: native `date_trunc('month', col)`.

When `time_grain` + `time_column` are set, the truncated expression becomes the
leading dimension (aliased to the column name), and it is added to `GROUP BY`.

```sql
-- monthly sales, SQLite
SELECT strftime('%Y-%m', "order_date") AS "order_date",
       SUM("sales") AS "monthly_sales"
FROM "sales"
GROUP BY strftime('%Y-%m', "order_date")
LIMIT 1000
```

> Note: spreadsheet-loaded dates often arrive as ISO strings (`2023-01-15`).
> SQLite's `strftime` operates on these directly, so time grain works on
> uploaded data without an explicit date type.

## Executor (`engine.py`)

- `get_engine(uri)` — LRU-cached SQLAlchemy engine per URI, `pool_pre_ping`.
- `dialect_of(uri)` — returns `"sqlite"`, `"postgresql"`, … for the builder.
- `inspect_columns(uri, table_name=… | sql=…)` — introspects a physical table,
  or runs `SELECT * FROM (<sql>) LIMIT 0` to read a virtual dataset's schema.
- `run_sql(uri, sql, params)` — executes and returns `{columns, rows}`.

## SQL Lab guardrail

`/api/sql/run` only permits statements beginning with `SELECT` or `WITH`, and
wraps them as `SELECT * FROM (<user sql>) LIMIT n`. Mutations (`DELETE`,
`UPDATE`, `DROP`, …) are rejected with HTTP 400 to protect the source.
