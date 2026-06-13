"""Translate a QuerySpec into safe parameterized SQL for a dataset.

Identifier safety: every column referenced in dimensions/metrics/filters/order_by
is validated against the dataset's known column list before it is placed into the
SQL string. Literal values are always passed as bound parameters, never inlined.
"""
from .config import MAX_ROW_LIMIT
from .models import Dataset
from .schemas import QuerySpec

_AGG_TEMPLATES = {
    "SUM": "SUM({col})",
    "AVG": "AVG({col})",
    "MIN": "MIN({col})",
    "MAX": "MAX({col})",
    "COUNT": "COUNT({col})",
    "COUNT_DISTINCT": "COUNT(DISTINCT {col})",
}

_VALUE_OPS = {"=", "!=", ">", ">=", "<", "<=", "LIKE"}
_LIST_OPS = {"IN", "NOT IN"}
_NULL_OPS = {"IS NULL", "IS NOT NULL"}


class QueryBuildError(ValueError):
    """Raised when a QuerySpec references unknown columns or is malformed."""


def _quote_ident(name: str) -> str:
    # Double-quote and escape embedded quotes (standard SQL identifier quoting).
    return '"' + name.replace('"', '""') + '"'


def _source_clause(dataset: Dataset) -> str:
    if dataset.table_name:
        return _quote_ident(dataset.table_name)
    # Virtual dataset: wrap the saved SQL as a subquery.
    return f"({dataset.sql.rstrip().rstrip(';')}) AS virtual_source"


# SQLite strftime() formats per grain. SQLite has no date_trunc, so we format
# the timestamp into a sortable string bucket.
_SQLITE_GRAIN = {
    "day": "strftime('%Y-%m-%d', {col})",
    "week": "strftime('%Y-W%W', {col})",
    "month": "strftime('%Y-%m', {col})",
    "year": "strftime('%Y', {col})",
    # quarter: derive Q1-Q4 from the month number.
    "quarter": "strftime('%Y', {col}) || '-Q' || "
               "((cast(strftime('%m', {col}) as integer) + 2) / 3)",
}


def _time_grain_expr(dialect: str, col_quoted: str, grain: str) -> str:
    """Return a SQL expression that truncates col_quoted to the given grain."""
    if dialect == "postgresql":
        # Postgres has native date_trunc for all supported grains.
        return f"date_trunc('{grain}', {col_quoted})"
    # Default / SQLite path.
    template = _SQLITE_GRAIN.get(grain)
    if not template:
        raise QueryBuildError(f"Unsupported time_grain: {grain}")
    return template.format(col=col_quoted)


def build_query(dataset: Dataset, spec: QuerySpec, dialect: str = "sqlite") -> tuple[str, dict]:
    """Return (sql, params). Raises QueryBuildError on invalid column references."""
    known = {c["name"] for c in (dataset.columns or [])}

    def require(col: str):
        if col not in known:
            raise QueryBuildError(
                f"Unknown column '{col}'. Known columns: {sorted(known)}"
            )

    select_parts: list[str] = []
    # Label -> SQL expression, used to resolve ORDER BY against aliases.
    label_exprs: dict[str, str] = {}
    # Grouping expressions for dimensions (kept separate so we can GROUP BY exprs).
    group_exprs: list[str] = []

    # Temporal dimension first, if a time grain is requested.
    if spec.time_grain and spec.time_column:
        require(spec.time_column)
        expr = _time_grain_expr(dialect, _quote_ident(spec.time_column), spec.time_grain)
        select_parts.append(f"{expr} AS {_quote_ident(spec.time_column)}")
        label_exprs[spec.time_column] = expr
        group_exprs.append(expr)

    for dim in spec.dimensions:
        if dim == spec.time_column and spec.time_grain:
            continue  # already added as the truncated temporal dimension
        require(dim)
        select_parts.append(_quote_ident(dim))
        label_exprs[dim] = _quote_ident(dim)
        group_exprs.append(_quote_ident(dim))

    for i, metric in enumerate(spec.metrics):
        if metric.aggregate == "COUNT" and metric.column is None:
            expr = "COUNT(*)"
        else:
            if metric.column is None:
                raise QueryBuildError(f"Metric #{i} requires a column for {metric.aggregate}")
            require(metric.column)
            expr = _AGG_TEMPLATES[metric.aggregate].format(col=_quote_ident(metric.column))
        label = metric.label or f"{metric.aggregate.lower()}_{metric.column or 'all'}"
        select_parts.append(f"{expr} AS {_quote_ident(label)}")
        label_exprs[label] = expr

    if not select_parts:
        select_parts.append("*")

    params: dict = {}
    where_parts: list[str] = []
    for i, f in enumerate(spec.filters):
        require(f.column)
        col = _quote_ident(f.column)
        if f.op in _NULL_OPS:
            where_parts.append(f"{col} {f.op}")
        elif f.op in _LIST_OPS:
            values = f.value if isinstance(f.value, (list, tuple)) else [f.value]
            keys = []
            for j, v in enumerate(values):
                pk = f"f{i}_{j}"
                params[pk] = v
                keys.append(f":{pk}")
            where_parts.append(f"{col} {f.op} ({', '.join(keys)})")
        elif f.op in _VALUE_OPS:
            pk = f"f{i}"
            params[pk] = f.value
            where_parts.append(f"{col} {f.op} :{pk}")
        else:
            raise QueryBuildError(f"Unsupported filter op: {f.op}")

    group_parts = group_exprs if spec.metrics else []

    order_parts: list[str] = []
    for ob in spec.order_by:
        if ob.field in label_exprs:
            order_parts.append(f"{_quote_ident(ob.field)} {'DESC' if ob.desc else 'ASC'}")
        elif ob.field in known:
            order_parts.append(f"{_quote_ident(ob.field)} {'DESC' if ob.desc else 'ASC'}")
        else:
            raise QueryBuildError(f"Cannot order by unknown field '{ob.field}'")

    limit = max(1, min(spec.row_limit or 1000, MAX_ROW_LIMIT))

    sql = f"SELECT {', '.join(select_parts)} FROM {_source_clause(dataset)}"
    if where_parts:
        sql += " WHERE " + " AND ".join(where_parts)
    if group_parts:
        sql += " GROUP BY " + ", ".join(group_parts)
    if order_parts:
        sql += " ORDER BY " + ", ".join(order_parts)
    sql += f" LIMIT {limit}"
    return sql, params
