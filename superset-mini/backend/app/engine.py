"""Connections to *user* data sources (not the metadata store).

Engines are cached per URI. We never interpolate user values into SQL strings
here — callers pass bound parameters. Identifier safety is enforced in
query_builder by validating against a dataset's known columns.
"""
from functools import lru_cache

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.engine import Engine


@lru_cache(maxsize=64)
def get_engine(sqlalchemy_uri: str) -> Engine:
    connect_args = {}
    if sqlalchemy_uri.startswith("sqlite"):
        connect_args["check_same_thread"] = False
    return create_engine(sqlalchemy_uri, connect_args=connect_args, pool_pre_ping=True)


def dialect_of(sqlalchemy_uri: str) -> str:
    """Return the SQLAlchemy dialect name (e.g. 'sqlite', 'postgresql')."""
    return get_engine(sqlalchemy_uri).dialect.name


def test_connection(sqlalchemy_uri: str) -> None:
    """Raise if the connection can't be established."""
    eng = get_engine(sqlalchemy_uri)
    with eng.connect() as conn:
        conn.execute(text("SELECT 1"))


def list_tables(sqlalchemy_uri: str) -> list[str]:
    eng = get_engine(sqlalchemy_uri)
    return sorted(inspect(eng).get_table_names())


def _normalize_type(sa_type) -> str:
    """Collapse SQLAlchemy/DBAPI types into coarse buckets the UI understands."""
    t = str(sa_type).upper()
    if any(k in t for k in ("INT", "DECIMAL", "NUMERIC", "FLOAT", "REAL", "DOUBLE")):
        return "number"
    if any(k in t for k in ("DATE", "TIME")):
        return "datetime"
    if "BOOL" in t:
        return "boolean"
    return "string"


def inspect_columns(sqlalchemy_uri: str, *, table_name: str | None = None,
                    sql: str | None = None) -> list[dict]:
    """Return [{"name", "type"}] for a physical table or a virtual SQL source."""
    eng = get_engine(sqlalchemy_uri)
    if table_name:
        cols = inspect(eng).get_columns(table_name)
        return [{"name": c["name"], "type": _normalize_type(c["type"])} for c in cols]

    if sql:
        # Wrap the user SQL and pull zero rows just to read the result schema.
        probe = f"SELECT * FROM ({sql.rstrip().rstrip(';')}) AS _t LIMIT 0"
        with eng.connect() as conn:
            result = conn.execute(text(probe))
            return [{"name": k, "type": "string"} for k in result.keys()]

    raise ValueError("Either table_name or sql must be provided")


def run_sql(sqlalchemy_uri: str, sql: str, params: dict | None = None) -> dict:
    """Execute SQL and return {"columns": [...], "rows": [...]}."""
    eng = get_engine(sqlalchemy_uri)
    with eng.connect() as conn:
        result = conn.execute(text(sql), params or {})
        columns = list(result.keys())
        rows = [dict(zip(columns, row)) for row in result.fetchall()]
    return {"columns": columns, "rows": rows}
