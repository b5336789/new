"""SQL Lab-style ad-hoc query execution."""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from .. import engine as data_engine
from ..config import MAX_ROW_LIMIT
from ..database import get_db
from ..models import Database
from ..schemas import QueryResult, SqlRunRequest

router = APIRouter(prefix="/api/sql", tags=["sql"])


@router.post("/run", response_model=QueryResult)
def run_sql(payload: SqlRunRequest, db: Session = Depends(get_db)):
    database = db.get(Database, payload.database_id)
    if not database:
        raise HTTPException(404, "Database not found")

    sql = payload.sql.strip().rstrip(";")
    if not sql:
        raise HTTPException(400, "Empty SQL")
    # SQL Lab is read-only here: reject obvious mutations to protect the source.
    first_word = sql.lstrip("(").split(None, 1)[0].upper() if sql else ""
    if first_word not in {"SELECT", "WITH"}:
        raise HTTPException(400, "Only SELECT / WITH queries are allowed in SQL Lab")

    limit = max(1, min(payload.row_limit or 1000, MAX_ROW_LIMIT))
    wrapped = f"SELECT * FROM ({sql}) AS _q LIMIT {limit}"
    try:
        out = data_engine.run_sql(database.sqlalchemy_uri, wrapped)
    except Exception as exc:
        raise HTTPException(400, f"Query failed: {exc}")
    return QueryResult(
        columns=out["columns"], rows=out["rows"], sql=wrapped, row_count=len(out["rows"])
    )
