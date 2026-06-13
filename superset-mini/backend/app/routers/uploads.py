"""Upload Excel/CSV files, load them into the uploads SQLite DB, and register
them as a dataset automatically.
"""
import re
import unicodedata

import pandas as pd
from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from sqlalchemy.orm import Session

from .. import engine as data_engine
from ..config import UPLOADS_DB_URI
from ..database import get_db
from ..models import Database, Dataset
from ..schemas import DatasetOut

router = APIRouter(prefix="/api/upload", tags=["upload"])

_IDENT_RE = re.compile(r"[^0-9a-zA-Z_]+")


def _sanitize(name: str, fallback: str = "col") -> str:
    """Turn an arbitrary header/sheet name into a safe SQL identifier."""
    name = unicodedata.normalize("NFKD", str(name)).encode("ascii", "ignore").decode()
    name = _IDENT_RE.sub("_", name).strip("_").lower()
    if not name:
        name = fallback
    if name[0].isdigit():
        name = f"_{name}"
    return name


def _get_uploads_database(db: Session) -> Database:
    """Find or create the dedicated 'uploads' database connection."""
    obj = db.query(Database).filter(Database.sqlalchemy_uri == UPLOADS_DB_URI).first()
    if obj:
        return obj
    obj = Database(name="Uploads", sqlalchemy_uri=UPLOADS_DB_URI)
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


@router.post("", response_model=DatasetOut, status_code=201)
async def upload_file(
    file: UploadFile = File(...),
    dataset_name: str = Form(...),
    db: Session = Depends(get_db),
):
    filename = (file.filename or "").lower()
    raw = await file.read()
    import io

    try:
        if filename.endswith(".csv"):
            frame = pd.read_csv(io.BytesIO(raw))
        elif filename.endswith((".xlsx", ".xls")):
            frame = pd.read_excel(io.BytesIO(raw))
        else:
            raise HTTPException(400, "Unsupported file type. Use .csv, .xlsx or .xls")
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(400, f"Failed to parse file: {exc}")

    if frame.empty:
        raise HTTPException(400, "Uploaded file has no rows")

    # Sanitize column names, de-duplicating collisions.
    seen: dict[str, int] = {}
    new_cols = []
    for i, col in enumerate(frame.columns):
        base = _sanitize(col, fallback=f"col_{i}")
        if base in seen:
            seen[base] += 1
            base = f"{base}_{seen[base]}"
        else:
            seen[base] = 0
        new_cols.append(base)
    frame.columns = new_cols

    table_name = _sanitize(dataset_name, fallback="uploaded_table")
    uploads_db = _get_uploads_database(db)

    if db.query(Dataset).filter(Dataset.table_name == table_name).first():
        raise HTTPException(409, f"A dataset/table named '{table_name}' already exists")

    try:
        eng = data_engine.get_engine(UPLOADS_DB_URI)
        frame.to_sql(table_name, eng, if_exists="fail", index=False)
    except Exception as exc:
        raise HTTPException(400, f"Failed to load data: {exc}")

    columns = data_engine.inspect_columns(UPLOADS_DB_URI, table_name=table_name)
    obj = Dataset(
        database_id=uploads_db.id,
        name=dataset_name,
        table_name=table_name,
        columns=columns,
    )
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj
