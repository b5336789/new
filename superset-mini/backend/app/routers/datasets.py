"""CRUD for datasets (physical tables or virtual SQL sources)."""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from .. import engine as data_engine
from ..database import get_db
from ..models import Database, Dataset
from ..schemas import DatasetCreate, DatasetOut

router = APIRouter(prefix="/api/datasets", tags=["datasets"])


@router.get("", response_model=list[DatasetOut])
def list_datasets(db: Session = Depends(get_db)):
    return db.query(Dataset).order_by(Dataset.id).all()


@router.get("/{ds_id}", response_model=DatasetOut)
def get_dataset(ds_id: int, db: Session = Depends(get_db)):
    obj = db.get(Dataset, ds_id)
    if not obj:
        raise HTTPException(404, "Dataset not found")
    return obj


@router.post("", response_model=DatasetOut, status_code=201)
def create_dataset(payload: DatasetCreate, db: Session = Depends(get_db)):
    if bool(payload.table_name) == bool(payload.sql):
        raise HTTPException(400, "Provide exactly one of table_name or sql")
    database = db.get(Database, payload.database_id)
    if not database:
        raise HTTPException(404, "Database not found")
    try:
        columns = data_engine.inspect_columns(
            database.sqlalchemy_uri, table_name=payload.table_name, sql=payload.sql
        )
    except Exception as exc:
        raise HTTPException(400, f"Failed to inspect columns: {exc}")
    if not columns:
        raise HTTPException(400, "Dataset has no columns")
    obj = Dataset(
        database_id=payload.database_id,
        name=payload.name,
        table_name=payload.table_name,
        sql=payload.sql,
        columns=columns,
    )
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


@router.post("/{ds_id}/refresh", response_model=DatasetOut)
def refresh_columns(ds_id: int, db: Session = Depends(get_db)):
    obj = db.get(Dataset, ds_id)
    if not obj:
        raise HTTPException(404, "Dataset not found")
    obj.columns = data_engine.inspect_columns(
        obj.database.sqlalchemy_uri, table_name=obj.table_name, sql=obj.sql
    )
    db.commit()
    db.refresh(obj)
    return obj


@router.delete("/{ds_id}", status_code=204)
def delete_dataset(ds_id: int, db: Session = Depends(get_db)):
    obj = db.get(Dataset, ds_id)
    if not obj:
        raise HTTPException(404, "Dataset not found")
    db.delete(obj)
    db.commit()
