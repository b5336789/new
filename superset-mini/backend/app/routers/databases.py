"""CRUD + connection test for database connections."""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from .. import engine as data_engine
from ..database import get_db
from ..models import Database
from ..schemas import DatabaseCreate, DatabaseOut

router = APIRouter(prefix="/api/databases", tags=["databases"])


@router.get("", response_model=list[DatabaseOut])
def list_databases(db: Session = Depends(get_db)):
    return db.query(Database).order_by(Database.id).all()


@router.post("", response_model=DatabaseOut, status_code=201)
def create_database(payload: DatabaseCreate, db: Session = Depends(get_db)):
    if db.query(Database).filter(Database.name == payload.name).first():
        raise HTTPException(409, f"Database named '{payload.name}' already exists")
    try:
        data_engine.test_connection(payload.sqlalchemy_uri)
    except Exception as exc:  # fail loud: surface the real driver error
        raise HTTPException(400, f"Connection failed: {exc}")
    obj = Database(name=payload.name, sqlalchemy_uri=payload.sqlalchemy_uri)
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


@router.get("/{db_id}/tables", response_model=list[str])
def list_tables(db_id: int, db: Session = Depends(get_db)):
    obj = db.get(Database, db_id)
    if not obj:
        raise HTTPException(404, "Database not found")
    try:
        return data_engine.list_tables(obj.sqlalchemy_uri)
    except Exception as exc:
        raise HTTPException(400, f"Failed to list tables: {exc}")


@router.delete("/{db_id}", status_code=204)
def delete_database(db_id: int, db: Session = Depends(get_db)):
    obj = db.get(Database, db_id)
    if not obj:
        raise HTTPException(404, "Database not found")
    db.delete(obj)
    db.commit()
