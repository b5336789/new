"""CRUD for dashboards (grid arrangements of charts)."""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from ..database import get_db
from ..models import Chart, Dashboard
from ..schemas import DashboardCreate, DashboardOut, DashboardUpdate

router = APIRouter(prefix="/api/dashboards", tags=["dashboards"])


def _validate_layout(layout, db: Session):
    for item in layout:
        if not db.get(Chart, item.chart_id):
            raise HTTPException(400, f"Chart {item.chart_id} in layout does not exist")


@router.get("", response_model=list[DashboardOut])
def list_dashboards(db: Session = Depends(get_db)):
    return db.query(Dashboard).order_by(Dashboard.id).all()


@router.get("/{dash_id}", response_model=DashboardOut)
def get_dashboard(dash_id: int, db: Session = Depends(get_db)):
    obj = db.get(Dashboard, dash_id)
    if not obj:
        raise HTTPException(404, "Dashboard not found")
    return obj


@router.post("", response_model=DashboardOut, status_code=201)
def create_dashboard(payload: DashboardCreate, db: Session = Depends(get_db)):
    _validate_layout(payload.layout, db)
    obj = Dashboard(
        name=payload.name,
        layout=[item.model_dump() for item in payload.layout],
    )
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


@router.put("/{dash_id}", response_model=DashboardOut)
def update_dashboard(dash_id: int, payload: DashboardUpdate, db: Session = Depends(get_db)):
    obj = db.get(Dashboard, dash_id)
    if not obj:
        raise HTTPException(404, "Dashboard not found")
    if payload.name is not None:
        obj.name = payload.name
    if payload.layout is not None:
        _validate_layout(payload.layout, db)
        obj.layout = [item.model_dump() for item in payload.layout]
    db.commit()
    db.refresh(obj)
    return obj


@router.delete("/{dash_id}", status_code=204)
def delete_dashboard(dash_id: int, db: Session = Depends(get_db)):
    obj = db.get(Dashboard, dash_id)
    if not obj:
        raise HTTPException(404, "Dashboard not found")
    db.delete(obj)
    db.commit()
