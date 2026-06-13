"""CRUD for charts + executing a chart's query to get data."""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from .. import engine as data_engine
from ..database import get_db
from ..models import Chart, Dataset
from ..query_builder import QueryBuildError, build_query
from ..schemas import ChartCreate, ChartOut, ChartUpdate, QueryResult, QuerySpec

router = APIRouter(prefix="/api/charts", tags=["charts"])


def _execute_chart(dataset: Dataset, spec: QuerySpec) -> QueryResult:
    try:
        sql, params = build_query(dataset, spec)
    except QueryBuildError as exc:
        raise HTTPException(400, str(exc))
    try:
        out = data_engine.run_sql(dataset.database.sqlalchemy_uri, sql, params)
    except Exception as exc:
        raise HTTPException(400, f"Query failed: {exc}")
    return QueryResult(
        columns=out["columns"], rows=out["rows"], sql=sql, row_count=len(out["rows"])
    )


@router.get("", response_model=list[ChartOut])
def list_charts(db: Session = Depends(get_db)):
    return db.query(Chart).order_by(Chart.id).all()


@router.get("/{chart_id}", response_model=ChartOut)
def get_chart(chart_id: int, db: Session = Depends(get_db)):
    obj = db.get(Chart, chart_id)
    if not obj:
        raise HTTPException(404, "Chart not found")
    return obj


@router.post("", response_model=ChartOut, status_code=201)
def create_chart(payload: ChartCreate, db: Session = Depends(get_db)):
    if not db.get(Dataset, payload.dataset_id):
        raise HTTPException(404, "Dataset not found")
    obj = Chart(
        name=payload.name,
        dataset_id=payload.dataset_id,
        viz_type=payload.viz_type,
        params=payload.params.model_dump(),
    )
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


@router.put("/{chart_id}", response_model=ChartOut)
def update_chart(chart_id: int, payload: ChartUpdate, db: Session = Depends(get_db)):
    obj = db.get(Chart, chart_id)
    if not obj:
        raise HTTPException(404, "Chart not found")
    if payload.name is not None:
        obj.name = payload.name
    if payload.viz_type is not None:
        obj.viz_type = payload.viz_type
    if payload.params is not None:
        obj.params = payload.params.model_dump()
    db.commit()
    db.refresh(obj)
    return obj


@router.delete("/{chart_id}", status_code=204)
def delete_chart(chart_id: int, db: Session = Depends(get_db)):
    obj = db.get(Chart, chart_id)
    if not obj:
        raise HTTPException(404, "Chart not found")
    db.delete(obj)
    db.commit()


@router.get("/{chart_id}/data", response_model=QueryResult)
def chart_data(chart_id: int, db: Session = Depends(get_db)):
    """Run the chart's saved query and return rows for rendering."""
    obj = db.get(Chart, chart_id)
    if not obj:
        raise HTTPException(404, "Chart not found")
    return _execute_chart(obj.dataset, QuerySpec(**obj.params))


@router.post("/explore", response_model=QueryResult)
def explore(payload: ChartCreate, db: Session = Depends(get_db)):
    """Run a query without saving — powers the live chart-builder preview."""
    dataset = db.get(Dataset, payload.dataset_id)
    if not dataset:
        raise HTTPException(404, "Dataset not found")
    return _execute_chart(dataset, payload.params)
