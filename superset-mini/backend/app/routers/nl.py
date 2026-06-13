"""Text-to-chart endpoint: NL prompt -> chart spec -> executed query."""
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.orm import Session

from .. import engine as data_engine
from ..database import get_db
from ..models import Dataset
from ..query_builder import QueryBuildError, build_query
from ..schemas import NLChartRequest, NLChartResponse, QueryResult
from ..services.text_to_chart import generate_chart_spec

router = APIRouter(prefix="/api/nl", tags=["text-to-chart"])


@router.post("/chart", response_model=NLChartResponse)
def nl_to_chart(payload: NLChartRequest, db: Session = Depends(get_db)):
    dataset = db.get(Dataset, payload.dataset_id)
    if not dataset:
        raise HTTPException(404, "Dataset not found")
    if not payload.prompt.strip():
        raise HTTPException(400, "Prompt is empty")

    try:
        viz_type, spec, explanation, source = generate_chart_spec(dataset, payload.prompt)
    except Exception as exc:  # fail loud on LLM/parse errors
        raise HTTPException(502, f"Text-to-chart generation failed: {exc}")

    # Validate + execute the generated spec against the real data.
    uri = dataset.database.sqlalchemy_uri
    try:
        sql, params = build_query(dataset, spec, dialect=data_engine.dialect_of(uri))
    except QueryBuildError as exc:
        raise HTTPException(
            422,
            f"Generated spec referenced invalid columns ({exc}). "
            f"Prompt: {payload.prompt!r}",
        )
    try:
        out = data_engine.run_sql(uri, sql, params)
    except Exception as exc:
        raise HTTPException(400, f"Generated query failed: {exc}")

    result = QueryResult(
        columns=out["columns"], rows=out["rows"], sql=sql, row_count=len(out["rows"])
    )
    return NLChartResponse(
        viz_type=viz_type, params=spec, explanation=explanation,
        source=source, result=result,
    )
