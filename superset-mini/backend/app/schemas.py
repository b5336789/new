"""Pydantic request/response schemas."""
from datetime import datetime
from typing import Any, Literal, Optional

from pydantic import BaseModel, ConfigDict, Field

# ---------------------------------------------------------------------------
# Query spec primitives (shared by charts and the query endpoint)
# ---------------------------------------------------------------------------

AggregateFn = Literal["SUM", "AVG", "MIN", "MAX", "COUNT", "COUNT_DISTINCT"]
FilterOp = Literal["=", "!=", ">", ">=", "<", "<=", "IN", "NOT IN", "LIKE", "IS NULL", "IS NOT NULL"]


class Metric(BaseModel):
    column: Optional[str] = None  # None allowed for COUNT(*)
    aggregate: AggregateFn = "SUM"
    label: Optional[str] = None


class Filter(BaseModel):
    column: str
    op: FilterOp = "="
    value: Any = None


class OrderBy(BaseModel):
    field: str  # references a metric label or a dimension column
    desc: bool = True


TimeGrain = Literal["day", "week", "month", "quarter", "year"]


class QuerySpec(BaseModel):
    dimensions: list[str] = Field(default_factory=list)
    metrics: list[Metric] = Field(default_factory=list)
    filters: list[Filter] = Field(default_factory=list)
    order_by: list[OrderBy] = Field(default_factory=list)
    row_limit: int = 1000
    # Temporal grouping (Superset-style "time grain"). When time_grain and
    # time_column are both set, the time column is truncated to the grain and
    # used as the leading dimension.
    time_grain: Optional[TimeGrain] = None
    time_column: Optional[str] = None


# ---------------------------------------------------------------------------
# Database
# ---------------------------------------------------------------------------

class DatabaseCreate(BaseModel):
    name: str
    sqlalchemy_uri: str


class DatabaseOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    name: str
    sqlalchemy_uri: str
    created_at: datetime


# ---------------------------------------------------------------------------
# Dataset
# ---------------------------------------------------------------------------

class DatasetCreate(BaseModel):
    database_id: int
    name: str
    table_name: Optional[str] = None
    sql: Optional[str] = None


class ColumnMeta(BaseModel):
    name: str
    type: str


class DatasetOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    database_id: int
    name: str
    table_name: Optional[str]
    sql: Optional[str]
    columns: list[ColumnMeta]
    created_at: datetime


# ---------------------------------------------------------------------------
# Chart
# ---------------------------------------------------------------------------

VizType = Literal["table", "bar", "line", "area", "pie", "scatter", "big_number"]


class ChartCreate(BaseModel):
    name: str
    dataset_id: int
    viz_type: VizType = "table"
    params: QuerySpec = Field(default_factory=QuerySpec)


class ChartUpdate(BaseModel):
    name: Optional[str] = None
    viz_type: Optional[VizType] = None
    params: Optional[QuerySpec] = None


class ChartOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    name: str
    dataset_id: int
    viz_type: str
    params: dict
    created_at: datetime


# ---------------------------------------------------------------------------
# Dashboard
# ---------------------------------------------------------------------------

class LayoutItem(BaseModel):
    chart_id: int
    x: int = 0
    y: int = 0
    w: int = 6
    h: int = 8


class DashboardCreate(BaseModel):
    name: str
    layout: list[LayoutItem] = Field(default_factory=list)


class DashboardUpdate(BaseModel):
    name: Optional[str] = None
    layout: Optional[list[LayoutItem]] = None


class DashboardOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    name: str
    layout: list[LayoutItem]
    created_at: datetime


# ---------------------------------------------------------------------------
# Query execution
# ---------------------------------------------------------------------------

class QueryResult(BaseModel):
    columns: list[str]
    rows: list[dict]
    sql: str
    row_count: int


class SqlRunRequest(BaseModel):
    database_id: int
    sql: str
    row_limit: int = 1000


# ---------------------------------------------------------------------------
# Text-to-chart
# ---------------------------------------------------------------------------

class NLChartRequest(BaseModel):
    dataset_id: int
    prompt: str


class NLChartResponse(BaseModel):
    viz_type: VizType
    params: QuerySpec
    explanation: str
    source: Literal["claude", "heuristic_fallback"]
    result: QueryResult
