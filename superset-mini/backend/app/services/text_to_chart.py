"""Text-to-chart: natural language -> (viz_type, QuerySpec).

Primary path uses the Claude API with a structured tool definition. When no
ANTHROPIC_API_KEY is configured we fall back to a deterministic keyword parser
and clearly tag the response `source="heuristic_fallback"` so the caller is
never misled into thinking the LLM ran (Fail Loud).
"""
from __future__ import annotations

import json

from ..config import ANTHROPIC_API_KEY, ANTHROPIC_MODEL
from ..models import Dataset
from ..schemas import QuerySpec

# The structured shape we ask Claude to fill in. Kept in sync with QuerySpec.
_CHART_TOOL = {
    "name": "create_chart",
    "description": "Define the visualization and query needed to answer the request.",
    "input_schema": {
        "type": "object",
        "properties": {
            "viz_type": {
                "type": "string",
                "enum": ["table", "bar", "line", "area", "pie", "scatter", "big_number"],
            },
            "dimensions": {
                "type": "array",
                "items": {"type": "string"},
                "description": "Columns to group by (categorical / time axes).",
            },
            "metrics": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "column": {"type": ["string", "null"]},
                        "aggregate": {
                            "type": "string",
                            "enum": ["SUM", "AVG", "MIN", "MAX", "COUNT", "COUNT_DISTINCT"],
                        },
                        "label": {"type": "string"},
                    },
                    "required": ["aggregate"],
                },
            },
            "filters": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "column": {"type": "string"},
                        "op": {"type": "string"},
                        "value": {},
                    },
                    "required": ["column", "op"],
                },
            },
            "order_by": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "field": {"type": "string"},
                        "desc": {"type": "boolean"},
                    },
                    "required": ["field"],
                },
            },
            "row_limit": {"type": "integer"},
            "time_grain": {
                "type": ["string", "null"],
                "enum": ["day", "week", "month", "quarter", "year", None],
                "description": "If grouping by time, the bucket size.",
            },
            "time_column": {
                "type": ["string", "null"],
                "description": "The datetime column to bucket when time_grain is set.",
            },
            "explanation": {
                "type": "string",
                "description": "One sentence explaining the chart in the user's language.",
            },
        },
        "required": ["viz_type", "explanation"],
    },
}


def _schema_text(dataset: Dataset) -> str:
    cols = ", ".join(f'{c["name"]} ({c["type"]})' for c in (dataset.columns or []))
    return f'Dataset "{dataset.name}" columns: {cols}'


def _spec_from_payload(payload: dict) -> tuple[str, QuerySpec, str]:
    viz_type = payload.get("viz_type", "table")
    explanation = payload.get("explanation", "")
    spec = QuerySpec(
        dimensions=payload.get("dimensions", []) or [],
        metrics=payload.get("metrics", []) or [],
        filters=payload.get("filters", []) or [],
        order_by=payload.get("order_by", []) or [],
        row_limit=payload.get("row_limit", 1000) or 1000,
        time_grain=payload.get("time_grain") or None,
        time_column=payload.get("time_column") or None,
    )
    return viz_type, spec, explanation


def _call_claude(dataset: Dataset, prompt: str) -> tuple[str, QuerySpec, str]:
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    system = (
        "You are a BI assistant that converts natural-language requests into a "
        "chart specification by calling the create_chart tool. Only reference "
        "columns that exist in the provided schema. Choose aggregates and a viz "
        "type that best answer the question. Reply only via the tool call."
    )
    message = client.messages.create(
        model=ANTHROPIC_MODEL,
        max_tokens=1024,
        system=system,
        tools=[_CHART_TOOL],
        tool_choice={"type": "tool", "name": "create_chart"},
        messages=[{
            "role": "user",
            "content": f"{_schema_text(dataset)}\n\nRequest: {prompt}",
        }],
    )
    for block in message.content:
        if block.type == "tool_use" and block.name == "create_chart":
            return _spec_from_payload(block.input)
    raise RuntimeError(f"Claude did not return a create_chart tool call: {message.content}")


# ---------------------------------------------------------------------------
# Deterministic fallback
# ---------------------------------------------------------------------------

_AGG_KEYWORDS = [
    ("count distinct", "COUNT_DISTINCT"),
    ("distinct", "COUNT_DISTINCT"),
    ("average", "AVG"),
    ("avg", "AVG"),
    ("mean", "AVG"),
    ("maximum", "MAX"),
    ("max", "MAX"),
    ("minimum", "MIN"),
    ("min", "MIN"),
    ("count", "COUNT"),
    ("number of", "COUNT"),
    ("sum", "SUM"),
    ("total", "SUM"),
]

_VIZ_KEYWORDS = [
    (("over time", "trend", "by month", "by year", "by day", "line"), "line"),
    (("pie", "share", "proportion", "percentage", "percent", "breakdown"), "pie"),
    (("scatter", "correlation", "relationship"), "scatter"),
    (("table", "list", "raw"), "table"),
    (("bar", "compare", "by ", "top", "per "), "bar"),
]


def _heuristic(dataset: Dataset, prompt: str) -> tuple[str, QuerySpec, str]:
    text = prompt.lower()
    columns = dataset.columns or []
    numeric = [c["name"] for c in columns if c["type"] == "number"]
    # Treat datetime-typed columns and date/time-named string columns as temporal.
    # (CSV/Excel-loaded dates often arrive as ISO strings; SQLite strftime handles them.)
    def looks_temporal(c):
        return c["type"] == "datetime" or (
            c["type"] == "string" and any(k in c["name"].lower() for k in ("date", "time", "_at", "year", "month"))
        )
    temporal = [c["name"] for c in columns if looks_temporal(c)]
    categorical = [
        c["name"] for c in columns
        if c["type"] in ("string", "boolean") and c["name"] not in temporal
    ]

    def mentioned(col: str) -> bool:
        return col.lower().replace("_", " ") in text or col.lower() in text

    # Aggregate
    aggregate = "SUM"
    for kw, agg in _AGG_KEYWORDS:
        if kw in text:
            aggregate = agg
            break

    # Metric column: a numeric column named in the prompt, else first numeric.
    metric_col = next((c for c in numeric if mentioned(c)), None) or (numeric[0] if numeric else None)
    if aggregate in ("COUNT", "COUNT_DISTINCT") and metric_col is None:
        metric = {"aggregate": "COUNT", "column": None, "label": "count"}
    elif metric_col is not None:
        metric = {"aggregate": aggregate, "column": metric_col,
                  "label": f"{aggregate.lower()}_{metric_col}"}
    else:
        metric = {"aggregate": "COUNT", "column": None, "label": "count"}

    # Detect a time-series request and pick a grain from keywords.
    time_words = ("over time", "trend", "by month", "by year", "by day",
                  "by week", "by quarter", "monthly", "yearly", "daily", "time")
    is_temporal = bool(temporal) and any(k in text for k in time_words)
    time_column = None
    time_grain = None
    if is_temporal:
        time_column = next((c for c in temporal if mentioned(c)), temporal[0])
        for grain in ("day", "week", "month", "quarter", "year"):
            if grain in text:
                time_grain = grain
                break
        time_grain = time_grain or "month"

    # Categorical dimension (used when not a pure time series).
    dimension = None
    if not is_temporal:
        dimension = next((c for c in categorical if mentioned(c)), None)
        if dimension is None and categorical:
            dimension = categorical[0]

    has_group = bool(time_column or dimension)

    # Viz type
    viz_type = "bar" if has_group else "big_number"
    if is_temporal:
        viz_type = "line"
    for keys, vt in _VIZ_KEYWORDS:
        if any(k in text for k in keys):
            viz_type = vt
            break
    if viz_type in ("bar", "line", "area", "pie", "scatter") and not has_group:
        viz_type = "big_number"

    dimensions = [dimension] if dimension and viz_type != "big_number" else []
    order_by = []
    if dimensions and viz_type in ("bar", "pie"):
        order_by = [{"field": metric["label"], "desc": True}]

    spec = QuerySpec(
        dimensions=dimensions,
        metrics=[metric],
        order_by=order_by,
        row_limit=1000,
        time_grain=time_grain if viz_type != "big_number" else None,
        time_column=time_column if viz_type != "big_number" else None,
    )
    grouped_by = time_column or (dimensions[0] if dimensions else None)
    explanation = (
        f"[heuristic] {viz_type} of {metric['label']}"
        + (f" by {grouped_by}" + (f" ({time_grain})" if time_grain else "") if grouped_by else "")
    )
    return viz_type, spec, explanation


def generate_chart_spec(dataset: Dataset, prompt: str) -> tuple[str, QuerySpec, str, str]:
    """Return (viz_type, QuerySpec, explanation, source)."""
    if ANTHROPIC_API_KEY:
        viz_type, spec, explanation = _call_claude(dataset, prompt)
        return viz_type, spec, explanation, "claude"
    viz_type, spec, explanation = _heuristic(dataset, prompt)
    return viz_type, spec, explanation, "heuristic_fallback"
