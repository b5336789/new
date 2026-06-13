# Text → Chart

Converts a natural-language request plus a dataset schema into an executable
`QuerySpec` + viz type. Implemented in `app/services/text_to_chart.py`,
exposed at `POST /api/nl/chart`.

## Flow

```
prompt + dataset.columns
        │
        ▼
generate_chart_spec()
        │  ANTHROPIC_API_KEY set?
        ├── yes ─► Claude API (tool use) ──► spec, source="claude"
        └── no  ─► heuristic parser      ──► spec, source="heuristic_fallback"
        │
        ▼
build_query(dataset, spec)   # validate columns, build SQL
        │
        ▼
engine.run_sql(...)          # execute against real data
        │
        ▼
{ viz_type, params, explanation, source, result }
```

The generated spec is **always validated and executed against the real data**
before returning. If the model invents a column, `build_query` raises and the
endpoint returns HTTP 422 with the offending prompt — we never return an
un-runnable chart silently.

## Claude path (primary)

Uses the Anthropic Messages API with **tool use** to force structured output.
A single tool, `create_chart`, has an `input_schema` mirroring `QuerySpec`
(dimensions, metrics, filters, order_by, row_limit, time_grain, time_column,
explanation). `tool_choice` forces the model to call it.

```python
client.messages.create(
    model=ANTHROPIC_MODEL,           # default: claude-opus-4-8
    tools=[_CHART_TOOL],
    tool_choice={"type": "tool", "name": "create_chart"},
    system="...convert NL → create_chart; only use columns in the schema...",
    messages=[{"role": "user",
               "content": f"{schema_text}\n\nRequest: {prompt}"}],
)
```

The tool input is parsed straight into a `QuerySpec`. Because the schema is sent
in the prompt and validated afterward, hallucinated columns are caught.

### Configuration

```bash
export ANTHROPIC_API_KEY=sk-ant-...
export ANTHROPIC_MODEL=claude-opus-4-8   # optional
```

`GET /api/health` reports `text_to_chart: "claude"` or `"heuristic_fallback"`
so the UI can warn the user which path is active.

## Heuristic fallback (no key)

A deterministic keyword parser so the feature is demonstrable without a key.
**It tags every response `source: "heuristic_fallback"`** — it never pretends
to be the LLM (a "fail loud" principle). It infers:

- **Aggregate** from words: *total/sum → SUM*, *average/avg → AVG*,
  *count/number of → COUNT*, *distinct → COUNT_DISTINCT*, *max/min*.
- **Metric column** = a numeric column named in the prompt, else the first.
- **Temporal intent** from *over time / trend / monthly / by month …*; picks a
  datetime (or date-named) column and a grain → `line` chart.
- **Dimension** = a categorical column named in the prompt, else the first.
- **Viz type** from *pie/share*, *scatter*, *table*, *bar/compare*, time → line.

### Example

`"monthly sales trend over time"` on the sample dataset →

```json
{
  "viz_type": "line",
  "params": {
    "metrics": [{"column": "sales", "aggregate": "SUM", "label": "sum_sales"}],
    "time_column": "order_date",
    "time_grain": "month"
  },
  "source": "heuristic_fallback"
}
```

## Why this design

- **Tool use over free-text JSON** — guarantees a parseable, schema-shaped
  result; no brittle JSON-from-prose extraction.
- **Validate-then-execute** — the LLM proposes, the deterministic engine
  disposes. The model can never produce an invalid or unsafe query that reaches
  the database, because `build_query` re-checks every identifier.
- **Transparent fallback** — useful offline / keyless, but always labeled.
