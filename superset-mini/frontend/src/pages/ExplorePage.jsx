import { useEffect, useState } from "react";
import { api, downloadBlob } from "../api.js";
import ChartView from "../components/ChartView.jsx";

const VIZ_TYPES = ["table", "bar", "line", "area", "pie", "scatter", "big_number"];
const AGGS = ["SUM", "AVG", "MIN", "MAX", "COUNT", "COUNT_DISTINCT"];
const OPS = ["=", "!=", ">", ">=", "<", "<=", "IN", "NOT IN", "LIKE", "IS NULL", "IS NOT NULL"];
const GRAINS = ["", "day", "week", "month", "quarter", "year"];

export default function ExplorePage({ ttcMode, editChartId, onSaved }) {
  const [datasets, setDatasets] = useState([]);
  const [datasetId, setDatasetId] = useState("");
  const [dataset, setDataset] = useState(null);

  const [vizType, setVizType] = useState("bar");
  const [dimensions, setDimensions] = useState([]);
  const [metrics, setMetrics] = useState([{ column: "", aggregate: "SUM", label: "" }]);
  const [filters, setFilters] = useState([]);
  const [rowLimit, setRowLimit] = useState(1000);
  const [timeColumn, setTimeColumn] = useState("");
  const [timeGrain, setTimeGrain] = useState("");
  const [editingId, setEditingId] = useState(null);

  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [chartName, setChartName] = useState("");
  const [msg, setMsg] = useState("");

  // text-to-chart
  const [prompt, setPrompt] = useState("");
  const [ttcSource, setTtcSource] = useState("");
  const [ttcExplanation, setTtcExplanation] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => { api.listDatasets().then(setDatasets); }, []);

  useEffect(() => {
    if (!datasetId) { setDataset(null); return; }
    api.getDataset(datasetId).then(setDataset);
  }, [datasetId]);

  // Load an existing chart for editing when navigated here from the Charts page.
  useEffect(() => {
    if (!editChartId) return;
    (async () => {
      const c = await api.getChart(editChartId);
      const p = c.params || {};
      setEditingId(c.id);
      setChartName(c.name);
      setDatasetId(String(c.dataset_id));
      setVizType(c.viz_type);
      setDimensions(p.dimensions || []);
      setMetrics((p.metrics || [{ column: "", aggregate: "SUM", label: "" }]).map((m) => ({
        column: m.column || "", aggregate: m.aggregate, label: m.label || "",
      })));
      setFilters((p.filters || []).map((f) => ({ column: f.column, op: f.op, value: f.value ?? "" })));
      setRowLimit(p.row_limit || 1000);
      setTimeColumn(p.time_column || "");
      setTimeGrain(p.time_grain || "");
    })();
  }, [editChartId]);

  function buildSpec() {
    return {
      dimensions,
      metrics: metrics
        .filter((m) => m.aggregate)
        .map((m) => ({
          column: m.column || null,
          aggregate: m.aggregate,
          label: m.label || undefined,
        })),
      filters: filters
        .filter((f) => f.column)
        .map((f) => ({ column: f.column, op: f.op, value: parseValue(f.value) })),
      order_by: [],
      row_limit: Number(rowLimit) || 1000,
      time_column: timeGrain ? timeColumn || null : null,
      time_grain: timeGrain || null,
    };
  }

  async function runPreview() {
    setError(""); setMsg("");
    if (!datasetId) { setError("Select a dataset."); return; }
    try {
      const r = await api.explore({
        name: "preview", dataset_id: Number(datasetId), viz_type: vizType, params: buildSpec(),
      });
      setResult(r);
    } catch (e) { setError(e.message); setResult(null); }
  }

  async function save(asNew = false) {
    setError(""); setMsg("");
    if (!chartName) { setError("Give the chart a name."); return; }
    try {
      if (editingId && !asNew) {
        const c = await api.updateChart(editingId, {
          name: chartName, viz_type: vizType, params: buildSpec(),
        });
        setMsg(`Updated chart "${c.name}" (#${c.id}).`);
      } else {
        const c = await api.createChart({
          name: chartName, dataset_id: Number(datasetId), viz_type: vizType, params: buildSpec(),
        });
        setEditingId(c.id);
        setMsg(`Saved chart "${c.name}" (#${c.id}).`);
      }
      onSaved && onSaved();
    } catch (e) { setError(e.message); }
  }

  async function exportCsv() {
    setError("");
    try {
      const blob = await api.exploreCsv({
        name: "export", dataset_id: Number(datasetId), viz_type: vizType, params: buildSpec(),
      });
      downloadBlob(blob, `${chartName || "export"}.csv`);
    } catch (e) { setError(e.message); }
  }

  async function generate() {
    setError(""); setMsg(""); setTtcSource(""); setBusy(true);
    if (!datasetId) { setError("Select a dataset first."); setBusy(false); return; }
    try {
      const r = await api.nlChart({ dataset_id: Number(datasetId), prompt });
      // Apply the generated spec into the builder so it stays editable.
      setVizType(r.viz_type);
      setDimensions(r.params.dimensions || []);
      setMetrics(
        (r.params.metrics || []).map((m) => ({
          column: m.column || "", aggregate: m.aggregate, label: m.label || "",
        }))
      );
      setFilters(
        (r.params.filters || []).map((f) => ({ column: f.column, op: f.op, value: f.value ?? "" }))
      );
      setRowLimit(r.params.row_limit || 1000);
      setTimeColumn(r.params.time_column || "");
      setTimeGrain(r.params.time_grain || "");
      setResult(r.result);
      setTtcSource(r.source);
      setTtcExplanation(r.explanation);
    } catch (e) { setError(e.message); }
    finally { setBusy(false); }
  }

  const columns = dataset ? dataset.columns : [];

  return (
    <div>
      <h2>{ttcMode ? "Text → Chart" : "Explore (Chart Builder)"}</h2>

      <div className="row">
        <select value={datasetId} onChange={(e) => setDatasetId(e.target.value)}>
          <option value="">Select dataset…</option>
          {datasets.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
        </select>
        {dataset && <span className="muted small">{columns.length} columns</span>}
      </div>

      {/* Text-to-chart panel */}
      <div className="panel ttc">
        <h4>🪄 Describe the chart you want</h4>
        <div className="row">
          <input style={{ flex: 3 }} value={prompt} onChange={(e) => setPrompt(e.target.value)}
                 placeholder='e.g. "total sales by region" or "average profit per category as a pie chart"' />
          <button onClick={generate} disabled={busy || !datasetId}>
            {busy ? "Generating…" : "Generate"}
          </button>
        </div>
        {ttcSource && (
          <p className={ttcSource === "claude" ? "success" : "warn"}>
            Source: <b>{ttcSource}</b>
            {ttcSource === "heuristic_fallback" &&
              " — ANTHROPIC_API_KEY not set, used local keyword parser."}
            {ttcExplanation && <> · {ttcExplanation}</>}
          </p>
        )}
      </div>

      <div className="grid-2">
        <div className="panel">
          <h4>Query</h4>
          <label>Visualization</label>
          <select value={vizType} onChange={(e) => setVizType(e.target.value)}>
            {VIZ_TYPES.map((v) => <option key={v} value={v}>{v}</option>)}
          </select>

          <label>Dimensions (group by)</label>
          <select multiple value={dimensions} size={Math.min(5, Math.max(2, columns.length))}
                  onChange={(e) =>
                    setDimensions(Array.from(e.target.selectedOptions, (o) => o.value))}>
            {columns.map((c) => <option key={c.name} value={c.name}>{c.name} ({c.type})</option>)}
          </select>

          <label>Metrics</label>
          {metrics.map((m, i) => (
            <div className="row" key={i}>
              <select value={m.aggregate}
                      onChange={(e) => updateMetric(metrics, setMetrics, i, "aggregate", e.target.value)}>
                {AGGS.map((a) => <option key={a}>{a}</option>)}
              </select>
              <select value={m.column}
                      onChange={(e) => updateMetric(metrics, setMetrics, i, "column", e.target.value)}>
                <option value="">(none / *)</option>
                {columns.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
              </select>
              <input placeholder="label" value={m.label}
                     onChange={(e) => updateMetric(metrics, setMetrics, i, "label", e.target.value)} />
              <button className="danger" onClick={() => setMetrics(metrics.filter((_, j) => j !== i))}>×</button>
            </div>
          ))}
          <button onClick={() => setMetrics([...metrics, { column: "", aggregate: "SUM", label: "" }])}>
            + metric
          </button>

          <label>Filters</label>
          {filters.map((f, i) => (
            <div className="row" key={i}>
              <select value={f.column}
                      onChange={(e) => updateMetric(filters, setFilters, i, "column", e.target.value)}>
                <option value="">column…</option>
                {columns.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
              </select>
              <select value={f.op}
                      onChange={(e) => updateMetric(filters, setFilters, i, "op", e.target.value)}>
                {OPS.map((o) => <option key={o}>{o}</option>)}
              </select>
              <input placeholder="value" value={f.value}
                     onChange={(e) => updateMetric(filters, setFilters, i, "value", e.target.value)} />
              <button className="danger" onClick={() => setFilters(filters.filter((_, j) => j !== i))}>×</button>
            </div>
          ))}
          <button onClick={() => setFilters([...filters, { column: "", op: "=", value: "" }])}>
            + filter
          </button>

          <label>Time grain (temporal grouping)</label>
          <div className="row">
            <select value={timeColumn} onChange={(e) => setTimeColumn(e.target.value)}>
              <option value="">time column…</option>
              {columns.map((c) => <option key={c.name} value={c.name}>{c.name}</option>)}
            </select>
            <select value={timeGrain} onChange={(e) => setTimeGrain(e.target.value)}>
              {GRAINS.map((g) => <option key={g} value={g}>{g || "(none)"}</option>)}
            </select>
          </div>

          <label>Row limit</label>
          <input type="number" value={rowLimit} onChange={(e) => setRowLimit(e.target.value)} />

          <div className="row" style={{ marginTop: 12 }}>
            <button onClick={runPreview} disabled={!datasetId}>Run</button>
            <input placeholder="Chart name" value={chartName}
                   onChange={(e) => setChartName(e.target.value)} />
            <button onClick={() => save(false)} disabled={!datasetId}>
              {editingId ? "Update" : "Save chart"}
            </button>
            {editingId && (
              <button onClick={() => save(true)} disabled={!datasetId}>Save as new</button>
            )}
            <button onClick={exportCsv} disabled={!datasetId || !result}>Export CSV</button>
          </div>
          {editingId && <p className="muted small">Editing chart #{editingId}</p>}
        </div>

        <div className="panel">
          <h4>Preview</h4>
          {error && <p className="error">{error}</p>}
          {msg && <p className="success">{msg}</p>}
          {result ? (
            <>
              <ChartView vizType={vizType} result={result} />
              <details>
                <summary className="muted small">SQL ({result.row_count} rows)</summary>
                <pre className="mono small">{result.sql}</pre>
              </details>
            </>
          ) : <p className="muted">Run a query to preview.</p>}
        </div>
      </div>
    </div>
  );
}

function updateMetric(list, setList, i, key, value) {
  const next = list.slice();
  next[i] = { ...next[i], [key]: value };
  setList(next);
}

function parseValue(v) {
  if (v === "" || v === null || v === undefined) return null;
  if (typeof v === "string" && v.includes(",")) return v.split(",").map((s) => coerce(s.trim()));
  return coerce(v);
}

function coerce(v) {
  if (typeof v !== "string") return v;
  const n = Number(v);
  return v.trim() !== "" && !Number.isNaN(n) ? n : v;
}
