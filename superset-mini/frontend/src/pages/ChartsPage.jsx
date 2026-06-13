import { useEffect, useState } from "react";
import { api } from "../api.js";
import ChartView from "../components/ChartView.jsx";

export default function ChartsPage({ onEdit }) {
  const [charts, setCharts] = useState([]);
  const [data, setData] = useState({}); // chartId -> result | {error}
  const [error, setError] = useState("");

  const load = () => api.listCharts().then(setCharts).catch((e) => setError(e.message));
  useEffect(() => { load(); }, []);

  // Lazily fetch each chart's data for thumbnails.
  useEffect(() => {
    charts.forEach((c) => {
      if (data[c.id]) return;
      api.chartData(c.id)
        .then((r) => setData((d) => ({ ...d, [c.id]: r })))
        .catch((e) => setData((d) => ({ ...d, [c.id]: { error: e.message } })));
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [charts]);

  async function remove(id) {
    if (!confirm("Delete chart?")) return;
    await api.deleteChart(id);
    load();
  }

  return (
    <div>
      <h2>Charts</h2>
      {error && <p className="error">{error}</p>}
      {charts.length === 0 && <p className="muted">No charts yet — build one in Explore.</p>}
      <div className="dashboard-grid">
        {charts.map((c) => (
          <div className="dashboard-tile" key={c.id}>
            <div className="row" style={{ justifyContent: "space-between" }}>
              <h4 style={{ margin: 0 }}>
                {c.name} <span className="muted small">{c.viz_type}</span>
              </h4>
              <div className="row">
                <button className="small" onClick={() => onEdit(c.id)}>Edit</button>
                <a className="small" href={api.chartCsvUrl(c.id)}>CSV</a>
                <button className="danger small" onClick={() => remove(c.id)}>Delete</button>
              </div>
            </div>
            {data[c.id]?.error ? (
              <p className="error">{data[c.id].error}</p>
            ) : data[c.id] ? (
              <ChartView vizType={c.viz_type} result={data[c.id]} height={220} />
            ) : (
              <p className="muted">Loading…</p>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}
