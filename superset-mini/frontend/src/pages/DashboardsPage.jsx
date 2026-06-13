import { useEffect, useState } from "react";
import { api } from "../api.js";
import ChartView from "../components/ChartView.jsx";

export default function DashboardsPage() {
  const [dashboards, setDashboards] = useState([]);
  const [charts, setCharts] = useState([]);
  const [name, setName] = useState("");
  const [selected, setSelected] = useState([]);
  const [openId, setOpenId] = useState(null);
  const [error, setError] = useState("");

  const load = () => {
    api.listDashboards().then(setDashboards);
    api.listCharts().then(setCharts);
  };
  useEffect(() => { load(); }, []);

  function toggle(id) {
    setSelected((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
  }

  async function create(e) {
    e.preventDefault();
    setError("");
    // Simple auto-layout: 2 columns, each tile 6 wide / 9 tall.
    const layout = selected.map((chart_id, i) => ({
      chart_id, x: (i % 2) * 6, y: Math.floor(i / 2) * 9, w: 6, h: 9,
    }));
    try {
      await api.createDashboard({ name, layout });
      setName(""); setSelected([]); load();
    } catch (e) { setError(e.message); }
  }

  async function remove(id) {
    if (!confirm("Delete dashboard?")) return;
    await api.deleteDashboard(id);
    if (openId === id) setOpenId(null);
    load();
  }

  return (
    <div>
      <h2>Dashboards</h2>
      {error && <p className="error">{error}</p>}

      <div className="grid-2">
        <div className="panel">
          <h4>Create dashboard</h4>
          <form onSubmit={create}>
            <input placeholder="Dashboard name" value={name}
                   onChange={(e) => setName(e.target.value)} required />
            <p className="muted small">Select charts to include:</p>
            <div className="checklist">
              {charts.map((c) => (
                <label key={c.id}>
                  <input type="checkbox" checked={selected.includes(c.id)}
                         onChange={() => toggle(c.id)} /> {c.name} <span className="muted">({c.viz_type})</span>
                </label>
              ))}
              {charts.length === 0 && <p className="muted">No charts yet — build some in Explore.</p>}
            </div>
            <button type="submit" disabled={selected.length === 0}>Create</button>
          </form>
        </div>

        <div className="panel">
          <h4>Saved dashboards</h4>
          <ul className="link-list">
            {dashboards.map((d) => (
              <li key={d.id}>
                <button className="link" onClick={() => setOpenId(d.id)}>{d.name}</button>
                <span className="muted small"> ({d.layout.length} charts)</span>{" "}
                <button className="danger small" onClick={() => remove(d.id)}>delete</button>
              </li>
            ))}
          </ul>
        </div>
      </div>

      {openId && <DashboardView key={openId} id={openId} />}
    </div>
  );
}

function DashboardView({ id }) {
  const [dash, setDash] = useState(null);
  const [tiles, setTiles] = useState([]);

  useEffect(() => {
    (async () => {
      const d = await api.getDashboard(id);
      setDash(d);
      const loaded = await Promise.all(
        d.layout.map(async (item) => {
          try {
            const chart = await api.getChart(item.chart_id);
            const data = await api.chartData(item.chart_id);
            return { chart, data, error: null };
          } catch (e) {
            return { chart: { id: item.chart_id, name: `#${item.chart_id}` }, data: null, error: e.message };
          }
        })
      );
      setTiles(loaded);
    })();
  }, [id]);

  if (!dash) return null;
  return (
    <div className="panel">
      <h3>{dash.name}</h3>
      <div className="dashboard-grid">
        {tiles.map((t) => (
          <div className="dashboard-tile" key={t.chart.id}>
            <h4>{t.chart.name} <span className="muted small">{t.chart.viz_type}</span></h4>
            {t.error ? <p className="error">{t.error}</p>
              : <ChartView vizType={t.chart.viz_type} result={t.data} height={260} />}
          </div>
        ))}
      </div>
    </div>
  );
}
