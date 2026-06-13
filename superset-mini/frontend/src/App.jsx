import { useEffect, useState } from "react";
import { api } from "./api.js";
import DatabasesPage from "./pages/DatabasesPage.jsx";
import DatasetsPage from "./pages/DatasetsPage.jsx";
import SqlLabPage from "./pages/SqlLabPage.jsx";
import ExplorePage from "./pages/ExplorePage.jsx";
import DashboardsPage from "./pages/DashboardsPage.jsx";

const TABS = [
  ["dashboards", "Dashboards"],
  ["explore", "Explore"],
  ["ttc", "Text → Chart"],
  ["sql", "SQL Lab"],
  ["datasets", "Datasets"],
  ["databases", "Databases"],
];

export default function App() {
  const [tab, setTab] = useState("explore");
  const [health, setHealth] = useState(null);

  useEffect(() => { api.health().then(setHealth).catch(() => setHealth({ status: "down" })); }, []);

  return (
    <div className="app">
      <header>
        <h1>Superset-mini</h1>
        <nav>
          {TABS.map(([key, label]) => (
            <button key={key} className={tab === key ? "active" : ""} onClick={() => setTab(key)}>
              {label}
            </button>
          ))}
        </nav>
        {health && (
          <span className={`health ${health.status === "ok" ? "ok" : "bad"}`}>
            backend: {health.status}
            {health.text_to_chart && ` · text-to-chart: ${health.text_to_chart}`}
          </span>
        )}
      </header>
      <main>
        {tab === "dashboards" && <DashboardsPage />}
        {tab === "explore" && <ExplorePage key="explore" />}
        {tab === "ttc" && <ExplorePage key="ttc" ttcMode />}
        {tab === "sql" && <SqlLabPage />}
        {tab === "datasets" && <DatasetsPage />}
        {tab === "databases" && <DatabasesPage />}
      </main>
    </div>
  );
}
