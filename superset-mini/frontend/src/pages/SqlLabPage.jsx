import { useEffect, useState } from "react";
import { api } from "../api.js";
import ResultTable from "../components/ResultTable.jsx";

export default function SqlLabPage() {
  const [databases, setDatabases] = useState([]);
  const [dbId, setDbId] = useState("");
  const [sql, setSql] = useState("SELECT * FROM sales LIMIT 100");
  const [result, setResult] = useState(null);
  const [error, setError] = useState("");
  const [running, setRunning] = useState(false);

  useEffect(() => {
    api.listDatabases().then((dbs) => {
      setDatabases(dbs);
      if (dbs.length) setDbId(String(dbs[0].id));
    });
  }, []);

  async function run() {
    setError(""); setRunning(true);
    try {
      const r = await api.runSql({ database_id: Number(dbId), sql, row_limit: 1000 });
      setResult(r);
    } catch (e) { setError(e.message); setResult(null); }
    finally { setRunning(false); }
  }

  return (
    <div>
      <h2>SQL Lab</h2>
      <div className="row">
        <select value={dbId} onChange={(e) => setDbId(e.target.value)}>
          {databases.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
        </select>
        <button onClick={run} disabled={running || !dbId}>{running ? "Running…" : "Run"}</button>
      </div>
      <textarea className="sql-editor" value={sql} onChange={(e) => setSql(e.target.value)} rows={8} />
      {error && <p className="error">{error}</p>}
      {result && (
        <div>
          <p className="muted">{result.row_count} rows</p>
          <ResultTable columns={result.columns} rows={result.rows} />
        </div>
      )}
    </div>
  );
}
