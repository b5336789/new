import { useEffect, useState } from "react";
import { api } from "../api.js";

export default function DatasetsPage() {
  const [datasets, setDatasets] = useState([]);
  const [databases, setDatabases] = useState([]);
  const [error, setError] = useState("");
  const [msg, setMsg] = useState("");

  // create-from-table / virtual sql state
  const [dbId, setDbId] = useState("");
  const [tables, setTables] = useState([]);
  const [mode, setMode] = useState("table"); // "table" | "sql"
  const [tableName, setTableName] = useState("");
  const [sql, setSql] = useState("");
  const [dsName, setDsName] = useState("");

  // upload state
  const [file, setFile] = useState(null);
  const [uploadName, setUploadName] = useState("");

  const load = () => {
    api.listDatasets().then(setDatasets).catch((e) => setError(e.message));
    api.listDatabases().then(setDatabases);
  };
  useEffect(() => { load(); }, []);

  async function onPickDb(id) {
    setDbId(id);
    setTableName("");
    if (id) {
      try { setTables(await api.listTables(id)); } catch { setTables([]); }
    }
  }

  async function createDataset(e) {
    e.preventDefault();
    setError(""); setMsg("");
    try {
      const body = {
        database_id: Number(dbId),
        name: dsName,
        table_name: mode === "table" ? tableName : null,
        sql: mode === "sql" ? sql : null,
      };
      await api.createDataset(body);
      setDsName(""); setTableName(""); setSql("");
      setMsg("Dataset created."); load();
    } catch (e) { setError(e.message); }
  }

  async function doUpload(e) {
    e.preventDefault();
    setError(""); setMsg("");
    if (!file) { setError("Choose a file first."); return; }
    try {
      const ds = await api.upload(file, uploadName || file.name.replace(/\.[^.]+$/, ""));
      setMsg(`Uploaded → dataset "${ds.name}" (#${ds.id}) with ${ds.columns.length} columns.`);
      setFile(null); setUploadName(""); load();
    } catch (e) { setError(e.message); }
  }

  async function remove(id) {
    if (!confirm("Delete this dataset?")) return;
    try { await api.deleteDataset(id); load(); }
    catch (e) { setError(e.message); }
  }

  return (
    <div>
      <h2>Datasets</h2>
      {error && <p className="error">{error}</p>}
      {msg && <p className="success">{msg}</p>}

      <div className="grid-2">
        <div className="panel">
          <h4>Upload Excel / CSV</h4>
          <p className="muted">Creates a table in the Uploads database and registers it as a dataset.</p>
          <form onSubmit={doUpload}>
            <input type="file" accept=".csv,.xlsx,.xls"
                   onChange={(e) => setFile(e.target.files[0])} />
            <input placeholder="Dataset name (optional)" value={uploadName}
                   onChange={(e) => setUploadName(e.target.value)} />
            <button type="submit">Upload</button>
          </form>
        </div>

        <div className="panel">
          <h4>Create from a database</h4>
          <form onSubmit={createDataset}>
            <select value={dbId} onChange={(e) => onPickDb(e.target.value)} required>
              <option value="">Select database…</option>
              {databases.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
            </select>
            <div className="row">
              <label><input type="radio" checked={mode === "table"} onChange={() => setMode("table")} /> Table</label>
              <label><input type="radio" checked={mode === "sql"} onChange={() => setMode("sql")} /> Virtual (SQL)</label>
            </div>
            {mode === "table" ? (
              <select value={tableName} onChange={(e) => setTableName(e.target.value)} required>
                <option value="">Select table…</option>
                {tables.map((t) => <option key={t} value={t}>{t}</option>)}
              </select>
            ) : (
              <textarea placeholder="SELECT ... FROM ..." value={sql}
                        onChange={(e) => setSql(e.target.value)} rows={4} />
            )}
            <input placeholder="Dataset name" value={dsName}
                   onChange={(e) => setDsName(e.target.value)} required />
            <button type="submit">Create dataset</button>
          </form>
        </div>
      </div>

      <h4>Existing datasets</h4>
      <table className="result-table">
        <thead><tr><th>ID</th><th>Name</th><th>Source</th><th>Columns</th><th></th></tr></thead>
        <tbody>
          {datasets.map((d) => (
            <tr key={d.id}>
              <td>{d.id}</td>
              <td>{d.name}</td>
              <td className="mono">{d.table_name || "(virtual SQL)"}</td>
              <td className="mono small">{d.columns.map((c) => c.name).join(", ")}</td>
              <td><button className="danger" onClick={() => remove(d.id)}>Delete</button></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
