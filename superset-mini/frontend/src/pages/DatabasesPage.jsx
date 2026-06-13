import { useEffect, useState } from "react";
import { api } from "../api.js";

export default function DatabasesPage() {
  const [databases, setDatabases] = useState([]);
  const [name, setName] = useState("");
  const [uri, setUri] = useState("");
  const [error, setError] = useState("");
  const [tablesFor, setTablesFor] = useState(null);
  const [tables, setTables] = useState([]);

  const load = () => api.listDatabases().then(setDatabases).catch((e) => setError(e.message));
  useEffect(() => { load(); }, []);

  async function add(e) {
    e.preventDefault();
    setError("");
    try {
      await api.createDatabase({ name, sqlalchemy_uri: uri });
      setName(""); setUri(""); load();
    } catch (e) { setError(e.message); }
  }

  async function showTables(db) {
    setError(""); setTablesFor(db.id);
    try { setTables(await api.listTables(db.id)); }
    catch (e) { setError(e.message); setTables([]); }
  }

  async function remove(id) {
    if (!confirm("Delete this database connection?")) return;
    try { await api.deleteDatabase(id); load(); }
    catch (e) { setError(e.message); }
  }

  return (
    <div>
      <h2>Databases</h2>
      <p className="muted">
        Connect a data source via SQLAlchemy URI. Example:{" "}
        <code>sqlite:////absolute/path/to/file.db</code> or{" "}
        <code>postgresql+psycopg2://user:pass@host/dbname</code>
      </p>
      <form onSubmit={add} className="row">
        <input placeholder="Connection name" value={name} onChange={(e) => setName(e.target.value)} required />
        <input placeholder="SQLAlchemy URI" value={uri} onChange={(e) => setUri(e.target.value)}
               required style={{ flex: 2 }} />
        <button type="submit">Test & Add</button>
      </form>
      {error && <p className="error">{error}</p>}

      <table className="result-table">
        <thead><tr><th>ID</th><th>Name</th><th>URI</th><th>Actions</th></tr></thead>
        <tbody>
          {databases.map((d) => (
            <tr key={d.id}>
              <td>{d.id}</td>
              <td>{d.name}</td>
              <td className="mono">{d.sqlalchemy_uri}</td>
              <td>
                <button onClick={() => showTables(d)}>Tables</button>{" "}
                <button className="danger" onClick={() => remove(d.id)}>Delete</button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {tablesFor && (
        <div className="panel">
          <h4>Tables in database #{tablesFor}</h4>
          {tables.length ? (
            <ul>{tables.map((t) => <li key={t} className="mono">{t}</li>)}</ul>
          ) : <p className="muted">No tables found.</p>}
        </div>
      )}
    </div>
  );
}
