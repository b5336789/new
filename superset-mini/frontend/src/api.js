// Thin fetch wrapper. All calls go through the Vite proxy to the backend.

async function request(method, path, body, isForm = false) {
  const opts = { method, headers: {} };
  if (body !== undefined) {
    if (isForm) {
      opts.body = body; // FormData; let the browser set the boundary
    } else {
      opts.headers["Content-Type"] = "application/json";
      opts.body = JSON.stringify(body);
    }
  }
  const res = await fetch(`/api${path}`, opts);
  if (res.status === 204) return null;
  const text = await res.text();
  let data;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = text;
  }
  if (!res.ok) {
    const detail = data && data.detail ? data.detail : text || res.statusText;
    throw new Error(typeof detail === "string" ? detail : JSON.stringify(detail));
  }
  return data;
}

export const api = {
  health: () => request("GET", "/health"),

  // Databases
  listDatabases: () => request("GET", "/databases"),
  createDatabase: (b) => request("POST", "/databases", b),
  listTables: (id) => request("GET", `/databases/${id}/tables`),
  deleteDatabase: (id) => request("DELETE", `/databases/${id}`),

  // Datasets
  listDatasets: () => request("GET", "/datasets"),
  getDataset: (id) => request("GET", `/datasets/${id}`),
  createDataset: (b) => request("POST", "/datasets", b),
  deleteDataset: (id) => request("DELETE", `/datasets/${id}`),

  // SQL Lab
  runSql: (b) => request("POST", "/sql/run", b),

  // Charts
  listCharts: () => request("GET", "/charts"),
  getChart: (id) => request("GET", `/charts/${id}`),
  createChart: (b) => request("POST", "/charts", b),
  updateChart: (id, b) => request("PUT", `/charts/${id}`, b),
  deleteChart: (id) => request("DELETE", `/charts/${id}`),
  chartData: (id) => request("GET", `/charts/${id}/data`),
  explore: (b) => request("POST", "/charts/explore", b),

  // Dashboards
  listDashboards: () => request("GET", "/dashboards"),
  getDashboard: (id) => request("GET", `/dashboards/${id}`),
  createDashboard: (b) => request("POST", "/dashboards", b),
  updateDashboard: (id, b) => request("PUT", `/dashboards/${id}`, b),
  deleteDashboard: (id) => request("DELETE", `/dashboards/${id}`),

  // Upload
  upload: (file, datasetName) => {
    const fd = new FormData();
    fd.append("file", file);
    fd.append("dataset_name", datasetName);
    return request("POST", "/upload", fd, true);
  },

  // Text-to-chart
  nlChart: (b) => request("POST", "/nl/chart", b),

  // CSV export
  chartCsvUrl: (id) => `/api/charts/${id}/data.csv`,
  exploreCsv: async (b) => {
    const res = await fetch("/api/charts/explore.csv", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(b),
    });
    if (!res.ok) throw new Error((await res.text()) || res.statusText);
    return res.blob();
  },
};

// Trigger a browser download of a Blob.
export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
