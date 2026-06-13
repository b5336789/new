import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Scatter,
  ScatterChart,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import ResultTable from "./ResultTable.jsx";

const COLORS = [
  "#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f",
  "#edc948", "#b07aa1", "#ff9da7", "#9c755f", "#bab0ac",
];

// Convention from the query builder: dimension columns come first, then metrics.
function splitColumns(result) {
  const cols = result.columns || [];
  const rows = result.rows || [];
  const numeric = new Set();
  if (rows.length) {
    for (const c of cols) {
      if (typeof rows[0][c] === "number") numeric.add(c);
    }
  }
  const metricCols = cols.filter((c) => numeric.has(c));
  const dimCols = cols.filter((c) => !numeric.has(c));
  return { dimCols, metricCols };
}

export default function ChartView({ vizType, result, height = 320 }) {
  if (!result) return null;
  const rows = result.rows || [];

  if (vizType === "table") {
    return <ResultTable columns={result.columns} rows={rows} />;
  }
  if (rows.length === 0) return <p className="muted">No data to plot.</p>;

  const { dimCols, metricCols } = splitColumns(result);
  const xKey = dimCols[0] || result.columns[0];

  if (vizType === "big_number") {
    const metric = metricCols[0] || result.columns[0];
    const value = rows[0][metric];
    return (
      <div className="big-number">
        <div className="big-number-value">
          {typeof value === "number" ? value.toLocaleString() : String(value)}
        </div>
        <div className="big-number-label">{metric}</div>
      </div>
    );
  }

  if (vizType === "pie") {
    const nameKey = xKey;
    const valueKey = metricCols[0] || result.columns[1];
    return (
      <ResponsiveContainer width="100%" height={height}>
        <PieChart>
          <Pie data={rows} dataKey={valueKey} nameKey={nameKey} outerRadius={110} label>
            {rows.map((_, i) => (
              <Cell key={i} fill={COLORS[i % COLORS.length]} />
            ))}
          </Pie>
          <Tooltip />
          <Legend />
        </PieChart>
      </ResponsiveContainer>
    );
  }

  if (vizType === "scatter") {
    const xNum = metricCols[0] || result.columns[0];
    const yNum = metricCols[1] || metricCols[0] || result.columns[1];
    return (
      <ResponsiveContainer width="100%" height={height}>
        <ScatterChart>
          <CartesianGrid />
          <XAxis dataKey={xNum} name={xNum} type="number" />
          <YAxis dataKey={yNum} name={yNum} type="number" />
          <Tooltip cursor={{ strokeDasharray: "3 3" }} />
          <Scatter data={rows} fill={COLORS[0]} />
        </ScatterChart>
      </ResponsiveContainer>
    );
  }

  if (vizType === "line") {
    return (
      <ResponsiveContainer width="100%" height={height}>
        <LineChart data={rows}>
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey={xKey} />
          <YAxis />
          <Tooltip />
          <Legend />
          {metricCols.map((m, i) => (
            <Line key={m} type="monotone" dataKey={m} stroke={COLORS[i % COLORS.length]} />
          ))}
        </LineChart>
      </ResponsiveContainer>
    );
  }

  if (vizType === "area") {
    return (
      <ResponsiveContainer width="100%" height={height}>
        <AreaChart data={rows}>
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey={xKey} />
          <YAxis />
          <Tooltip />
          <Legend />
          {metricCols.map((m, i) => (
            <Area key={m} type="monotone" dataKey={m} stroke={COLORS[i % COLORS.length]}
                  fill={COLORS[i % COLORS.length]} fillOpacity={0.4} />
          ))}
        </AreaChart>
      </ResponsiveContainer>
    );
  }

  // default: bar
  return (
    <ResponsiveContainer width="100%" height={height}>
      <BarChart data={rows}>
        <CartesianGrid strokeDasharray="3 3" />
        <XAxis dataKey={xKey} />
        <YAxis />
        <Tooltip />
        <Legend />
        {metricCols.map((m, i) => (
          <Bar key={m} dataKey={m} fill={COLORS[i % COLORS.length]} />
        ))}
      </BarChart>
    </ResponsiveContainer>
  );
}
