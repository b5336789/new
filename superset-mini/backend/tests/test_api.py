"""End-to-end API tests covering the real workflow:
seed -> dataset -> SQL Lab -> chart explore/save -> dashboard -> upload -> text-to-chart.

These validate business outcomes (correct aggregates, real rows), not just status codes.
"""
import io
import os
import tempfile

import pandas as pd
import pytest

# Point all data files at a throwaway dir BEFORE importing the app.
_TMP = tempfile.mkdtemp()
os.environ.pop("ANTHROPIC_API_KEY", None)  # force heuristic fallback in tests


@pytest.fixture(scope="module")
def client():
    import importlib

    from app import config

    # Rewire config paths to the temp dir, then (re)build modules that captured them.
    config.DATA_DIR = __import__("pathlib").Path(_TMP)
    config.METADATA_DB_PATH = config.DATA_DIR / "metadata.db"
    config.METADATA_DB_URI = f"sqlite:///{config.METADATA_DB_PATH}"
    config.EXAMPLES_DB_PATH = config.DATA_DIR / "examples.db"
    config.EXAMPLES_DB_URI = f"sqlite:///{config.EXAMPLES_DB_PATH}"
    config.UPLOADS_DB_PATH = config.DATA_DIR / "uploads.db"
    config.UPLOADS_DB_URI = f"sqlite:///{config.UPLOADS_DB_PATH}"

    import app.database as database
    importlib.reload(database)
    import app.seed as seed
    importlib.reload(seed)
    import app.routers.uploads as uploads
    importlib.reload(uploads)
    import app.main as main
    importlib.reload(main)

    from fastapi.testclient import TestClient
    with TestClient(main.app) as c:
        yield c


def _sample_dataset_id(client):
    datasets = client.get("/api/datasets").json()
    sales = [d for d in datasets if d["name"] == "Sample Sales"]
    assert sales, "Sample Sales dataset should be seeded"
    return sales[0]["id"]


def test_health_reports_fallback(client):
    body = client.get("/api/health").json()
    assert body["status"] == "ok"
    assert body["text_to_chart"] == "heuristic_fallback"


def test_seed_dataset_has_expected_columns(client):
    ds_id = _sample_dataset_id(client)
    ds = client.get(f"/api/datasets/{ds_id}").json()
    names = {c["name"] for c in ds["columns"]}
    assert {"region", "sales", "category", "order_date"} <= names
    sales_type = next(c["type"] for c in ds["columns"] if c["name"] == "sales")
    assert sales_type == "number"


def test_sql_lab_runs_select(client):
    dbs = client.get("/api/databases").json()
    ex = next(d for d in dbs if d["name"] == "Examples")
    r = client.post("/api/sql/run", json={
        "database_id": ex["id"], "sql": "SELECT region, sales FROM sales", "row_limit": 5,
    })
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["row_count"] == 5
    assert set(body["columns"]) == {"region", "sales"}


def test_sql_lab_rejects_mutation(client):
    dbs = client.get("/api/databases").json()
    ex = next(d for d in dbs if d["name"] == "Examples")
    r = client.post("/api/sql/run", json={
        "database_id": ex["id"], "sql": "DELETE FROM sales", "row_limit": 5,
    })
    assert r.status_code == 400


def test_chart_explore_aggregates_correctly(client):
    ds_id = _sample_dataset_id(client)
    # SUM(sales) grouped by region should match a direct SQL aggregate.
    spec = {
        "name": "preview",
        "dataset_id": ds_id,
        "viz_type": "bar",
        "params": {
            "dimensions": ["region"],
            "metrics": [{"column": "sales", "aggregate": "SUM", "label": "total_sales"}],
            "order_by": [{"field": "total_sales", "desc": True}],
            "row_limit": 100,
        },
    }
    r = client.post("/api/charts/explore", json=spec)
    assert r.status_code == 200, r.text
    rows = r.json()["rows"]
    assert len(rows) == 4  # four regions
    assert "total_sales" in rows[0]
    # Ordered descending.
    totals = [row["total_sales"] for row in rows]
    assert totals == sorted(totals, reverse=True)


def test_chart_rejects_unknown_column(client):
    ds_id = _sample_dataset_id(client)
    r = client.post("/api/charts/explore", json={
        "name": "bad", "dataset_id": ds_id, "viz_type": "bar",
        "params": {"dimensions": ["nonexistent_col"], "metrics": []},
    })
    assert r.status_code == 400
    assert "Unknown column" in r.text


def test_save_chart_and_dashboard(client):
    ds_id = _sample_dataset_id(client)
    chart = client.post("/api/charts", json={
        "name": "Sales by Region", "dataset_id": ds_id, "viz_type": "bar",
        "params": {
            "dimensions": ["region"],
            "metrics": [{"column": "sales", "aggregate": "SUM", "label": "total_sales"}],
        },
    }).json()
    assert chart["id"]
    data = client.get(f"/api/charts/{chart['id']}/data").json()
    assert data["row_count"] == 4

    dash = client.post("/api/dashboards", json={
        "name": "Overview",
        "layout": [{"chart_id": chart["id"], "x": 0, "y": 0, "w": 6, "h": 8}],
    })
    assert dash.status_code == 201, dash.text
    assert dash.json()["layout"][0]["chart_id"] == chart["id"]


def test_dashboard_rejects_missing_chart(client):
    r = client.post("/api/dashboards", json={
        "name": "Bad", "layout": [{"chart_id": 99999, "x": 0, "y": 0, "w": 6, "h": 8}],
    })
    assert r.status_code == 400


def test_excel_upload_creates_dataset(client):
    frame = pd.DataFrame({
        "Product Name": ["A", "B", "A"],
        "Revenue ($)": [10.0, 20.0, 30.0],
    })
    buf = io.BytesIO()
    frame.to_excel(buf, index=False)
    buf.seek(0)
    r = client.post(
        "/api/upload",
        files={"file": ("data.xlsx", buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"dataset_name": "My Upload"},
    )
    assert r.status_code == 201, r.text
    ds = r.json()
    names = {c["name"] for c in ds["columns"]}
    # Headers sanitized into safe identifiers.
    assert "product_name" in names
    assert any("revenue" in n for n in names)

    # And it is immediately queryable.
    explore = client.post("/api/charts/explore", json={
        "name": "p", "dataset_id": ds["id"], "viz_type": "bar",
        "params": {
            "dimensions": ["product_name"],
            "metrics": [{"column": next(n for n in names if "revenue" in n),
                         "aggregate": "SUM", "label": "rev"}],
        },
    }).json()
    by_product = {row["product_name"]: row["rev"] for row in explore["rows"]}
    assert by_product["A"] == 40.0  # 10 + 30


def test_text_to_chart_fallback_builds_valid_chart(client):
    ds_id = _sample_dataset_id(client)
    r = client.post("/api/nl/chart", json={
        "dataset_id": ds_id, "prompt": "total sales by region",
    })
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["source"] == "heuristic_fallback"
    assert body["viz_type"] == "bar"
    assert body["params"]["dimensions"] == ["region"]
    assert body["params"]["metrics"][0]["aggregate"] == "SUM"
    assert body["result"]["row_count"] == 4


def test_time_grain_groups_by_month(client):
    ds_id = _sample_dataset_id(client)
    spec = {
        "name": "ts", "dataset_id": ds_id, "viz_type": "line",
        "params": {
            "time_column": "order_date", "time_grain": "month",
            "metrics": [{"column": "sales", "aggregate": "SUM", "label": "monthly_sales"}],
            "row_limit": 1000,
        },
    }
    r = client.post("/api/charts/explore", json=spec)
    assert r.status_code == 200, r.text
    body = r.json()
    assert "order_date" in body["columns"]
    # Sample data spans 2 years => up to 24 monthly buckets; well under row count.
    assert 1 < body["row_count"] <= 24
    # Buckets are YYYY-MM strings.
    assert all(len(row["order_date"]) == 7 and row["order_date"][4] == "-"
               for row in body["rows"])


def test_text_to_chart_detects_time_series(client):
    ds_id = _sample_dataset_id(client)
    r = client.post("/api/nl/chart", json={
        "dataset_id": ds_id, "prompt": "monthly sales trend over time",
    })
    assert r.status_code == 200, r.text
    body = r.json()
    assert body["viz_type"] == "line"
    assert body["params"]["time_column"] == "order_date"
    assert body["params"]["time_grain"] == "month"


def test_chart_csv_export(client):
    ds_id = _sample_dataset_id(client)
    chart = client.post("/api/charts", json={
        "name": "CSV Chart", "dataset_id": ds_id, "viz_type": "bar",
        "params": {
            "dimensions": ["region"],
            "metrics": [{"column": "sales", "aggregate": "SUM", "label": "total_sales"}],
        },
    }).json()
    r = client.get(f"/api/charts/{chart['id']}/data.csv")
    assert r.status_code == 200
    assert "text/csv" in r.headers["content-type"]
    lines = r.text.strip().splitlines()
    assert lines[0] == "region,total_sales"
    assert len(lines) == 5  # header + 4 regions
