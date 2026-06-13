"""Seed a sample data source + dataset so the app is useful on first launch.

Creates an `examples.db` SQLite file with a `sales` table and registers it as a
Database + Dataset in the metadata store (idempotent).
"""
import random
from datetime import date, timedelta

import pandas as pd

from . import engine as data_engine
from .config import EXAMPLES_DB_URI
from .database import SessionLocal
from .models import Database, Dataset

REGIONS = ["North", "South", "East", "West"]
CATEGORIES = ["Electronics", "Furniture", "Office Supplies", "Apparel"]
SEGMENTS = ["Consumer", "Corporate", "Home Office"]


def _build_sales_frame(n: int = 2000) -> pd.DataFrame:
    random.seed(42)
    start = date(2023, 1, 1)
    rows = []
    for i in range(n):
        d = start + timedelta(days=random.randint(0, 729))
        quantity = random.randint(1, 20)
        unit_price = round(random.uniform(5, 500), 2)
        sales = round(quantity * unit_price, 2)
        rows.append({
            "order_id": 1000 + i,
            "order_date": d.isoformat(),
            "region": random.choice(REGIONS),
            "category": random.choice(CATEGORIES),
            "segment": random.choice(SEGMENTS),
            "quantity": quantity,
            "unit_price": unit_price,
            "sales": sales,
            "profit": round(sales * random.uniform(-0.1, 0.4), 2),
        })
    return pd.DataFrame(rows)


def seed_examples():
    """Idempotently create the sample DB, Database row, and Dataset row."""
    db = SessionLocal()
    try:
        existing = db.query(Database).filter(
            Database.sqlalchemy_uri == EXAMPLES_DB_URI
        ).first()
        if existing:
            return  # already seeded

        eng = data_engine.get_engine(EXAMPLES_DB_URI)
        frame = _build_sales_frame()
        frame.to_sql("sales", eng, if_exists="replace", index=False)

        database = Database(name="Examples", sqlalchemy_uri=EXAMPLES_DB_URI)
        db.add(database)
        db.commit()
        db.refresh(database)

        columns = data_engine.inspect_columns(EXAMPLES_DB_URI, table_name="sales")
        dataset = Dataset(
            database_id=database.id,
            name="Sample Sales",
            table_name="sales",
            columns=columns,
        )
        db.add(dataset)
        db.commit()
    finally:
        db.close()
