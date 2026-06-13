"""ORM models for the metadata store.

These mirror the core Superset object graph in miniature:
    Database --< Dataset --< Chart >-- Dashboard (via layout references)
"""
from datetime import datetime

from sqlalchemy import JSON, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import Mapped, mapped_column, relationship

from .database import Base


def _now() -> datetime:
    return datetime.utcnow()


class Database(Base):
    """A connection to a physical data source (SQLite/Postgres/...)."""

    __tablename__ = "databases"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(255), unique=True, nullable=False)
    sqlalchemy_uri: Mapped[str] = mapped_column(String(1024), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=_now)

    datasets: Mapped[list["Dataset"]] = relationship(
        back_populates="database", cascade="all, delete-orphan"
    )


class Dataset(Base):
    """A queryable table or virtual (SQL-defined) source within a Database."""

    __tablename__ = "datasets"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    database_id: Mapped[int] = mapped_column(ForeignKey("databases.id"), nullable=False)
    name: Mapped[str] = mapped_column(String(255), nullable=False)
    # Exactly one of table_name / sql is set. table_name = physical, sql = virtual.
    table_name: Mapped[str | None] = mapped_column(String(255), nullable=True)
    sql: Mapped[str | None] = mapped_column(Text, nullable=True)
    # Cached column metadata: [{"name": str, "type": str}, ...]
    columns: Mapped[list] = mapped_column(JSON, default=list)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=_now)

    database: Mapped["Database"] = relationship(back_populates="datasets")
    charts: Mapped[list["Chart"]] = relationship(
        back_populates="dataset", cascade="all, delete-orphan"
    )


class Chart(Base):
    """A saved visualization: a dataset + viz type + query params."""

    __tablename__ = "charts"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(255), nullable=False)
    dataset_id: Mapped[int] = mapped_column(ForeignKey("datasets.id"), nullable=False)
    viz_type: Mapped[str] = mapped_column(String(64), nullable=False, default="table")
    # Query spec: {"dimensions": [...], "metrics": [...], "filters": [...], ...}
    params: Mapped[dict] = mapped_column(JSON, default=dict)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=_now)

    dataset: Mapped["Dataset"] = relationship(back_populates="charts")


class Dashboard(Base):
    """A collection of charts arranged on a grid layout."""

    __tablename__ = "dashboards"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String(255), nullable=False)
    # Grid layout: [{"chart_id": int, "x": int, "y": int, "w": int, "h": int}, ...]
    layout: Mapped[list] = mapped_column(JSON, default=list)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=_now)
