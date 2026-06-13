"""Metadata-store session management (SQLAlchemy 2.0)."""
from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, sessionmaker

from .config import METADATA_DB_URI

engine = create_engine(
    METADATA_DB_URI,
    connect_args={"check_same_thread": False},  # FastAPI uses a threadpool for sync deps
)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)


class Base(DeclarativeBase):
    pass


def get_db():
    """FastAPI dependency that yields a metadata-store session."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db():
    """Create metadata tables. Import models first so they register on Base."""
    from . import models  # noqa: F401

    Base.metadata.create_all(bind=engine)
