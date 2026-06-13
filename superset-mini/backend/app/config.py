"""Application configuration.

All paths are resolved relative to the backend directory so the app runs the
same way regardless of the current working directory.
"""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent  # .../backend
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

# Metadata store: holds Superset-mini's own objects (databases, datasets, charts...).
METADATA_DB_PATH = DATA_DIR / "metadata.db"
METADATA_DB_URI = f"sqlite:///{METADATA_DB_PATH}"

# Default sample data source that ships with the app so charts work out of the box.
EXAMPLES_DB_PATH = DATA_DIR / "examples.db"
EXAMPLES_DB_URI = f"sqlite:///{EXAMPLES_DB_PATH}"

# SQLite DB that user-uploaded spreadsheets get written into.
UPLOADS_DB_PATH = DATA_DIR / "uploads.db"
UPLOADS_DB_URI = f"sqlite:///{UPLOADS_DB_PATH}"

# Hard cap on rows returned by any query, mirroring Superset's row_limit guardrail.
MAX_ROW_LIMIT = 50_000
DEFAULT_ROW_LIMIT = 1_000

# Text-to-chart (Claude) configuration. Read lazily at request time so the server
# can boot without a key; the NL endpoint fails loud if the key is missing.
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
ANTHROPIC_MODEL = os.environ.get("ANTHROPIC_MODEL", "claude-opus-4-8")
