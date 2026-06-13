"""FastAPI application entrypoint for Superset-mini."""
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from .config import ANTHROPIC_API_KEY, BASE_DIR
from .database import init_db
from .routers import charts, dashboards, databases, datasets, nl, sql, uploads
from .seed import seed_examples


@asynccontextmanager
async def lifespan(app: FastAPI):
    init_db()
    seed_examples()
    yield


app = FastAPI(title="Superset-mini", version="0.1.0", lifespan=lifespan)

# Dev CORS: the Vite frontend runs on a different origin.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(databases.router)
app.include_router(datasets.router)
app.include_router(sql.router)
app.include_router(charts.router)
app.include_router(dashboards.router)
app.include_router(uploads.router)
app.include_router(nl.router)


# Serve the technical-documentation site at /docs-site when present.
_DOCS_DIR = BASE_DIR.parent / "docs"
if (_DOCS_DIR / "site" / "index.html").exists():
    app.mount("/docs-site", StaticFiles(directory=_DOCS_DIR / "site", html=True),
              name="docs-site")


@app.get("/api/health")
def health():
    return {
        "status": "ok",
        # Surfaced so the UI can warn that text-to-chart will use the fallback.
        "text_to_chart": "claude" if ANTHROPIC_API_KEY else "heuristic_fallback",
    }
