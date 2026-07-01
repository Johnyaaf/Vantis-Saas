from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.core.config import settings

vantis = FastAPI(
    title=settings.APP_NAME,
    version=settings.APP_VERSION,
    description="VANTIS ERP — API SaaS para PYMEs chilenas",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    openapi_url="/api/openapi.json",
)

# ── CORS ────────────────────────────────────────────────────────────
vantis.add_middleware(
    CORSMiddleware,
    allow_origins=settings.ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Health check ────────────────────────────────────────────────────
@vantis.get("/api/health", tags=["Sistema"])
async def health_check():
    return {
        "status": "ok",
        "app": settings.APP_NAME,
        "version": settings.APP_VERSION,
    }


# ── Root ────────────────────────────────────────────────────────────
@vantis.get("/", tags=["Sistema"])
async def root():
    return {"message": "VANTIS ERP API", "docs": "/api/docs"}