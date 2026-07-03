from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.core.config import settings
from app.modules.auth.router import router as auth_router
from app.modules.clientes.router import router as clientes_router
from app.modules.proveedores.router import router as proveedores_router
from app.modules.productos.router import router as productos_router
from app.modules.ventas.router import router as ventas_router
from app.modules.compras.router import router as compras_router
from app.modules.cobros.router import router as cobros_router

vantis = FastAPI(
    title=settings.APP_NAME,
    version=settings.APP_VERSION,
    description="VANTIS ERP — API SaaS para PYMEs chilenas",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    openapi_url="/api/openapi.json",
)

vantis.add_middleware(
    CORSMiddleware,
    allow_origins=settings.ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

vantis.include_router(auth_router)
vantis.include_router(clientes_router)
vantis.include_router(proveedores_router)
vantis.include_router(productos_router)
vantis.include_router(ventas_router)
vantis.include_router(compras_router)
vantis.include_router(cobros_router)


@vantis.get("/api/health", tags=["Sistema"])
async def health_check() -> dict:
    return {
        "status": "ok",
        "app": settings.APP_NAME,
        "version": settings.APP_VERSION,
    }


@vantis.get("/", tags=["Sistema"])
async def root() -> dict:
    return {"message": "VANTIS ERP API", "docs": "/api/docs"}