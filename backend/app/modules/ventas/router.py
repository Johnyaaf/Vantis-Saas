from fastapi import APIRouter, Depends, HTTPException, Request
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.core.dependencies import get_current_user, require_modulo
from app.modules.ventas.schemas import VentaCreate
from app.modules.ventas import service

router = APIRouter(prefix="/api/ventas", tags=["Ventas"])


@router.get("")
async def listar(
    limite: int = 50,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("transacciones"))
) -> list:
    return await service.listar_ventas(db, user["schema"], limite)


@router.get("/{venta_id}")
async def obtener(
    venta_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("transacciones"))
) -> dict:
    venta = await service.obtener_venta(db, user["schema"], venta_id)
    if not venta:
        raise HTTPException(status_code=404, detail="Venta no encontrada")
    return venta


@router.post("", status_code=201)
async def crear(
    datos: VentaCreate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("transacciones"))
) -> dict:
    resultado = await service.crear_venta(db, user["schema"], datos, user["email"])
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado
