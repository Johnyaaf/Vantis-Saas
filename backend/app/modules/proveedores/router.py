from fastapi import APIRouter, Depends, HTTPException
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.core.dependencies import get_current_user, require_modulo
from app.modules.proveedores.schemas import ProveedorCreate, ProveedorUpdate
from app.modules.proveedores import service

router = APIRouter(prefix="/api/proveedores", tags=["Proveedores"])


@router.get("")
async def listar(
    buscar: Optional[str] = None,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("proveedores"))
) -> list:
    return await service.listar_proveedores(db, user["schema"], buscar)


@router.get("/{proveedor_id}")
async def obtener(
    proveedor_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    proveedor = await service.obtener_proveedor(db, user["schema"], proveedor_id)
    if not proveedor:
        raise HTTPException(status_code=404, detail="Proveedor no encontrado")
    return proveedor


@router.post("", status_code=201)
async def crear(
    datos: ProveedorCreate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    resultado = await service.crear_proveedor(db, user["schema"], datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado


@router.put("/{proveedor_id}")
async def actualizar(
    proveedor_id: str,
    datos: ProveedorUpdate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    resultado = await service.actualizar_proveedor(db, user["schema"], proveedor_id, datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado


@router.delete("/{proveedor_id}")
async def eliminar(
    proveedor_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    return await service.eliminar_proveedor(db, user["schema"], proveedor_id)
