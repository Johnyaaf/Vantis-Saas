from fastapi import APIRouter, Depends, HTTPException
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.core.dependencies import get_current_user, require_modulo
from app.modules.productos.schemas import ProductoCreate, ProductoUpdate
from app.modules.productos import service

router = APIRouter(prefix="/api/productos", tags=["Productos"])


@router.get("")
async def listar(
    buscar: Optional[str] = None,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("inventario"))
) -> list:
    return await service.listar_productos(db, user["schema"], buscar)


@router.get("/{producto_id}")
async def obtener(
    producto_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    producto = await service.obtener_producto(db, user["schema"], producto_id)
    if not producto:
        raise HTTPException(status_code=404, detail="Producto no encontrado")
    return producto


@router.post("", status_code=201)
async def crear(
    datos: ProductoCreate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("inventario"))
) -> dict:
    resultado = await service.crear_producto(db, user["schema"], datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado


@router.put("/{producto_id}")
async def actualizar(
    producto_id: str,
    datos: ProductoUpdate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    resultado = await service.actualizar_producto(db, user["schema"], producto_id, datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado


@router.delete("/{producto_id}")
async def eliminar(
    producto_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    return await service.eliminar_producto(db, user["schema"], producto_id)
