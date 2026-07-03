from typing import Optional
from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.core.dependencies import require_modulo
from app.modules.clientes.schemas import ClienteCreate, ClienteUpdate
from app.modules.clientes import service

router = APIRouter(prefix="/api/clientes", tags=["Clientes"])


@router.get("")
async def listar(
    buscar: Optional[str] = None,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("clientes"))
) -> list:
    return await service.listar_clientes(db, user["schema"], buscar)


@router.get("/{cliente_id}")
async def obtener(
    cliente_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("clientes"))
) -> dict:
    cliente = await service.obtener_cliente(db, user["schema"], cliente_id)
    if not cliente:
        raise HTTPException(status_code=404, detail="Cliente no encontrado")
    return cliente


@router.post("", status_code=201)
async def crear(
    datos: ClienteCreate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("clientes"))
) -> dict:
    resultado = await service.crear_cliente(db, user["schema"], datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado


@router.put("/{cliente_id}")
async def actualizar(
    cliente_id: str,
    datos: ClienteUpdate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("clientes"))
) -> dict:
    resultado = await service.actualizar_cliente(db, user["schema"], cliente_id, datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado


@router.delete("/{cliente_id}")
async def eliminar(
    cliente_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("clientes"))
) -> dict:
    return await service.eliminar_cliente(db, user["schema"], cliente_id)
