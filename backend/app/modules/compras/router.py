from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.core.dependencies import get_current_user, require_modulo
from app.modules.compras.schemas import CompraCreate
from app.modules.compras import service

router = APIRouter(prefix="/api/compras", tags=["Compras"])


@router.get("")
async def listar(
    limite: int = 50,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("transacciones"))
) -> list:
    return await service.listar_compras(db, user["schema"], limite)


@router.get("/{compra_id}")
async def obtener(
    compra_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("transacciones"))
) -> dict:
    compra = await service.obtener_compra(db, user["schema"], compra_id)
    if not compra:
        raise HTTPException(status_code=404, detail="Compra no encontrada")
    return compra


@router.post("", status_code=201)
async def crear(
    datos: CompraCreate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("transacciones"))
) -> dict:
    resultado = await service.crear_compra(db, user["schema"], datos, user["email"])
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado
