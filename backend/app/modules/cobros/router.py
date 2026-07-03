from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.core.dependencies import get_current_user, require_modulo
from app.modules.cobros.schemas import CobroCreate, PagoCreate
from app.modules.cobros import service

router = APIRouter(prefix="/api", tags=["Cobranza y Pagos"])


@router.get("/cxc")
async def listar_cxc(
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("cobranza"))
) -> list:
    return await service.listar_cxc(db, user["schema"])


@router.post("/cxc/cobrar", status_code=201)
async def cobrar(
    datos: CobroCreate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("cobranza"))
) -> dict:
    resultado = await service.registrar_cobro(db, user["schema"], datos, user["email"])
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado


@router.get("/cxp")
async def listar_cxp(
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("proveedores"))
) -> list:
    return await service.listar_cxp(db, user["schema"])


@router.post("/cxp/pagar", status_code=201)
async def pagar(
    datos: PagoCreate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("proveedores"))
) -> dict:
    resultado = await service.registrar_pago(db, user["schema"], datos, user["email"])
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado
