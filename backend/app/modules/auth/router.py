from fastapi import APIRouter, Depends, HTTPException
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.modules.auth.schemas import RegisterRequest, LoginRequest, TokenResponse, RefreshRequest
from app.modules.auth.service import registrar_tenant, login
from app.core.security import decode_token

router = APIRouter(prefix="/api/auth", tags=["Autenticación"])


@router.post("/register", status_code=201)
async def register(datos: RegisterRequest, db: AsyncSession = Depends(get_db)):
    resultado = await registrar_tenant(db, datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return {"mensaje": "Empresa registrada exitosamente", "data": resultado}


@router.post("/login", response_model=TokenResponse)
async def login_endpoint(datos: LoginRequest, db: AsyncSession = Depends(get_db)):
    resultado = await login(db, datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=401, detail=resultado["error"])
    return resultado


@router.post("/refresh")
async def refresh_token(datos: RefreshRequest):
    payload = decode_token(datos.refresh_token)
    if not payload or payload.get("type") != "refresh":
        raise HTTPException(status_code=401, detail="Token inválido o expirado")
    from app.core.security import create_access_token
    nuevo_token = create_access_token({
        "sub": payload["sub"],
        "tenant_id": payload["tenant_id"],
    })
    return {"access_token": nuevo_token, "token_type": "bearer"}


@router.get("/verify")
async def verify_token(token: str):
    payload = decode_token(token)
    if not payload:
        raise HTTPException(status_code=401, detail="Token inválido")
    return {"valido": True, "payload": payload}