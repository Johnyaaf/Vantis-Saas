import os

base = 'app/modules/proveedores'
os.makedirs(base, exist_ok=True)

# __init__.py
with open(f'{base}/__init__.py', 'w', encoding='utf-8') as f:
    f.write('')

# schemas.py
with open(f'{base}/schemas.py', 'w', encoding='utf-8') as f:
    f.write('''from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class ProveedorCreate(BaseModel):
    rut: str
    razon_social: str
    nombre_fantasia: Optional[str] = None
    giro: Optional[str] = None
    email: Optional[str] = None
    telefono: Optional[str] = None
    direccion: Optional[str] = None
    comuna: Optional[str] = None
    region: Optional[str] = None
    condicion_pago: str = "CONTADO"
    plazo_entrega_dias: int = 7
    moneda_habitual: str = "CLP"
    critico: bool = False


class ProveedorUpdate(BaseModel):
    razon_social: Optional[str] = None
    nombre_fantasia: Optional[str] = None
    giro: Optional[str] = None
    email: Optional[str] = None
    telefono: Optional[str] = None
    direccion: Optional[str] = None
    comuna: Optional[str] = None
    region: Optional[str] = None
    condicion_pago: Optional[str] = None
    plazo_entrega_dias: Optional[int] = None
    moneda_habitual: Optional[str] = None
    critico: Optional[bool] = None
    activo: Optional[bool] = None


class ProveedorResponse(BaseModel):
    id: str
    rut: str
    razon_social: str
    nombre_fantasia: Optional[str]
    giro: Optional[str]
    email: Optional[str]
    telefono: Optional[str]
    direccion: Optional[str]
    comuna: Optional[str]
    region: Optional[str]
    condicion_pago: str
    plazo_entrega_dias: int
    moneda_habitual: str
    critico: bool
    activo: bool
    creado_en: datetime

    class Config:
        from_attributes = True
''')

# service.py
with open(f'{base}/service.py', 'w', encoding='utf-8') as f:
    f.write('''import uuid
from datetime import datetime, timezone
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import text
from app.core.security import validate_rut_chile
from app.modules.proveedores.schemas import ProveedorCreate, ProveedorUpdate


async def listar_proveedores(db: AsyncSession, schema: str, buscar: Optional[str] = None) -> list:
    if buscar:
        result = await db.execute(text(f"""
            SELECT id, rut, razon_social, nombre_fantasia, giro, email, telefono,
                   direccion, comuna, region, condicion_pago, plazo_entrega_dias,
                   moneda_habitual, critico, activo, creado_en
            FROM "{schema}".proveedores
            WHERE activo = true AND (LOWER(razon_social) LIKE :buscar OR rut LIKE :buscar)
            ORDER BY razon_social
        """), {"buscar": f"%{buscar.lower()}%"})
    else:
        result = await db.execute(text(f"""
            SELECT id, rut, razon_social, nombre_fantasia, giro, email, telefono,
                   direccion, comuna, region, condicion_pago, plazo_entrega_dias,
                   moneda_habitual, critico, activo, creado_en
            FROM "{schema}".proveedores
            WHERE activo = true
            ORDER BY razon_social
        """))
    rows = result.fetchall()
    return [dict(row._mapping) for row in rows]


async def obtener_proveedor(db: AsyncSession, schema: str, proveedor_id: str) -> dict:
    result = await db.execute(text(f"""
        SELECT id, rut, razon_social, nombre_fantasia, giro, email, telefono,
               direccion, comuna, region, condicion_pago, plazo_entrega_dias,
               moneda_habitual, critico, activo, creado_en
        FROM "{schema}".proveedores WHERE id = :id
    """), {"id": proveedor_id})
    row = result.fetchone()
    return dict(row._mapping) if row else None


async def crear_proveedor(db: AsyncSession, schema: str, datos: ProveedorCreate) -> dict:
    if not validate_rut_chile(datos.rut):
        return {"ok": False, "error": "RUT invalido"}

    result = await db.execute(text(f"""
        SELECT id FROM "{schema}".proveedores WHERE rut = :rut
    """), {"rut": datos.rut})
    if result.fetchone():
        return {"ok": False, "error": "Ya existe un proveedor con ese RUT"}

    proveedor_id = str(uuid.uuid4())
    ahora = datetime.now(timezone.utc)

    await db.execute(text(f"""
        INSERT INTO "{schema}".proveedores
        (id, rut, razon_social, nombre_fantasia, giro, email, telefono,
         direccion, comuna, region, condicion_pago, plazo_entrega_dias,
         moneda_habitual, critico, activo, creado_en)
        VALUES (:id, :rut, :razon_social, :nombre_fantasia, :giro, :email, :telefono,
                :direccion, :comuna, :region, :condicion_pago, :plazo_entrega_dias,
                :moneda_habitual, :critico, true, :creado_en)
    """), {
        "id": proveedor_id,
        "rut": datos.rut,
        "razon_social": datos.razon_social,
        "nombre_fantasia": datos.nombre_fantasia,
        "giro": datos.giro,
        "email": datos.email,
        "telefono": datos.telefono,
        "direccion": datos.direccion,
        "comuna": datos.comuna,
        "region": datos.region,
        "condicion_pago": datos.condicion_pago,
        "plazo_entrega_dias": datos.plazo_entrega_dias,
        "moneda_habitual": datos.moneda_habitual,
        "critico": datos.critico,
        "creado_en": ahora,
    })
    await db.commit()
    return {"ok": True, "id": proveedor_id}


async def actualizar_proveedor(db: AsyncSession, schema: str, proveedor_id: str, datos: ProveedorUpdate) -> dict:
    campos = {k: v for k, v in datos.model_dump().items() if v is not None}
    if not campos:
        return {"ok": False, "error": "Sin campos para actualizar"}

    sets = ", ".join([f"{k} = :{k}" for k in campos])
    campos["id"] = proveedor_id

    await db.execute(text(f"""
        UPDATE "{schema}".proveedores SET {sets} WHERE id = :id
    """), campos)
    await db.commit()
    return {"ok": True}


async def eliminar_proveedor(db: AsyncSession, schema: str, proveedor_id: str) -> dict:
    await db.execute(text(f"""
        UPDATE "{schema}".proveedores SET activo = false WHERE id = :id
    """), {"id": proveedor_id})
    await db.commit()
    return {"ok": True}
''')

# router.py
with open(f'{base}/router.py', 'w', encoding='utf-8') as f:
    f.write('''from fastapi import APIRouter, Depends, HTTPException
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
''')

print("Modulo proveedores creado correctamente")

# Verificar
for archivo in ['__init__.py', 'schemas.py', 'service.py', 'router.py']:
    ruta = f'{base}/{archivo}'
    size = os.path.getsize(ruta)
    print(f"  {ruta}: {size} bytes")