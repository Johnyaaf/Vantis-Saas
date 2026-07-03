import uuid
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
