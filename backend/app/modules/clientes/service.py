import uuid
from datetime import datetime, timezone
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import text
from app.core.security import validate_rut_chile
from app.modules.clientes.schemas import ClienteCreate, ClienteUpdate


async def listar_clientes(db: AsyncSession, schema: str, buscar: Optional[str] = None) -> list:
    query = f"""
        SELECT id, rut, razon_social, nombre_fantasia, giro, email, telefono,
               direccion, comuna, region, condicion_pago, limite_credito, activo, creado_en
        FROM "{schema}".clientes
        WHERE activo = true
    """
    params = {}
    if buscar:
        query += " AND (LOWER(razon_social) LIKE LOWER(:buscar) OR rut LIKE :buscar)"
        params["buscar"] = f"%{buscar}%"
    query += " ORDER BY razon_social"

    result = await db.execute(text(query), params)
    rows = result.fetchall()
    return [dict(row._mapping) for row in rows]


async def obtener_cliente(db: AsyncSession, schema: str, cliente_id: str) -> Optional[dict]:
    result = await db.execute(text(f"""
        SELECT id, rut, razon_social, nombre_fantasia, giro, email, telefono,
               direccion, comuna, region, condicion_pago, limite_credito, activo, creado_en
        FROM "{schema}".clientes WHERE id = :id
    """), {"id": cliente_id})
    row = result.fetchone()
    return dict(row._mapping) if row else None


async def crear_cliente(db: AsyncSession, schema: str, datos: ClienteCreate) -> dict:
    if not validate_rut_chile(datos.rut):
        return {"ok": False, "error": "RUT invalido"}

    # Verificar duplicado
    result = await db.execute(text(f"""
        SELECT id FROM "{schema}".clientes WHERE rut = :rut
    """), {"rut": datos.rut})
    if result.fetchone():
        return {"ok": False, "error": "Ya existe un cliente con ese RUT"}

    cliente_id = str(uuid.uuid4())
    ahora = datetime.now(timezone.utc)

    await db.execute(text(f"""
        INSERT INTO "{schema}".clientes
        (id, rut, razon_social, nombre_fantasia, giro, email, telefono,
         direccion, comuna, region, condicion_pago, limite_credito, activo, creado_en)
        VALUES (:id, :rut, :razon_social, :nombre_fantasia, :giro, :email, :telefono,
                :direccion, :comuna, :region, :condicion_pago, :limite_credito, true, :creado_en)
    """), {
        "id": cliente_id,
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
        "limite_credito": datos.limite_credito,
        "creado_en": ahora,
    })
    await db.commit()
    return {"ok": True, "id": cliente_id}


async def actualizar_cliente(db: AsyncSession, schema: str, cliente_id: str, datos: ClienteUpdate) -> dict:
    campos = {k: v for k, v in datos.model_dump().items() if v is not None}
    if not campos:
        return {"ok": False, "error": "Sin campos para actualizar"}

    sets = ", ".join([f"{k} = :{k}" for k in campos])
    campos["id"] = cliente_id

    await db.execute(text(f"""
        UPDATE "{schema}".clientes SET {sets} WHERE id = :id
    """), campos)
    await db.commit()
    return {"ok": True}


async def eliminar_cliente(db: AsyncSession, schema: str, cliente_id: str) -> dict:
    await db.execute(text(f"""
        UPDATE "{schema}".clientes SET activo = false WHERE id = :id
    """), {"id": cliente_id})
    await db.commit()
    return {"ok": True}
