import uuid
from datetime import datetime, timezone
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import text
from app.modules.cobros.schemas import CobroCreate, PagoCreate


async def registrar_cobro(db: AsyncSession, schema: str, datos: CobroCreate, usuario: str) -> dict:
    result = await db.execute(text(f"""
        SELECT id, cliente_id, saldo_pendiente, estado
        FROM "{schema}".cxc WHERE id = :id
    """), {"id": datos.cxc_id})
    cxc = result.fetchone()
    if not cxc:
        return {"ok": False, "error": "Documento CxC no encontrado"}

    c = cxc._mapping
    if c["estado"] == "PAGADO":
        return {"ok": False, "error": "Este documento ya esta completamente pagado"}

    saldo = float(c["saldo_pendiente"])
    if datos.monto > saldo:
        return {"ok": False, "error": f"Monto supera el saldo pendiente. Disponible: ${saldo:,.0f}"}

    ahora = datetime.now(timezone.utc)
    cobro_id = str(uuid.uuid4())
    saldo_nuevo = round(saldo - datos.monto, 2)
    estado_nuevo = "PAGADO" if saldo_nuevo <= 0 else "PARCIAL"

    await db.execute(text(f"""
        INSERT INTO "{schema}".cobros
        (id, cxc_id, cliente_id, monto, medio_pago, referencia, fecha, creado_por, creado_en)
        VALUES (:id, :cxc_id, :cliente_id, :monto, :medio_pago, :referencia, :fecha, :creado_por, :creado_en)
    """), {
        "id": cobro_id,
        "cxc_id": datos.cxc_id,
        "cliente_id": c["cliente_id"],
        "monto": datos.monto,
        "medio_pago": datos.medio_pago,
        "referencia": datos.referencia,
        "fecha": datetime.fromisoformat(datos.fecha),
        "creado_por": usuario,
        "creado_en": ahora,
    })

    await db.execute(text(f"""
        UPDATE "{schema}".cxc
        SET saldo_pendiente = :saldo, estado = :estado
        WHERE id = :id
    """), {"saldo": saldo_nuevo, "estado": estado_nuevo, "id": datos.cxc_id})

    await db.commit()
    return {
        "ok": True,
        "id": cobro_id,
        "saldo_anterior": saldo,
        "monto_cobrado": datos.monto,
        "saldo_nuevo": saldo_nuevo,
        "estado": estado_nuevo,
    }


async def registrar_pago(db: AsyncSession, schema: str, datos: PagoCreate, usuario: str) -> dict:
    result = await db.execute(text(f"""
        SELECT id, proveedor_id, saldo_pendiente, estado
        FROM "{schema}".cxp WHERE id = :id
    """), {"id": datos.cxp_id})
    cxp = result.fetchone()
    if not cxp:
        return {"ok": False, "error": "Documento CxP no encontrado"}

    p = cxp._mapping
    if p["estado"] == "PAGADO":
        return {"ok": False, "error": "Este documento ya esta completamente pagado"}

    saldo = float(p["saldo_pendiente"])
    if datos.monto > saldo:
        return {"ok": False, "error": f"Monto supera el saldo pendiente. Disponible: ${saldo:,.0f}"}

    ahora = datetime.now(timezone.utc)
    pago_id = str(uuid.uuid4())
    saldo_nuevo = round(saldo - datos.monto, 2)
    estado_nuevo = "PAGADO" if saldo_nuevo <= 0 else "PARCIAL"

    await db.execute(text(f"""
        INSERT INTO "{schema}".pagos
        (id, cxp_id, proveedor_id, monto, medio_pago, referencia, fecha, creado_por, creado_en)
        VALUES (:id, :cxp_id, :proveedor_id, :monto, :medio_pago, :referencia, :fecha, :creado_por, :creado_en)
    """), {
        "id": pago_id,
        "cxp_id": datos.cxp_id,
        "proveedor_id": p["proveedor_id"],
        "monto": datos.monto,
        "medio_pago": datos.medio_pago,
        "referencia": datos.referencia,
        "fecha": datetime.fromisoformat(datos.fecha),
        "creado_por": usuario,
        "creado_en": ahora,
    })

    await db.execute(text(f"""
        UPDATE "{schema}".cxp
        SET saldo_pendiente = :saldo, estado = :estado
        WHERE id = :id
    """), {"saldo": saldo_nuevo, "estado": estado_nuevo, "id": datos.cxp_id})

    await db.commit()
    return {
        "ok": True,
        "id": pago_id,
        "saldo_anterior": saldo,
        "monto_pagado": datos.monto,
        "saldo_nuevo": saldo_nuevo,
        "estado": estado_nuevo,
    }


async def listar_cxc(db: AsyncSession, schema: str) -> list:
    result = await db.execute(text(f"""
        SELECT cxc.id, cxc.venta_id, cxc.cliente_id, c.razon_social as cliente_nombre,
               cxc.monto_original, cxc.saldo_pendiente, cxc.fecha_emision,
               cxc.fecha_vencimiento, cxc.estado
        FROM "{schema}".cxc
        LEFT JOIN "{schema}".clientes c ON cxc.cliente_id = c.id
        WHERE cxc.estado != 'PAGADO'
        ORDER BY cxc.fecha_emision DESC
    """))
    return [dict(r._mapping) for r in result.fetchall()]


async def listar_cxp(db: AsyncSession, schema: str) -> list:
    result = await db.execute(text(f"""
        SELECT cxp.id, cxp.compra_id, cxp.proveedor_id, p.razon_social as proveedor_nombre,
               cxp.monto_original, cxp.saldo_pendiente, cxp.fecha_emision,
               cxp.fecha_vencimiento, cxp.estado
        FROM "{schema}".cxp
        LEFT JOIN "{schema}".proveedores p ON cxp.proveedor_id = p.id
        WHERE cxp.estado != 'PAGADO'
        ORDER BY cxp.fecha_emision DESC
    """))
    return [dict(r._mapping) for r in result.fetchall()]
