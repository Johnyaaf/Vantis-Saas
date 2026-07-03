import uuid
from datetime import datetime, timezone
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import text
from app.modules.ventas.schemas import VentaCreate


def _calcular_numero_documento(tipo: str, ahora: datetime) -> str:
    prefijo = {"BOLETA": "BOL", "FACTURA": "FAC", "NOTA_CREDITO": "NC", "NOTA_DEBITO": "ND"}.get(tipo, "DOC")
    return f"{prefijo}-{ahora.strftime('%Y%m%d%H%M%S')}"


async def crear_venta(db: AsyncSession, schema: str, datos: VentaCreate, usuario: str) -> dict:
    tipos_validos = ["BOLETA", "FACTURA", "NOTA_CREDITO", "NOTA_DEBITO"]
    if datos.tipo_documento not in tipos_validos:
        return {"ok": False, "error": f"Tipo de documento invalido. Use: {tipos_validos}"}

    if not datos.lineas:
        return {"ok": False, "error": "Debes agregar al menos una linea de producto"}

    result = await db.execute(text(f"""
        SELECT id FROM "{schema}".clientes WHERE id = :id AND activo = true
    """), {"id": datos.cliente_id})
    if not result.fetchone():
        return {"ok": False, "error": "Cliente no encontrado o inactivo"}

    tasa_iva = 0.19
    ahora = datetime.now(timezone.utc)
    venta_id = str(uuid.uuid4())
    numero_documento = _calcular_numero_documento(datos.tipo_documento, ahora)
    fecha_venta = datetime.fromisoformat(datos.fecha)

    total_neto = 0
    total_iva = 0
    total_bruto = 0
    lineas_procesadas = []

    for linea in datos.lineas:
        result = await db.execute(text(f"""
            SELECT id, nombre, costo_unitario, stock_actual, maneja_stock, precio_venta_neto
            FROM "{schema}".productos WHERE id = :id AND activo = true
        """), {"id": linea.producto_id})
        prod = result.fetchone()
        if not prod:
            return {"ok": False, "error": f"Producto {linea.producto_id} no encontrado"}

        p = prod._mapping
        if p["maneja_stock"] and float(p["stock_actual"]) < float(linea.cantidad):
            return {"ok": False, "error": f"Stock insuficiente para {p['nombre']}. Disponible: {p['stock_actual']}"}

        subtotal_neto = round(linea.cantidad * linea.precio_unitario_neto * (1 - linea.descuento_pct / 100), 2)
        subtotal_iva = round(subtotal_neto * tasa_iva, 2)
        subtotal_bruto = subtotal_neto + subtotal_iva
        costo_total = float(linea.cantidad) * float(p["costo_unitario"])
        margen = round(subtotal_neto - costo_total, 2)

        total_neto += subtotal_neto
        total_iva += subtotal_iva
        total_bruto += subtotal_bruto

        lineas_procesadas.append({
            "id": str(uuid.uuid4()),
            "venta_id": venta_id,
            "producto_id": linea.producto_id,
            "cantidad": linea.cantidad,
            "precio_unitario_neto": linea.precio_unitario_neto,
            "descuento_pct": linea.descuento_pct,
            "subtotal_neto": subtotal_neto,
            "subtotal_iva": subtotal_iva,
            "subtotal_bruto": subtotal_bruto,
            "costo_unitario": p["costo_unitario"],
            "costo_total": costo_total,
            "margen": margen,
            "stock_antes": float(p["stock_actual"]),
        })

    await db.execute(text(f"""
        INSERT INTO "{schema}".ventas
        (id, numero_documento, tipo_documento, cliente_id, fecha, condicion_pago,
         total_neto, total_iva, total_bruto, estado, observaciones,
         documento_origen_id, creado_por, creado_en)
        VALUES (:id, :numero_documento, :tipo_documento, :cliente_id, :fecha, :condicion_pago,
                :total_neto, :total_iva, :total_bruto, 'VIGENTE', :observaciones,
                :documento_origen_id, :creado_por, :creado_en)
    """), {
        "id": venta_id,
        "numero_documento": numero_documento,
        "tipo_documento": datos.tipo_documento,
        "cliente_id": datos.cliente_id,
        "fecha": fecha_venta,
        "condicion_pago": datos.condicion_pago,
        "total_neto": round(total_neto, 2),
        "total_iva": round(total_iva, 2),
        "total_bruto": round(total_bruto, 2),
        "observaciones": datos.observaciones,
        "documento_origen_id": datos.documento_origen_id,
        "creado_por": usuario,
        "creado_en": ahora,
    })

    for linea in lineas_procesadas:
        await db.execute(text(f"""
            INSERT INTO "{schema}".venta_lineas
            (id, venta_id, producto_id, cantidad, precio_unitario_neto, descuento_pct,
             subtotal_neto, subtotal_iva, subtotal_bruto, costo_unitario, costo_total, margen)
            VALUES (:id, :venta_id, :producto_id, :cantidad, :precio_unitario_neto, :descuento_pct,
                    :subtotal_neto, :subtotal_iva, :subtotal_bruto, :costo_unitario, :costo_total, :margen)
        """), linea)

        if lineas_procesadas[0]["stock_antes"] >= 0:
            await db.execute(text(f"""
                UPDATE "{schema}".productos
                SET stock_actual = stock_actual - :cantidad
                WHERE id = :id
            """), {"cantidad": linea["cantidad"], "id": linea["producto_id"]})

            mov_id = str(uuid.uuid4())
            await db.execute(text(f"""
                INSERT INTO "{schema}".movimientos_inventario
                (id, producto_id, tipo, cantidad, stock_resultante, referencia_id, creado_en)
                VALUES (:id, :producto_id, 'SALIDA_VENTA', :cantidad,
                        :stock_resultante, :referencia_id, :creado_en)
            """), {
                "id": mov_id,
                "producto_id": linea["producto_id"],
                "cantidad": linea["cantidad"],
                "stock_resultante": float(linea["stock_antes"]) - float(linea["cantidad"]),
                "referencia_id": venta_id,
                "creado_en": ahora,
            })

    if datos.condicion_pago == "CXC":
        cxc_id = str(uuid.uuid4())
        await db.execute(text(f"""
            INSERT INTO "{schema}".cxc
            (id, venta_id, cliente_id, monto_original, saldo_pendiente,
             fecha_emision, fecha_vencimiento, estado, creado_en)
            VALUES (:id, :venta_id, :cliente_id, :monto, :monto,
                    :fecha_emision, :fecha_emision, 'PENDIENTE', :creado_en)
        """), {
            "id": cxc_id,
            "venta_id": venta_id,
            "cliente_id": datos.cliente_id,
            "monto": round(total_bruto, 2),
            "fecha_emision": fecha_venta,
            "creado_en": ahora,
        })

    await db.commit()
    return {
        "ok": True,
        "id": venta_id,
        "numero_documento": numero_documento,
        "total_neto": round(total_neto, 2),
        "total_iva": round(total_iva, 2),
        "total_bruto": round(total_bruto, 2),
        "lineas": len(lineas_procesadas),
        "genero_cxc": datos.condicion_pago == "CXC",
    }


async def listar_ventas(db: AsyncSession, schema: str, limite: int = 50) -> list:
    result = await db.execute(text(f"""
        SELECT v.id, v.numero_documento, v.tipo_documento, v.cliente_id,
               c.razon_social as cliente_nombre, v.fecha, v.condicion_pago,
               v.total_neto, v.total_iva, v.total_bruto, v.estado, v.creado_en
        FROM "{schema}".ventas v
        LEFT JOIN "{schema}".clientes c ON v.cliente_id = c.id
        WHERE v.estado != 'ANULADO'
        ORDER BY v.creado_en DESC
        LIMIT :limite
    """), {"limite": limite})
    rows = result.fetchall()
    return [dict(row._mapping) for row in rows]


async def obtener_venta(db: AsyncSession, schema: str, venta_id: str) -> dict:
    result = await db.execute(text(f"""
        SELECT v.*, c.razon_social as cliente_nombre
        FROM "{schema}".ventas v
        LEFT JOIN "{schema}".clientes c ON v.cliente_id = c.id
        WHERE v.id = :id
    """), {"id": venta_id})
    venta = result.fetchone()
    if not venta:
        return None

    result2 = await db.execute(text(f"""
        SELECT vl.*, p.nombre as producto_nombre, p.sku
        FROM "{schema}".venta_lineas vl
        LEFT JOIN "{schema}".productos p ON vl.producto_id = p.id
        WHERE vl.venta_id = :venta_id
    """), {"venta_id": venta_id})
    lineas = [dict(r._mapping) for r in result2.fetchall()]

    data = dict(venta._mapping)
    data["lineas"] = lineas
    return data
