import os

base = 'app/modules/ventas'
os.makedirs(base, exist_ok=True)

with open(f'{base}/__init__.py', 'w', encoding='utf-8') as f:
    f.write('')

with open(f'{base}/schemas.py', 'w', encoding='utf-8') as f:
    f.write('''from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime


class LineaVenta(BaseModel):
    producto_id: str
    cantidad: float
    precio_unitario_neto: float
    descuento_pct: float = 0


class VentaCreate(BaseModel):
    tipo_documento: str = "BOLETA"
    cliente_id: str
    fecha: str
    condicion_pago: str = "CONTADO"
    lineas: List[LineaVenta]
    observaciones: Optional[str] = None
    documento_origen_id: Optional[str] = None


class VentaResponse(BaseModel):
    id: str
    numero_documento: str
    tipo_documento: str
    cliente_id: str
    fecha: datetime
    condicion_pago: str
    total_neto: float
    total_iva: float
    total_bruto: float
    estado: str
    creado_en: datetime

    class Config:
        from_attributes = True
''')

with open(f'{base}/service.py', 'w', encoding='utf-8') as f:
    f.write('''import uuid
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
        if p["maneja_stock"] and p["stock_actual"] < linea.cantidad:
            return {"ok": False, "error": f"Stock insuficiente para {p['nombre']}. Disponible: {p['stock_actual']}"}

        subtotal_neto = round(linea.cantidad * linea.precio_unitario_neto * (1 - linea.descuento_pct / 100), 2)
        subtotal_iva = round(subtotal_neto * tasa_iva, 2)
        subtotal_bruto = subtotal_neto + subtotal_iva
        costo_total = linea.cantidad * p["costo_unitario"]
        margen = subtotal_neto - costo_total

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
            "stock_antes": p["stock_actual"],
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
                "stock_resultante": linea["stock_antes"] - linea["cantidad"],
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
        WHERE v.estado != \'ANULADO\'
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
''')

with open(f'{base}/router.py', 'w', encoding='utf-8') as f:
    f.write('''from fastapi import APIRouter, Depends, HTTPException, Request
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.core.dependencies import get_current_user, require_modulo
from app.modules.ventas.schemas import VentaCreate
from app.modules.ventas import service

router = APIRouter(prefix="/api/ventas", tags=["Ventas"])


@router.get("")
async def listar(
    limite: int = 50,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("transacciones"))
) -> list:
    return await service.listar_ventas(db, user["schema"], limite)


@router.get("/{venta_id}")
async def obtener(
    venta_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("transacciones"))
) -> dict:
    venta = await service.obtener_venta(db, user["schema"], venta_id)
    if not venta:
        raise HTTPException(status_code=404, detail="Venta no encontrada")
    return venta


@router.post("", status_code=201)
async def crear(
    datos: VentaCreate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("transacciones"))
) -> dict:
    resultado = await service.crear_venta(db, user["schema"], datos, user["email"])
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado
''')

print("Modulo ventas creado correctamente")
for archivo in ['__init__.py', 'schemas.py', 'service.py', 'router.py']:
    ruta = f'{base}/{archivo}'
    size = os.path.getsize(ruta)
    print(f"  {ruta}: {size} bytes")