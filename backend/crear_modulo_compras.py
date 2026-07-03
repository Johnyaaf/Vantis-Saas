import os

base = 'app/modules/compras'
os.makedirs(base, exist_ok=True)

with open(f'{base}/__init__.py', 'w', encoding='utf-8') as f:
    f.write('')

with open(f'{base}/schemas.py', 'w', encoding='utf-8') as f:
    f.write('''from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime


class LineaCompra(BaseModel):
    producto_id: str
    cantidad: float
    costo_unitario: float
    descuento_pct: float = 0


class CompraCreate(BaseModel):
    tipo_documento: str = "FACTURA"
    proveedor_id: str
    numero_documento: str
    fecha: str
    condicion_pago: str = "CONTADO"
    lineas: List[LineaCompra]
    observaciones: Optional[str] = None


class CompraResponse(BaseModel):
    id: str
    tipo_documento: str
    proveedor_id: str
    numero_documento: str
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
from app.modules.compras.schemas import CompraCreate


async def crear_compra(db: AsyncSession, schema: str, datos: CompraCreate, usuario: str) -> dict:
    tipos_validos = ["FACTURA", "BOLETA", "NOTA_CREDITO", "NOTA_DEBITO"]
    if datos.tipo_documento not in tipos_validos:
        return {"ok": False, "error": f"Tipo de documento invalido. Use: {tipos_validos}"}

    if not datos.lineas:
        return {"ok": False, "error": "Debes agregar al menos una linea de producto"}

    if not datos.numero_documento or not datos.numero_documento.strip():
        return {"ok": False, "error": "El numero de documento es obligatorio"}

    result = await db.execute(text(f"""
        SELECT id FROM "{schema}".proveedores WHERE id = :id AND activo = true
    """), {"id": datos.proveedor_id})
    if not result.fetchone():
        return {"ok": False, "error": "Proveedor no encontrado o inactivo"}

    tasa_iva = 0.19
    ahora = datetime.now(timezone.utc)
    compra_id = str(uuid.uuid4())
    fecha_compra = datetime.fromisoformat(datos.fecha)

    total_neto = 0
    total_iva = 0
    total_bruto = 0
    lineas_procesadas = []

    for linea in datos.lineas:
        result = await db.execute(text(f"""
            SELECT id, nombre, costo_unitario, stock_actual, maneja_stock
            FROM "{schema}".productos WHERE id = :id AND activo = true
        """), {"id": linea.producto_id})
        prod = result.fetchone()
        if not prod:
            return {"ok": False, "error": f"Producto {linea.producto_id} no encontrado"}

        p = prod._mapping
        subtotal_neto = round(float(linea.cantidad) * float(linea.costo_unitario) * (1 - linea.descuento_pct / 100), 2)
        subtotal_iva = round(subtotal_neto * tasa_iva, 2)
        subtotal_bruto = subtotal_neto + subtotal_iva

        total_neto += subtotal_neto
        total_iva += subtotal_iva
        total_bruto += subtotal_bruto

        stock_actual = float(p["stock_actual"])
        costo_actual = float(p["costo_unitario"])
        stock_nuevo = stock_actual + float(linea.cantidad)
        if stock_nuevo > 0:
            costo_cpp = round((stock_actual * costo_actual + float(linea.cantidad) * float(linea.costo_unitario)) / stock_nuevo, 2)
        else:
            costo_cpp = float(linea.costo_unitario)

        lineas_procesadas.append({
            "id": str(uuid.uuid4()),
            "compra_id": compra_id,
            "producto_id": linea.producto_id,
            "cantidad": float(linea.cantidad),
            "costo_unitario": float(linea.costo_unitario),
            "descuento_pct": linea.descuento_pct,
            "subtotal_neto": subtotal_neto,
            "subtotal_iva": subtotal_iva,
            "subtotal_bruto": subtotal_bruto,
            "stock_antes": stock_actual,
            "stock_nuevo": stock_nuevo,
            "costo_cpp": costo_cpp,
            "maneja_stock": bool(p["maneja_stock"]),
        })

    await db.execute(text(f"""
        INSERT INTO "{schema}".compras
        (id, tipo_documento, proveedor_id, numero_documento, fecha, condicion_pago,
         total_neto, total_iva, total_bruto, estado, observaciones, creado_por, creado_en)
        VALUES (:id, :tipo_documento, :proveedor_id, :numero_documento, :fecha, :condicion_pago,
                :total_neto, :total_iva, :total_bruto, \'RECIBIDA\', :observaciones, :creado_por, :creado_en)
    """), {
        "id": compra_id,
        "tipo_documento": datos.tipo_documento,
        "proveedor_id": datos.proveedor_id,
        "numero_documento": datos.numero_documento,
        "fecha": fecha_compra,
        "condicion_pago": datos.condicion_pago,
        "total_neto": round(total_neto, 2),
        "total_iva": round(total_iva, 2),
        "total_bruto": round(total_bruto, 2),
        "observaciones": datos.observaciones,
        "creado_por": usuario,
        "creado_en": ahora,
    })

    for linea in lineas_procesadas:
        await db.execute(text(f"""
            INSERT INTO "{schema}".compra_lineas
            (id, compra_id, producto_id, cantidad, costo_unitario, descuento_pct,
             subtotal_neto, subtotal_iva, subtotal_bruto)
            VALUES (:id, :compra_id, :producto_id, :cantidad, :costo_unitario, :descuento_pct,
                    :subtotal_neto, :subtotal_iva, :subtotal_bruto)
        """), linea)

        if linea["maneja_stock"]:
            await db.execute(text(f"""
                UPDATE "{schema}".productos
                SET stock_actual = :stock_nuevo, costo_unitario = :costo_cpp
                WHERE id = :id
            """), {"stock_nuevo": linea["stock_nuevo"], "costo_cpp": linea["costo_cpp"], "id": linea["producto_id"]})

            await db.execute(text(f"""
                INSERT INTO "{schema}".movimientos_inventario
                (id, producto_id, tipo, cantidad, stock_resultante, referencia_id, creado_en)
                VALUES (:id, :producto_id, \'ENTRADA_COMPRA\', :cantidad, :stock_resultante, :referencia_id, :creado_en)
            """), {
                "id": str(uuid.uuid4()),
                "producto_id": linea["producto_id"],
                "cantidad": linea["cantidad"],
                "stock_resultante": linea["stock_nuevo"],
                "referencia_id": compra_id,
                "creado_en": ahora,
            })

    if datos.condicion_pago == "CXP":
        await db.execute(text(f"""
            INSERT INTO "{schema}".cxp
            (id, compra_id, proveedor_id, monto_original, saldo_pendiente,
             fecha_emision, fecha_vencimiento, estado, creado_en)
            VALUES (:id, :compra_id, :proveedor_id, :monto, :monto,
                    :fecha_emision, :fecha_emision, \'PENDIENTE\', :creado_en)
        """), {
            "id": str(uuid.uuid4()),
            "compra_id": compra_id,
            "proveedor_id": datos.proveedor_id,
            "monto": round(total_bruto, 2),
            "fecha_emision": fecha_compra,
            "creado_en": ahora,
        })

    await db.commit()
    return {
        "ok": True,
        "id": compra_id,
        "numero_documento": datos.numero_documento,
        "total_neto": round(total_neto, 2),
        "total_iva": round(total_iva, 2),
        "total_bruto": round(total_bruto, 2),
        "lineas": len(lineas_procesadas),
        "genero_cxp": datos.condicion_pago == "CXP",
    }


async def listar_compras(db: AsyncSession, schema: str, limite: int = 50) -> list:
    result = await db.execute(text(f"""
        SELECT c.id, c.tipo_documento, c.proveedor_id,
               p.razon_social as proveedor_nombre, c.numero_documento,
               c.fecha, c.condicion_pago, c.total_neto, c.total_iva,
               c.total_bruto, c.estado, c.creado_en
        FROM "{schema}".compras c
        LEFT JOIN "{schema}".proveedores p ON c.proveedor_id = p.id
        WHERE c.estado != \'ANULADA\'
        ORDER BY c.creado_en DESC
        LIMIT :limite
    """), {"limite": limite})
    rows = result.fetchall()
    return [dict(row._mapping) for row in rows]


async def obtener_compra(db: AsyncSession, schema: str, compra_id: str) -> dict:
    result = await db.execute(text(f"""
        SELECT c.*, p.razon_social as proveedor_nombre
        FROM "{schema}".compras c
        LEFT JOIN "{schema}".proveedores p ON c.proveedor_id = p.id
        WHERE c.id = :id
    """), {"id": compra_id})
    compra = result.fetchone()
    if not compra:
        return None

    result2 = await db.execute(text(f"""
        SELECT cl.*, pr.nombre as producto_nombre, pr.sku
        FROM "{schema}".compra_lineas cl
        LEFT JOIN "{schema}".productos pr ON cl.producto_id = pr.id
        WHERE cl.compra_id = :compra_id
    """), {"compra_id": compra_id})
    lineas = [dict(r._mapping) for r in result2.fetchall()]

    data = dict(compra._mapping)
    data["lineas"] = lineas
    return data
''')

with open(f'{base}/router.py', 'w', encoding='utf-8') as f:
    f.write('''from fastapi import APIRouter, Depends, HTTPException
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
''')

print("Modulo compras creado correctamente")
for archivo in ['__init__.py', 'schemas.py', 'service.py', 'router.py']:
    ruta = f'{base}/{archivo}'
    size = os.path.getsize(ruta)
    print(f"  {ruta}: {size} bytes")