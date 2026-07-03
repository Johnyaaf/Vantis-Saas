import uuid
from datetime import datetime, timezone
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import text
from app.modules.productos.schemas import ProductoCreate, ProductoUpdate


async def listar_productos(db: AsyncSession, schema: str, buscar: Optional[str] = None) -> list:
    if buscar:
        result = await db.execute(text(f"""
            SELECT id, sku, nombre, descripcion, id_categoria, tipo, unidad_medida,
                   precio_venta_neto, costo_unitario, stock_actual, stock_minimo,
                   stock_maximo, stock_critico, aplica_iva, maneja_stock, marca, activo, creado_en
            FROM "{schema}".productos
            WHERE activo = true AND (LOWER(nombre) LIKE :buscar OR LOWER(sku) LIKE :buscar)
            ORDER BY nombre
        """), {"buscar": f"%{buscar.lower()}%"})
    else:
        result = await db.execute(text(f"""
            SELECT id, sku, nombre, descripcion, id_categoria, tipo, unidad_medida,
                   precio_venta_neto, costo_unitario, stock_actual, stock_minimo,
                   stock_maximo, stock_critico, aplica_iva, maneja_stock, marca, activo, creado_en
            FROM "{schema}".productos
            WHERE activo = true
            ORDER BY nombre
        """))
    rows = result.fetchall()
    return [dict(row._mapping) for row in rows]


async def obtener_producto(db: AsyncSession, schema: str, producto_id: str) -> dict:
    result = await db.execute(text(f"""
        SELECT id, sku, nombre, descripcion, id_categoria, tipo, unidad_medida,
               precio_venta_neto, costo_unitario, stock_actual, stock_minimo,
               stock_maximo, stock_critico, aplica_iva, maneja_stock, marca, activo, creado_en
        FROM "{schema}".productos WHERE id = :id
    """), {"id": producto_id})
    row = result.fetchone()
    return dict(row._mapping) if row else None


async def crear_producto(db: AsyncSession, schema: str, datos: ProductoCreate) -> dict:
    result = await db.execute(text(f"""
        SELECT id FROM "{schema}".productos WHERE sku = :sku
    """), {"sku": datos.sku})
    if result.fetchone():
        return {"ok": False, "error": "Ya existe un producto con ese SKU"}

    producto_id = str(uuid.uuid4())
    ahora = datetime.now(timezone.utc)
    precio_bruto = round(datos.precio_venta_neto * 1.19) if datos.aplica_iva else datos.precio_venta_neto
    margen = round((datos.precio_venta_neto - datos.costo_unitario) / datos.precio_venta_neto * 100, 2) if datos.precio_venta_neto > 0 else 0

    await db.execute(text(f"""
        INSERT INTO "{schema}".productos
        (id, sku, nombre, descripcion, id_categoria, tipo, unidad_medida,
         precio_venta_neto, precio_venta_bruto, costo_unitario, margen_pct,
         stock_actual, stock_minimo, stock_maximo, stock_critico,
         aplica_iva, maneja_stock, marca, activo, creado_en)
        VALUES (:id, :sku, :nombre, :descripcion, :id_categoria, :tipo, :unidad_medida,
                :precio_venta_neto, :precio_venta_bruto, :costo_unitario, :margen_pct,
                0, :stock_minimo, :stock_maximo, :stock_critico,
                :aplica_iva, :maneja_stock, :marca, true, :creado_en)
    """), {
        "id": producto_id,
        "sku": datos.sku,
        "nombre": datos.nombre,
        "descripcion": datos.descripcion,
        "id_categoria": datos.id_categoria,
        "tipo": datos.tipo,
        "unidad_medida": datos.unidad_medida,
        "precio_venta_neto": datos.precio_venta_neto,
        "precio_venta_bruto": precio_bruto,
        "costo_unitario": datos.costo_unitario,
        "margen_pct": margen,
        "stock_minimo": datos.stock_minimo,
        "stock_maximo": datos.stock_maximo,
        "stock_critico": datos.stock_critico,
        "aplica_iva": datos.aplica_iva,
        "maneja_stock": datos.maneja_stock,
        "marca": datos.marca,
        "creado_en": ahora,
    })
    await db.commit()
    return {"ok": True, "id": producto_id}


async def actualizar_costo_cpp(db: AsyncSession, schema: str, producto_id: str,
                                cantidad_comprada: int, costo_compra: float) -> dict:
    result = await db.execute(text(f"""
        SELECT stock_actual, costo_unitario FROM "{schema}".productos WHERE id = :id
    """), {"id": producto_id})
    row = result.fetchone()
    if not row:
        return {"ok": False, "error": "Producto no encontrado"}

    stock_actual = row._mapping["stock_actual"]
    costo_actual = row._mapping["costo_unitario"]
    stock_nuevo = stock_actual + cantidad_comprada

    if stock_nuevo > 0:
        costo_nuevo = round((stock_actual * costo_actual + cantidad_comprada * costo_compra) / stock_nuevo, 2)
    else:
        costo_nuevo = costo_compra

    await db.execute(text(f"""
        UPDATE "{schema}".productos
        SET stock_actual = :stock_nuevo, costo_unitario = :costo_nuevo
        WHERE id = :id
    """), {"stock_nuevo": stock_nuevo, "costo_nuevo": costo_nuevo, "id": producto_id})
    await db.commit()
    return {"ok": True, "stock_nuevo": stock_nuevo, "costo_nuevo": costo_nuevo}


async def actualizar_producto(db: AsyncSession, schema: str, producto_id: str, datos: ProductoUpdate) -> dict:
    campos = {k: v for k, v in datos.model_dump().items() if v is not None}
    if not campos:
        return {"ok": False, "error": "Sin campos para actualizar"}

    sets = ", ".join([f"{k} = :{k}" for k in campos])
    campos["id"] = producto_id

    await db.execute(text(f"""
        UPDATE "{schema}".productos SET {sets} WHERE id = :id
    """), campos)
    await db.commit()
    return {"ok": True}


async def eliminar_producto(db: AsyncSession, schema: str, producto_id: str) -> dict:
    await db.execute(text(f"""
        UPDATE "{schema}".productos SET activo = false WHERE id = :id
    """), {"id": producto_id})
    await db.commit()
    return {"ok": True}
