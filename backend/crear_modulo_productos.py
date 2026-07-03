import os

base = 'app/modules/productos'
os.makedirs(base, exist_ok=True)

with open(f'{base}/__init__.py', 'w', encoding='utf-8') as f:
    f.write('')

with open(f'{base}/schemas.py', 'w', encoding='utf-8') as f:
    f.write('''from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class ProductoCreate(BaseModel):
    sku: str
    nombre: str
    descripcion: Optional[str] = None
    id_categoria: Optional[str] = None
    tipo: str = "mercaderia"
    unidad_medida: str = "unidad"
    precio_venta_neto: float = 0
    costo_unitario: float = 0
    stock_minimo: int = 0
    stock_maximo: int = 0
    stock_critico: int = 0
    aplica_iva: bool = True
    maneja_stock: bool = True
    marca: Optional[str] = None


class ProductoUpdate(BaseModel):
    nombre: Optional[str] = None
    descripcion: Optional[str] = None
    id_categoria: Optional[str] = None
    unidad_medida: Optional[str] = None
    precio_venta_neto: Optional[float] = None
    stock_minimo: Optional[int] = None
    stock_maximo: Optional[int] = None
    stock_critico: Optional[int] = None
    aplica_iva: Optional[bool] = None
    maneja_stock: Optional[bool] = None
    marca: Optional[str] = None
    activo: Optional[bool] = None


class ProductoResponse(BaseModel):
    id: str
    sku: str
    nombre: str
    descripcion: Optional[str]
    id_categoria: Optional[str]
    tipo: str
    unidad_medida: str
    precio_venta_neto: float
    costo_unitario: float
    stock_actual: int
    stock_minimo: int
    stock_maximo: int
    stock_critico: int
    aplica_iva: bool
    maneja_stock: bool
    marca: Optional[str]
    activo: bool
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
''')

with open(f'{base}/router.py', 'w', encoding='utf-8') as f:
    f.write('''from fastapi import APIRouter, Depends, HTTPException
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.database import get_db
from app.core.dependencies import get_current_user, require_modulo
from app.modules.productos.schemas import ProductoCreate, ProductoUpdate
from app.modules.productos import service

router = APIRouter(prefix="/api/productos", tags=["Productos"])


@router.get("")
async def listar(
    buscar: Optional[str] = None,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("inventario"))
) -> list:
    return await service.listar_productos(db, user["schema"], buscar)


@router.get("/{producto_id}")
async def obtener(
    producto_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    producto = await service.obtener_producto(db, user["schema"], producto_id)
    if not producto:
        raise HTTPException(status_code=404, detail="Producto no encontrado")
    return producto


@router.post("", status_code=201)
async def crear(
    datos: ProductoCreate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(require_modulo("inventario"))
) -> dict:
    resultado = await service.crear_producto(db, user["schema"], datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado


@router.put("/{producto_id}")
async def actualizar(
    producto_id: str,
    datos: ProductoUpdate,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    resultado = await service.actualizar_producto(db, user["schema"], producto_id, datos)
    if not resultado["ok"]:
        raise HTTPException(status_code=400, detail=resultado["error"])
    return resultado


@router.delete("/{producto_id}")
async def eliminar(
    producto_id: str,
    db: AsyncSession = Depends(get_db),
    user: dict = Depends(get_current_user)
) -> dict:
    return await service.eliminar_producto(db, user["schema"], producto_id)
''')

print("Modulo productos creado correctamente")
for archivo in ['__init__.py', 'schemas.py', 'service.py', 'router.py']:
    ruta = f'{base}/{archivo}'
    size = os.path.getsize(ruta)
    print(f"  {ruta}: {size} bytes")