from pydantic import BaseModel
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
