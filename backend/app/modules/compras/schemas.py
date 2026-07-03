from pydantic import BaseModel
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
