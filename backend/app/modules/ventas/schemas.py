from pydantic import BaseModel
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
