from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class ProveedorCreate(BaseModel):
    rut: str
    razon_social: str
    nombre_fantasia: Optional[str] = None
    giro: Optional[str] = None
    email: Optional[str] = None
    telefono: Optional[str] = None
    direccion: Optional[str] = None
    comuna: Optional[str] = None
    region: Optional[str] = None
    condicion_pago: str = "CONTADO"
    plazo_entrega_dias: int = 7
    moneda_habitual: str = "CLP"
    critico: bool = False


class ProveedorUpdate(BaseModel):
    razon_social: Optional[str] = None
    nombre_fantasia: Optional[str] = None
    giro: Optional[str] = None
    email: Optional[str] = None
    telefono: Optional[str] = None
    direccion: Optional[str] = None
    comuna: Optional[str] = None
    region: Optional[str] = None
    condicion_pago: Optional[str] = None
    plazo_entrega_dias: Optional[int] = None
    moneda_habitual: Optional[str] = None
    critico: Optional[bool] = None
    activo: Optional[bool] = None


class ProveedorResponse(BaseModel):
    id: str
    rut: str
    razon_social: str
    nombre_fantasia: Optional[str]
    giro: Optional[str]
    email: Optional[str]
    telefono: Optional[str]
    direccion: Optional[str]
    comuna: Optional[str]
    region: Optional[str]
    condicion_pago: str
    plazo_entrega_dias: int
    moneda_habitual: str
    critico: bool
    activo: bool
    creado_en: datetime

    class Config:
        from_attributes = True
