from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class ClienteCreate(BaseModel):
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
    limite_credito: float = 0


class ClienteUpdate(BaseModel):
    razon_social: Optional[str] = None
    nombre_fantasia: Optional[str] = None
    giro: Optional[str] = None
    email: Optional[str] = None
    telefono: Optional[str] = None
    direccion: Optional[str] = None
    comuna: Optional[str] = None
    region: Optional[str] = None
    condicion_pago: Optional[str] = None
    limite_credito: Optional[float] = None
    activo: Optional[bool] = None


class ClienteResponse(BaseModel):
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
    limite_credito: float
    activo: bool
    creado_en: datetime

    class Config:
        from_attributes = True
