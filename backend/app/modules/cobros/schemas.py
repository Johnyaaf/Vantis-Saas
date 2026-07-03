from pydantic import BaseModel
from typing import Optional
from datetime import datetime


class CobroCreate(BaseModel):
    cxc_id: str
    monto: float
    medio_pago: str
    fecha: str
    referencia: Optional[str] = None


class PagoCreate(BaseModel):
    cxp_id: str
    monto: float
    medio_pago: str
    fecha: str
    referencia: Optional[str] = None
