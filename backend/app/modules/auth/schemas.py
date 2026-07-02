from pydantic import BaseModel, EmailStr


class RegisterRequest(BaseModel):
    nombre: str
    email: EmailStr
    password: str
    rut_empresa: str
    nombre_empresa: str
    plan: str = "lite"


class LoginRequest(BaseModel):
    email: EmailStr
    password: str


class TokenResponse(BaseModel):
    access_token: str
    refresh_token: str
    token_type: str = "bearer"
    usuario: dict
    tenant: dict


class RefreshRequest(BaseModel):
    refresh_token: str