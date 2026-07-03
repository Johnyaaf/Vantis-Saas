import re
import unicodedata
import uuid

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.security import (
    create_access_token,
    create_refresh_token,
    hash_password,
    validate_rut_chile,
    verify_password,
)
from app.core.tenant_schema import create_tenant_schema
from app.models.tenant import ModuloPlan, Tenant, UsuarioGlobal, UsuarioTenant
from app.modules.auth.schemas import LoginRequest, RegisterRequest


def _slug_from_nombre(nombre: str) -> str:
    slug = unicodedata.normalize("NFKD", nombre.lower())
    slug = "".join(char for char in slug if not unicodedata.combining(char))
    slug = re.sub(r"[^a-z0-9]", "-", slug)
    slug = re.sub(r"-+", "-", slug).strip("-")
    return slug[:32] or "empresa"


async def registrar_tenant(db: AsyncSession, datos: RegisterRequest) -> dict:
    if not validate_rut_chile(datos.rut_empresa):
        return {"ok": False, "error": "RUT de empresa invalido"}

    result = await db.execute(select(UsuarioGlobal).where(UsuarioGlobal.email == datos.email))
    if result.scalar_one_or_none():
        return {"ok": False, "error": "Este email ya esta registrado"}

    result = await db.execute(select(Tenant).where(Tenant.rut == datos.rut_empresa))
    if result.scalar_one_or_none():
        return {"ok": False, "error": "Este RUT ya esta registrado"}

    slug_base = _slug_from_nombre(datos.nombre_empresa)
    suffix = str(uuid.uuid4())[:8]
    slug = f"{slug_base}-{suffix}"[:50]
    schema_name = f"tenant_{slug_base.replace('-', '_')}_{suffix}"

    tenant = Tenant(
        slug=slug,
        rut=datos.rut_empresa,
        nombre=datos.nombre_empresa,
        plan=datos.plan,
        schema_name=schema_name,
    )
    db.add(tenant)
    await db.flush()

    usuario = UsuarioGlobal(
        email=datos.email,
        password_hash=hash_password(datos.password),
        nombre=datos.nombre,
    )
    db.add(usuario)
    await db.flush()

    vinculo = UsuarioTenant(
        usuario_id=usuario.id,
        tenant_id=tenant.id,
        rol="admin",
    )
    db.add(vinculo)

    await create_tenant_schema(db, schema_name)
    await db.commit()

    return {"ok": True, "tenant_id": tenant.id, "usuario_id": usuario.id}


async def login(db: AsyncSession, datos: LoginRequest) -> dict:
    result = await db.execute(
        select(UsuarioGlobal).where(UsuarioGlobal.email == datos.email)
    )
    usuario = result.scalar_one_or_none()

    if not usuario or not verify_password(datos.password, usuario.password_hash):
        return {"ok": False, "error": "Credenciales invalidas"}

    if not usuario.activo:
        return {"ok": False, "error": "Usuario inactivo"}

    result = await db.execute(
        select(UsuarioTenant, Tenant)
        .join(Tenant, UsuarioTenant.tenant_id == Tenant.id)
        .where(UsuarioTenant.usuario_id == usuario.id)
        .where(UsuarioTenant.activo.is_(True))
    )
    row = result.first()

    if not row:
        return {"ok": False, "error": "Usuario sin empresa asignada"}

    vinculo, tenant = row

    result = await db.execute(
        select(ModuloPlan).where(
            ModuloPlan.plan == tenant.plan,
            ModuloPlan.activo.is_(True),
        )
    )
    modulos = [m.modulo for m in result.scalars().all()]

    token_data = {
        "sub": usuario.id,
        "email": usuario.email,
        "tenant_id": tenant.id,
        "schema": tenant.schema_name,
        "rol": vinculo.rol,
        "plan": tenant.plan,
        "modulos": modulos,
    }

    return {
        "ok": True,
        "access_token": create_access_token(token_data),
        "refresh_token": create_refresh_token({"sub": usuario.id, "tenant_id": tenant.id}),
        "token_type": "bearer",
        "usuario": {
            "id": usuario.id,
            "nombre": usuario.nombre,
            "email": usuario.email,
            "rol": vinculo.rol,
        },
        "tenant": {
            "id": tenant.id,
            "nombre": tenant.nombre,
            "plan": tenant.plan,
            "modulos": modulos,
        },
    }
