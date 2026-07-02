import uuid
import re
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, text
from app.models.tenant import UsuarioGlobal, Tenant, UsuarioTenant, ModuloPlan
from app.core.security import hash_password, verify_password, create_access_token, create_refresh_token, validate_rut_chile
from app.modules.auth.schemas import RegisterRequest, LoginRequest


def _slug_from_nombre(nombre: str) -> str:
    slug = nombre.lower()
    slug = re.sub(r'[áàä]', 'a', slug)
    slug = re.sub(r'[éèë]', 'e', slug)
    slug = re.sub(r'[íìï]', 'i', slug)
    slug = re.sub(r'[óòö]', 'o', slug)
    slug = re.sub(r'[úùü]', 'u', slug)
    slug = re.sub(r'[^a-z0-9]', '-', slug)
    slug = re.sub(r'-+', '-', slug).strip('-')
    return slug[:40]


async def registrar_tenant(db: AsyncSession, datos: RegisterRequest) -> dict:
    if not validate_rut_chile(datos.rut_empresa):
        return {"ok": False, "error": "RUT de empresa inválido"}

    result = await db.execute(select(UsuarioGlobal).where(UsuarioGlobal.email == datos.email))
    if result.scalar_one_or_none():
        return {"ok": False, "error": "Este email ya está registrado"}

    result = await db.execute(select(Tenant).where(Tenant.rut == datos.rut_empresa))
    if result.scalar_one_or_none():
        return {"ok": False, "error": "Este RUT ya está registrado"}

    slug = _slug_from_nombre(datos.nombre_empresa)
    schema_name = f"tenant_{slug}_{str(uuid.uuid4())[:8]}"

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

    await db.execute(text(f'CREATE SCHEMA IF NOT EXISTS "{schema_name}"'))
    await db.commit()

    return {"ok": True, "tenant_id": tenant.id, "usuario_id": usuario.id}


async def login(db: AsyncSession, datos: LoginRequest) -> dict:
    result = await db.execute(
        select(UsuarioGlobal).where(UsuarioGlobal.email == datos.email)
    )
    usuario = result.scalar_one_or_none()

    if not usuario or not verify_password(datos.password, usuario.password_hash):
        return {"ok": False, "error": "Credenciales inválidas"}

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
            ModuloPlan.activo.is_(True)
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