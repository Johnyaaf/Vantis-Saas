import uuid
from datetime import datetime, timezone
from sqlalchemy import String, Boolean, DateTime, Text
from sqlalchemy.orm import Mapped, mapped_column
from app.core.database import Base


class Tenant(Base):
    __tablename__ = "tenants"
    __table_args__ = {"schema": "public"}

    id: Mapped[str] = mapped_column(
        String(36), primary_key=True, default=lambda: str(uuid.uuid4())
    )
    slug: Mapped[str] = mapped_column(String(50), unique=True, nullable=False)
    rut: Mapped[str] = mapped_column(String(20), unique=True, nullable=False)
    nombre: Mapped[str] = mapped_column(String(200), nullable=False)
    plan: Mapped[str] = mapped_column(String(20), default="lite")
    estado: Mapped[str] = mapped_column(String(20), default="activo")
    schema_name: Mapped[str] = mapped_column(String(63), unique=True, nullable=False)
    creado_en: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), default=lambda: datetime.now(timezone.utc)
    )
    vence_en: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)


class ModuloPlan(Base):
    __tablename__ = "modulos_plan"
    __table_args__ = {"schema": "public"}

    plan: Mapped[str] = mapped_column(String(20), primary_key=True)
    modulo: Mapped[str] = mapped_column(String(50), primary_key=True)
    activo: Mapped[bool] = mapped_column(Boolean, default=True)


class UsuarioGlobal(Base):
    __tablename__ = "usuarios_globales"
    __table_args__ = {"schema": "public"}

    id: Mapped[str] = mapped_column(
        String(36), primary_key=True, default=lambda: str(uuid.uuid4())
    )
    email: Mapped[str] = mapped_column(String(255), unique=True, nullable=False)
    password_hash: Mapped[str] = mapped_column(String(255), nullable=False)
    nombre: Mapped[str] = mapped_column(String(200), nullable=False)
    activo: Mapped[bool] = mapped_column(Boolean, default=True)
    creado_en: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), default=lambda: datetime.now(timezone.utc)
    )


class UsuarioTenant(Base):
    __tablename__ = "usuario_tenant"
    __table_args__ = {"schema": "public"}

    id: Mapped[str] = mapped_column(
        String(36), primary_key=True, default=lambda: str(uuid.uuid4())
    )
    usuario_id: Mapped[str] = mapped_column(String(36), nullable=False)
    tenant_id: Mapped[str] = mapped_column(String(36), nullable=False)
    rol: Mapped[str] = mapped_column(String(30), default="admin")
    activo: Mapped[bool] = mapped_column(Boolean, default=True)
    creado_en: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), default=lambda: datetime.now(timezone.utc)
    )