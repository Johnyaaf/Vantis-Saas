import re

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession


_IDENTIFIER_RE = re.compile(r"^[a-z][a-z0-9_]{0,62}$")


def validate_schema_name(schema_name: str) -> str:
    """Validate a generated PostgreSQL schema identifier before interpolation."""
    if not _IDENTIFIER_RE.fullmatch(schema_name):
        raise ValueError("Nombre de schema tenant invalido")
    return schema_name


async def create_tenant_schema(session: AsyncSession, schema_name: str) -> None:
    schema = validate_schema_name(schema_name)

    await session.execute(text(f'CREATE SCHEMA IF NOT EXISTS "{schema}"'))
    await session.execute(text(f"""
        CREATE TABLE IF NOT EXISTS "{schema}".clientes (
            id VARCHAR(36) PRIMARY KEY,
            rut VARCHAR(20) NOT NULL UNIQUE,
            razon_social VARCHAR(200) NOT NULL,
            nombre_fantasia VARCHAR(200),
            giro VARCHAR(200),
            email VARCHAR(255),
            telefono VARCHAR(50),
            direccion VARCHAR(255),
            comuna VARCHAR(100),
            region VARCHAR(100),
            condicion_pago VARCHAR(30) NOT NULL DEFAULT 'CONTADO',
            limite_credito NUMERIC(14, 2) NOT NULL DEFAULT 0,
            activo BOOLEAN NOT NULL DEFAULT true,
            creado_en TIMESTAMP WITH TIME ZONE NOT NULL,
            actualizado_en TIMESTAMP WITH TIME ZONE
        )
    """))
    await session.execute(text(f"""
        CREATE INDEX IF NOT EXISTS ix_clientes_razon_social
        ON "{schema}".clientes (razon_social)
    """))
