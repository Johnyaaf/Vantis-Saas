from sqlalchemy.ext.asyncio import create_async_engine, AsyncSession, async_sessionmaker
from sqlalchemy.orm import DeclarativeBase
from sqlalchemy import text
from app.core.config import settings


# Motor principal de la base de datos
engine = create_async_engine(
    settings.DATABASE_URL,
    echo=settings.DEBUG,
    pool_pre_ping=True,
    pool_size=10,
    max_overflow=20,
)

# Fábrica de sesiones asíncronas
AsyncSessionLocal = async_sessionmaker(
    engine,
    class_=AsyncSession,
    expire_on_commit=False,
)


# Base para todos los modelos ORM
class Base(DeclarativeBase):
    pass


async def get_db():
    """Dependencia FastAPI: entrega una sesión de BD y la cierra al terminar."""
    async with AsyncSessionLocal() as session:
        try:
            yield session
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()


async def set_schema(session: AsyncSession, schema: str):
    """Cambia el search_path al schema del tenant activo."""
    await session.execute(text(f"SET search_path TO {schema}, public"))


async def get_tenant_db(schema: str):
    """Entrega una sesión ya apuntando al schema del tenant."""
    async with AsyncSessionLocal() as session:
        await set_schema(session, schema)
        try:
            yield session
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()