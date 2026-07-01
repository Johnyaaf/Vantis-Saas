from pydantic_settings import BaseSettings
from typing import List


class Settings(BaseSettings):
    # Aplicación
    APP_NAME: str = "VANTIS ERP"
    APP_VERSION: str = "1.0.0"
    DEBUG: bool = True
    SECRET_KEY: str = "vantis-secret-key-cambiar-en-produccion-2026"

    # Base de datos
    DATABASE_URL: str = "postgresql+asyncpg://vantis:vantis123@localhost:5432/vantis_db"

    # Redis
    REDIS_URL: str = "redis://localhost:6379/0"

    # JWT
    JWT_SECRET_KEY: str = "jwt-secret-key-cambiar-en-produccion-2026"
    JWT_ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 15
    REFRESH_TOKEN_EXPIRE_DAYS: int = 7

    # CORS
    ALLOWED_ORIGINS: List[str] = ["http://localhost:3000", "http://localhost:8000"]

    model_config = {
        "env_file": ".env",
        "env_file_encoding": "utf-8",
    }


settings = Settings()