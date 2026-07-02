from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from app.core.security import decode_token

security = HTTPBearer()


def get_current_user(credentials: HTTPAuthorizationCredentials = Depends(security)) -> dict:
    """Extrae y valida el usuario del JWT en cada request protegido."""
    token = credentials.credentials
    payload = decode_token(token)

    if not payload:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Token inválido o expirado",
            headers={"WWW-Authenticate": "Bearer"},
        )

    if payload.get("type") != "access":
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Tipo de token incorrecto",
        )

    return {
        "id": payload.get("sub"),
        "email": payload.get("email"),
        "tenant_id": payload.get("tenant_id"),
        "schema": payload.get("schema"),
        "rol": payload.get("rol"),
        "plan": payload.get("plan"),
        "modulos": payload.get("modulos", []),
    }


def require_modulo(modulo: str):
    """Verifica que el plan del tenant incluya el módulo requerido."""
    def _check(user: dict = Depends(get_current_user)) -> dict:
        if modulo not in user.get("modulos", []):
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail=f"Tu plan no incluye el módulo '{modulo}'",
            )
        return user
    return _check


def require_rol(roles: list[str]):
    """Verifica que el usuario tenga uno de los roles requeridos."""
    def _check(user: dict = Depends(get_current_user)) -> dict:
        if user.get("rol") not in roles:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="No tienes permisos para esta acción",
            )
        return user
    return _check