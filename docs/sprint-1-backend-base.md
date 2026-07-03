# Sprint 1: backend base SaaS

Objetivo del sprint:

- Registrar empresa.
- Crear schema tenant.
- Crear tabla `clientes` dentro del schema tenant.
- Iniciar sesion.
- Obtener token con modulos por plan.
- Consumir `/api/clientes` con JWT.

## Cambios incluidos

- `backend/requirements.txt`
- `backend/app/core/tenant_schema.py`
- `backend/migrations/versions/4fb7a31c9c21_seed_modulos_plan.py`
- Ajuste de `backend/app/modules/auth/service.py`
- Ajuste de permisos en `backend/app/modules/clientes/router.py`

## Instalar dependencias

Desde la raiz del repo:

```bash
cd backend
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## Ejecutar migraciones

PostgreSQL debe tener creada la base:

```txt
vantis_db
```

Y el usuario configurado en `backend/app/core/config.py` o `.env`:

```txt
vantis / vantis123
```

Luego:

```bash
alembic upgrade head
```

## Levantar API

```bash
uvicorn app.main:vantis --reload --host 127.0.0.1 --port 8000
```

Abrir:

```txt
http://127.0.0.1:8000/api/docs
```

## Flujo esperado

1. `POST /api/auth/register`
2. `POST /api/auth/login`
3. Copiar `access_token`
4. Autorizar Swagger con `Bearer <token>`
5. Probar `POST /api/clientes`
6. Probar `GET /api/clientes`

## Nota tecnica

La tabla `clientes` ahora se crea automaticamente dentro del schema del tenant durante el registro. Esto desbloquea el primer flujo SaaS real sin Google Sheets.
