"""seed modulos por plan

Revision ID: 4fb7a31c9c21
Revises: d43b54232cc2
Create Date: 2026-07-02 20:20:00.000000

"""
from typing import Sequence, Union

from alembic import op


revision: str = "4fb7a31c9c21"
down_revision: Union[str, Sequence[str], None] = "d43b54232cc2"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    modulos = [
        ("lite", "dashboard"),
        ("lite", "clientes"),
        ("lite", "proveedores"),
        ("lite", "productos"),
        ("pro", "dashboard"),
        ("pro", "clientes"),
        ("pro", "proveedores"),
        ("pro", "productos"),
        ("pro", "ventas"),
        ("pro", "compras"),
        ("pro", "inventario"),
        ("pro", "cobranza"),
        ("full", "dashboard"),
        ("full", "clientes"),
        ("full", "proveedores"),
        ("full", "productos"),
        ("full", "ventas"),
        ("full", "compras"),
        ("full", "inventario"),
        ("full", "cobranza"),
        ("full", "finanzas"),
        ("full", "reportes"),
        ("full", "configuracion"),
    ]

    values = ", ".join(f"('{plan}', '{modulo}', true)" for plan, modulo in modulos)
    op.execute(f"""
        INSERT INTO public.modulos_plan (plan, modulo, activo)
        VALUES {values}
        ON CONFLICT (plan, modulo) DO UPDATE SET activo = EXCLUDED.activo
    """)


def downgrade() -> None:
    op.execute("""
        DELETE FROM public.modulos_plan
        WHERE plan IN ('lite', 'pro', 'full')
    """)
