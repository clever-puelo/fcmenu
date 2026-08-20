"""FK NOT VALID Ctascte/FcivaVta/DtoxClte hacia Clientes

Revision ID: bad21413e66c
Revises: ca859aa3ee33
Create Date: 2026-08-19 21:10:00.000000

Segundo paso de integridad referencial real (pedido del usuario,
2026-08-19: "como se puede implementar FK en el resto del sistema?
[...] lo mas moderno, solido, funcional posible") — primeras
`ForeignKey()` reales del proyecto (hasta ahora, cero en toda la base,
ver `Documentacion_Tecnica_Arquitectura.md` §3.2).

`Ctasctes.CLTE`/`FcivaVta.CLTE`/`DtoxClte.CLTE` → `Clientes.CODIGO`,
las 3 como `NOT VALID` (Postgres: protege todo INSERT/UPDATE nuevo de
ahora en más, sin exigir limpiar antes el historial ya inconsistente) —
verificado contra `fcmenu_dev` real antes de escribir esto: 30/6/117
filas huérfanas respectivamente (un CLTE que ya no existe en Clientes),
el usuario decidió explícitamente NO revisarlas a mano ahora.

**No incluye `Despachos` → `Articulo`** — no es una decisión de negocio
pendiente, es una limitación técnica real de Postgres: el índice único
de `Articulo` es por EXPRESIÓN (`trim(COD1,COD2)`, ver `ix_articulo_
cod1_cod2_trim_unique` en la migración anterior) porque el dato real
tiene padding de Access en ambos lados — Postgres no permite que el
lado referenciado de una FK sea un índice por expresión, sólo una
constraint UNIQUE/PK sobre columnas literales. Requeriría reescribir el
dato histórico primero (decisión aparte, más grande). Sigue con sólo
índice simple.

**Tampoco incluye `Fcestad1` → `Articulo`** — deliberado, no técnico:
~17% de esas filas (41.078 de 238.143) no tienen Artículo real A
PROPÓSITO (ítem libre "**"/líneas que usan la descripción de Sección,
ver `factura_renglon.SECCION_ITEM_LIBRE`) — una FK ahí estaría mal,
no sólo sucia.
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = 'bad21413e66c'
down_revision: Union[str, Sequence[str], None] = 'ca859aa3ee33'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    op.create_foreign_key(
        "fk_ctasctes_clte_clientes",
        "Ctasctes", "Clientes",
        ["CLTE"], ["CODIGO"],
        postgresql_not_valid=True,
    )
    op.create_foreign_key(
        "fk_fcivavta_clte_clientes",
        "FcivaVta", "Clientes",
        ["CLTE"], ["CODIGO"],
        postgresql_not_valid=True,
    )
    op.create_foreign_key(
        "fk_dtoxclte_clte_clientes",
        "DtoxClte", "Clientes",
        ["CLTE"], ["CODIGO"],
        postgresql_not_valid=True,
    )


def downgrade() -> None:
    """Downgrade schema."""
    op.drop_constraint("fk_dtoxclte_clte_clientes", "DtoxClte", type_="foreignkey")
    op.drop_constraint("fk_fcivavta_clte_clientes", "FcivaVta", type_="foreignkey")
    op.drop_constraint("fk_ctasctes_clte_clientes", "Ctasctes", type_="foreignkey")
