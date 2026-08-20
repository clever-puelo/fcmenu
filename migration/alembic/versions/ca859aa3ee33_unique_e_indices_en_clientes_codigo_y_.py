"""unique e indices en Clientes.CODIGO y Articulo COD1 COD2

Revision ID: ca859aa3ee33
Revises: 8cd556285914
Create Date: 2026-08-19 20:27:30.094638

Primer paso, de bajo riesgo, hacia integridad referencial real (pedido
del usuario, 2026-08-19, "el fk de articulos y clientes que son los
simples") — ver `Documentacion_Tecnica_Arquitectura.md` §3.2 para el
hallazgo original ("ni un solo `ForeignKey()`/`UNIQUE`/índice en toda
la base"). Alcance de ESTA migración, verificado contra `fcmenu_dev`
real antes de escribirla:

- `UNIQUE` en `Clientes.CODIGO` — sin duplicados reales.
- Índice único por expresión (`trim(COD1), trim(COD2)`) en `Articulo`
  — sin colisiones reales una vez descartado el padding fijo de Access
  (mismo bug ya documentado en varios repos de este proyecto). Se usa
  una expresión en vez de `UNIQUE` directo sobre las columnas para no
  tener que reescribir el dato histórico migrado 1:1 del `.mdb`.
- Índices simples (sin `UNIQUE`, sin `ForeignKey` todavía) en las
  columnas que referencian a Cliente/Articulo desde otras tablas
  (`Ctasctes.CLTE`, `FcivaVta.CLTE`, `DtoxClte.CLTE`,
  `Despachos(COD1,COD2)`, `Fcestad1(COD1,COD2)`) — todas tienen algún
  huérfano real (30/6/117/58 filas respectivamente; Fcestad1 tiene
  41.078, la mayoría por diseño — ítem libre "**"/líneas sin Artículo
  real, ver `Fcestad1.__table_args__`), así que declarar `ForeignKey()`
  ahora rompería la carga de esos datos históricos. Eso queda para una
  migración futura, evaluando `NOT VALID` y/o limpieza puntual de cada
  caso.
- De yapa, `FcivaVta`/`Fcestad1` también suman un índice sobre
  TIPO/LETRA/PTOVTA/CPBTE (`by_comprobante()`) — el lookup que el bug
  real de "no funciona el click en Cuentas Corrientes" (mismo día) dejó
  andando recién ahora.
"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision: str = 'ca859aa3ee33'
down_revision: Union[str, Sequence[str], None] = '8cd556285914'
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    """Upgrade schema."""
    op.create_unique_constraint("uq_clientes_codigo", "Clientes", ["CODIGO"])

    op.create_index(
        "ix_articulo_cod1_cod2_trim_unique",
        "Articulo",
        [sa.text('trim("COD1")'), sa.text('trim("COD2")')],
        unique=True,
    )

    op.create_index("ix_ctasctes_clte", "Ctasctes", ["CLTE"])
    op.create_index("ix_fcivavta_clte", "FcivaVta", ["CLTE"])
    op.create_index("ix_fcivavta_comprobante", "FcivaVta", ["TIPO", "LETRA", "PTOVTA", "CPBTE"])
    op.create_index("ix_dtoxclte_clte_seccion", "DtoxClte", ["CLTE", "SECCION"])
    op.create_index("ix_despachos_cod1_cod2", "Despachos", ["COD1", "COD2"])
    op.create_index("ix_despachos_nrodesp_fecent", "Despachos", ["NRODESP", "FECENT"])
    op.create_index("ix_fcestad1_cod1_cod2", "Fcestad1", ["COD1", "COD2"])
    op.create_index("ix_fcestad1_comprobante", "Fcestad1", ["TIPO", "LETRA", "PTOVTA", "CPBTE"])


def downgrade() -> None:
    """Downgrade schema."""
    op.drop_index("ix_fcestad1_comprobante", table_name="Fcestad1")
    op.drop_index("ix_fcestad1_cod1_cod2", table_name="Fcestad1")
    op.drop_index("ix_despachos_nrodesp_fecent", table_name="Despachos")
    op.drop_index("ix_despachos_cod1_cod2", table_name="Despachos")
    op.drop_index("ix_dtoxclte_clte_seccion", table_name="DtoxClte")
    op.drop_index("ix_fcivavta_comprobante", table_name="FcivaVta")
    op.drop_index("ix_fcivavta_clte", table_name="FcivaVta")
    op.drop_index("ix_ctasctes_clte", table_name="Ctasctes")
    op.drop_index("ix_articulo_cod1_cod2_trim_unique", table_name="Articulo")
    op.drop_constraint("uq_clientes_codigo", "Clientes", type_="unique")
