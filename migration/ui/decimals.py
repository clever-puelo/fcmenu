"""Re-export de compatibilidad — la implementación real vive en
`migration/decimals.py` (movida ahí 2026-08-15 para que `services.py`
también pueda usar `format_decimal()` sin que `migration` dependa de
`migration.ui`). Se mantiene este módulo para no tocar los ~11 imports
`from .decimals import ...` ya existentes en toda la capa de UI.
"""

from __future__ import annotations

from migration.decimals import format_decimal, parse_decimal

__all__ = ["parse_decimal", "format_decimal"]
