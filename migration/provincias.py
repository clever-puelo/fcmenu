"""Decodificador de `Cliente.PCIA`/`FcivaVta.PCIA` (código de 1 letra, ej.
`"V "`) al nombre completo de la Provincia — réplica del array global
`PCIA()` de `FCMENU.bas:358-361` (`PCIA = Array("ASalta    ", "BBs. As.  ",
..., "VT.del Fgo", ...)`, decodificado en el resto del legacy vía
`Mid(PCIA(Asc(letra) - 65), 2, 10)`).

Ningún formulario ya migrado necesitaba mostrar el nombre completo (sólo
el código crudo, ver `ClienteDetalleDialog.txt_pcia` y
`FacturaService.letra_comprobante`, que sólo comparan contra `"V "` para
la regla de Tierra del Fuego) — este es el primer uso real de la
decodificación completa, para el panel de Cliente de sólo lectura de
`VerFact.frm`. Reusable por cualquier pantalla futura que la necesite
(`CtaCte.frm`, `EmiFact.frm`, `VerCheq.frm`, `Listados.frm` la usan en el
legacy con el mismo mecanismo).
"""

from __future__ import annotations

# Letras "I" y "O" quedan sin provincia asignada en el legacy (huecos
# reales del array original, no un error de transcripción).
NOMBRE_PROVINCIA: dict[str, str] = {
    "A": "Salta",
    "B": "Bs. As.",
    "C": "Cap.Fed.",
    "D": "San Luis",
    "E": "Ent.Ríos",
    "F": "La Rioja",
    "G": "Sgo.del E",
    "H": "Chaco",
    "J": "San Juan",
    "K": "Catamarca",
    "L": "La Pampa",
    "M": "Mendoza",
    "N": "Misiones",
    "P": "Formosa",
    "Q": "Neuquén",
    "R": "Río Negro",
    "S": "Santa Fe",
    "T": "Tucumán",
    "U": "Chubut",
    "V": "T.del Fgo",
    "W": "Corr'tes.",
    "X": "Córdoba",
    "Y": "Jujuy",
    "Z": "Sta.Cruz",
}


def nombre_provincia(codigo: str | None) -> str:
    """Nombre completo de la Provincia a partir de su código de 1 letra.

    Tolera el padding real de Access (`"V "`, ver notas de padding en
    `repository.py`) y devuelve `""` si el código está vacío o no
    corresponde a ninguna provincia conocida (mismo criterio que el
    legacy, que dejaba el 27º slot del array en blanco)."""
    if not codigo:
        return ""
    letra = codigo.strip()[:1].upper()
    return NOMBRE_PROVINCIA.get(letra, "")
