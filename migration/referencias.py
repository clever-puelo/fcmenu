"""Generación del PDF de "Referencias" (Ayuda -> Referencias) — una
página de tips y ayudamemorias de atajos reales ya implementados en la
app, pedida por el usuario (2026-08-17): "un PDF en una carpeta interna
del sistema de una página con Tips y ayudamemorias".

Usa **ReportLab puro** (`reportlab.platypus`), mismo criterio que
[pdf.py](pdf.py) (CLAUDE.md: "Generación de PDF automáticos... mediante
ReportLab"). El PDF se genera una sola vez y se guarda en
`assets/docs/` (carpeta interna del proyecto, junto al logo) — llamadas
siguientes reusan el archivo ya generado.

El contenido no es inventado: son los atajos/comportamientos reales
documentados en el código de cada pantalla (`facturador_window.py`,
`recibo_window.py`, `stock_movimiento_window.py`, `widgets.py`
`EnterAsTabFilter`, etc.).
"""

from __future__ import annotations

from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import ListFlowable, ListItem, Paragraph, SimpleDocTemplate, Spacer

DIR_SALIDA_DEFAULT = Path("assets") / "docs"
NOMBRE_ARCHIVO = "Referencias_FCMENU.pdf"

# (sección, [tips]) — un tip por atajo/comportamiento real ya migrado.
SECCIONES: list[tuple[str, list[str]]] = [
    (
        "Navegación general",
        [
            "Enter avanza al siguiente campo, igual que Tab (no hace falta usar el mouse).",
            "F2, o Enter en un campo de código en blanco, abre el buscador de Cliente/Artículo.",
            "El botón «Cerrar» de cualquier pantalla la cierra y la elimina — no hace falta volver a cerrarla.",
        ],
    ),
    (
        "Factura / Recibo",
        [
            "Al abrir Factura o Recibo, primero aparece la pantalla y recién después el buscador de Cliente.",
            "En la grilla de Detalle: Sección → segmentos → cantidad, todo tipeado, sin diálogos.",
            "«Salir» avisa antes de cerrar si hay algo cargado sin emitir — evita perder el ingreso por error.",
            "El botón «Nota Clte.» se pinta de amarillo si el cliente elegido tiene una Nota guardada.",
        ],
    ),
    (
        "Recibo — Pendientes de Cobro",
        [
            "Un clic en cualquier columna (menos Aplicado/Descuento) aplica o quita el comprobante completo.",
            "Doble clic en Aplicado o Descuento permite un cobro parcial, con descuento por pronto pago.",
            "«Emitir» se habilita sólo cuando la Diferencia da exactamente $ 0,00.",
        ],
    ),
    (
        "Stock",
        [
            "Enter en «Nº Cpbante.» pasa directo a la grilla, lista para cargar renglones.",
            "Supr quita el renglón seleccionado de la grilla.",
        ],
    ),
    (
        "Tablas Varias",
        [
            "El combo «TABLA:» cambia de categoría (Secciones, Vendedores, Zonas, etc.) sin cerrar la ventana.",
        ],
    ),
]


def _construir_estilos():
    hoja = getSampleStyleSheet()
    return {
        "titulo": ParagraphStyle(
            "TituloReferencias", parent=hoja["Title"], fontSize=18, textColor="#1B4332", spaceAfter=4,
        ),
        "subtitulo": ParagraphStyle(
            "SubtituloReferencias", parent=hoja["Normal"], fontSize=10, textColor="#4A5D52", spaceAfter=14,
        ),
        "seccion": ParagraphStyle(
            "SeccionReferencias", parent=hoja["Heading2"], fontSize=12, textColor="#2D6A4F", spaceBefore=10, spaceAfter=4,
        ),
        "tip": ParagraphStyle("TipReferencias", parent=hoja["Normal"], fontSize=9.5, leading=13),
    }


def generar_pdf_referencias(directorio_salida: Path | None = None, *, forzar: bool = False) -> Path:
    """Genera (si no existe todavía, o si `forzar=True`) el PDF de
    Referencias y devuelve su ruta. Idempotente: si el archivo ya
    existe y no se pide forzar, no lo vuelve a generar."""
    directorio = Path(directorio_salida) if directorio_salida is not None else DIR_SALIDA_DEFAULT
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / NOMBRE_ARCHIVO

    if ruta.exists() and not forzar:
        return ruta

    estilos = _construir_estilos()
    doc = SimpleDocTemplate(
        str(ruta), pagesize=A4,
        topMargin=18 * mm, bottomMargin=18 * mm, leftMargin=20 * mm, rightMargin=20 * mm,
    )

    elementos = [
        Paragraph("FCMENU — Referencias", estilos["titulo"]),
        Paragraph("Tips y ayudamemorias de uso rápido", estilos["subtitulo"]),
    ]
    for titulo_seccion, tips in SECCIONES:
        elementos.append(Paragraph(titulo_seccion, estilos["seccion"]))
        elementos.append(
            ListFlowable(
                [ListItem(Paragraph(tip, estilos["tip"])) for tip in tips],
                bulletType="bullet",
                start="•",
                leftIndent=12,
            )
        )
    elementos.append(Spacer(1, 10 * mm))

    doc.build(elementos)
    return ruta
