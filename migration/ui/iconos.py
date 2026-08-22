"""Íconos simples y evidentes para el menú principal y las pantallas —
pedido del usuario (2026-08-16): "Iconos: Simples, practicos, evidentes,
nada rebuscado".

**Preparado para cuando el usuario deje íconos reales en una carpeta**:
`icono(clave)` busca primero un archivo en `assets/icons/<clave>.png`
(o `.ico`/`.svg`) junto a este módulo — si no lo encuentra, genera un
ícono simple (glifo/emoji sobre una tarjeta redondeada verde pastel) al
vuelo con `QPainter`, para que el menú nunca se vea vacío mientras
tanto. Reemplazar un ícono más adelante es sólo copiar el archivo real
a `assets/icons/`, sin tocar código.
"""

from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import QRectF, QSize, Qt
from PyQt6.QtGui import QColor, QFont, QIcon, QPainter, QPixmap

from .theme import Verde

DIR_ICONOS = Path(__file__).resolve().parent / "assets" / "icons"

# clave -> (glifo, color de fondo) — un glifo evidente por pantalla/
# sección, mismo criterio del pedido ("evidentes, nada rebuscado").
GLIFOS: dict[str, tuple[str, str]] = {
    "clientes": ("👤", Verde.PASTEL),
    "articulos": ("📦", Verde.PASTEL),
    "descuentos": ("💲", Verde.PASTEL),
    "precios": ("🏷️", Verde.PASTEL),
    "tablas": ("🗂️", Verde.PASTEL),
    "cotizacion": ("💱", Verde.PASTEL),
    "parametros": ("⚙️", Verde.PASTEL),
    "constancia_afip": ("📄", Verde.CLARO),
    "busqueda": ("🔍", Verde.PASTEL),
    "detalle": ("🔎", Verde.PASTEL),
    "ctacte": ("📒", Verde.PASTEL),
    "ventas": ("📈", Verde.PASTEL),
    "stock": ("🏬", Verde.PASTEL),
    "despachos": ("🚚", Verde.PASTEL),
    "cobranzas": ("🤝", Verde.PASTEL),
    "facturas_emitidas": ("🧾", Verde.PASTEL),
    "totales": ("📊", Verde.PASTEL),
    "cheques": ("🏦", Verde.PASTEL),
    "factura": ("🧾", Verde.MEDIO_CLARO),
    "recibo": ("💵", Verde.MEDIO_CLARO),
    "ncredito": ("↩️", Verde.CLARO),
    # Mismo glifo que "ncredito" — "N.Créd. x Mercadería" es el mismo
    # concepto (Nota de Crédito), sólo la variante con reversa de Stock.
    # Bug real encontrado el 2026-08-20: esta clave se usaba en la barra
    # de tareas (main_menu_window.py) sin existir acá, así que caía
    # siempre al bullet genérico "•" — se nota mucho más ahora que los
    # íconos se agrandaron al doble.
    "ncredito_merc": ("↩️", Verde.CLARO),
    "ingreso_stock": ("📥", Verde.MEDIO_CLARO),
    "listados": ("📋", Verde.MEDIO_CLARO),
    "calculadora": ("🧮", Verde.PASTEL),
    "calendario": ("📅", Verde.PASTEL),
    "notas": ("📝", Verde.PASTEL),
    "arreglos": ("🔧", Verde.CLARO),
    "salir": ("🚪", Verde.CLARO),
    "ayuda": ("❓", Verde.CLARO),
    "app": ("🧾", Verde.MEDIO),
    # -- Botones del panel lateral de MainMenuWindow (secciones, no pantallas
    # puntuales — claves separadas de las de arriba a propósito, aunque
    # compartan glifo con alguna). "nav_listados" reusa el glifo de
    # "listados": misma idea (una lista), no hace falta uno nuevo.
    "nav_abms": ("🗃️", Verde.MEDIO_CLARO),
    "nav_ingreso": ("📥", Verde.MEDIO_CLARO),
    "nav_consultas": ("🔎", Verde.MEDIO_CLARO),
    "nav_listados": ("📋", Verde.MEDIO_CLARO),
    "nav_varios": ("🧰", Verde.MEDIO_CLARO),
    # "salvavidas" — sección Ayuda (pedido del usuario, 2026-08-17).
    "nav_ayuda": ("🛟", Verde.MEDIO_CLARO),
    # -- Botones de la sección Ayuda.
    "acerca_de": ("ℹ️", Verde.CLARO),
    "referencias": ("📑", Verde.CLARO),
    "manual": ("📖", Verde.CLARO),
    "sitio_web": ("🌐", Verde.CLARO),
    "lista_precios_pdf": ("🏷️", Verde.CLARO),
    # -- Comprobantes de Ingreso todavía sin implementar (2026-08-17):
    # botones visibles pero deshabilitados ("Próximamente"), mismo
    # criterio ya usado para "Constancia AFIP" en Varios.
    "ndebito": ("↪️", Verde.CLARO),
    "cotizacion_venta": ("📃", Verde.CLARO),
    "remito": ("🚛", Verde.CLARO),
    # Nota de Crédito Interna (2026-08-20) — mismo glifo que "ncredito"/
    # "ncredito_merc" (es la misma familia conceptual, Nota de Crédito),
    # color distinto para diferenciarla de un vistazo en la barra.
    "nci": ("↩️", Verde.MEDIO_CLARO),
    # Botón "Ver en el Mapa" de la ficha de Cliente (2026-08-21) — pin de
    # ubicación genérico, no el logo real de Google Maps (evitar usar una
    # marca ajena como ícono propio). Abre la dirección del cliente en un
    # navegador embebido dentro de la propia app (`navegador_dialog.py`),
    # no en el navegador externo — ver `cliente_detalle_dialog.py`.
    #
    # Bug real reportado por el usuario (2026-08-21): el botón apuntaba
    # primero a Google EARTH (globo 3D, `"earth"`) — es una SPA pesada
    # (Flutter + WebGL/WebGPU) que tira errores de consola y no anda bien
    # embebida (Chromium de QtWebEngine sin aceleración GPU completa).
    # Para el objetivo real ("ver la ubicación en el mapa") alcanza y
    # sobra con Google MAPS en modo `embed` — mapa 2D liviano, sin 3D/
    # WebGL, mismo `NavegadorDialog` (QWebEngineView) ya andando.
    "mapa": ("📍", Verde.PASTEL),
}

_CACHE: dict[tuple[str, int], QIcon] = {}


def icono(clave: str, tamano: int = 40) -> QIcon:
    cache_key = (clave, tamano)
    if cache_key in _CACHE:
        return _CACHE[cache_key]

    for extension in (".png", ".ico", ".svg"):
        ruta = DIR_ICONOS / f"{clave}{extension}"
        if ruta.exists():
            resultado = QIcon(str(ruta))
            _CACHE[cache_key] = resultado
            return resultado

    resultado = QIcon(_generar_pixmap(clave, tamano))
    _CACHE[cache_key] = resultado
    return resultado


def _generar_pixmap(clave: str, tamano: int) -> QPixmap:
    glifo, color_fondo = GLIFOS.get(clave, ("•", Verde.PASTEL))

    pixmap = QPixmap(QSize(tamano, tamano))
    pixmap.fill(Qt.GlobalColor.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)

    radio = tamano * 0.22
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor(color_fondo))
    painter.drawRoundedRect(QRectF(0, 0, tamano, tamano), radio, radio)

    fuente = QFont("Segoe UI Emoji")
    fuente.setPixelSize(int(tamano * 0.55))
    painter.setFont(fuente)
    painter.setPen(QColor(Verde.OSCURO))
    painter.drawText(QRectF(0, 0, tamano, tamano), Qt.AlignmentFlag.AlignCenter, glifo)

    painter.end()
    return pixmap
