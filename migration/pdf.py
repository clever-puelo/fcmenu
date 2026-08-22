"""Generación de PDF de Factura Electrónica — reemplaza el circuito de
impresión de `EmiFact.frm` (VSPrinter/formulario preimpreso, fuera de
alcance) y el QR externo (`Shell("qrcode.exe ...")`, bug real ya
corregido en `migration/afip.py`).

Usa **ReportLab puro** (`reportlab.pdfgen.canvas` + `reportlab.graphics.
barcode.qr.QrCodeWidget`) — el QR se dibuja nativamente dentro del PDF,
sin depender de un generador de imágenes externo ni de una librería de
QR aparte (CLAUDE.md: "Generación de PDF automáticos... mediante
ReportLab").

**Réplica visual del legacy (2026-08-20, 2 rondas)**: el usuario mandó
una Factura y un Recibo REALES ("deben quedar iguales") — primera
ronda: se releyó `EmiFact.frm`/`EmiRec.frm`/`NCInterna.frm` con esas
muestras en mano. Segunda ronda (el resultado seguía sin coincidir): el
usuario aclaró que el procedimiento real de impresión electrónica es
`EmiFact.frm Sub IMPRIME_Electronica()` (línea 1250) — no `Sub IMPRIME()`
(línea 700, código muerto/de impresión en papel, que es lo que se había
leído primero por error — mismo nombre de variable `TIPO`/casi mismo
código, fácil de confundir). Desde `IMPRIME_Electronica` se transcribieron
las coordenadas ABSOLUTAS reales (mm desde el borde de la hoja, sistema
propio de `Escribe(Font,Size,Negr,Ital,Lin,Col,Texto)` en `FCMENU.bas`) —
ver docstring de `_dibujar_pagina_factura` para el detalle línea por
línea. Lo mismo para Recibo (`EmiRec.frm Sub IMPRIME2en1`, confirmado
por el usuario como el procedimiento real) y Nota de Crédito Interna
(`NCInterna.frm Sub ImprimeNCI`, mismo letterhead).

Hallazgos/correcciones reales de esta segunda ronda:
- El logo real es `assets/LOGO2.jpg` (`Picture1` en los 3 `.frm`) — NO
  `Logo-Alestel.png` (ése es el ícono de `MainMenuWindow`, con un
  círculo de fondo, un logo distinto).
- **30 renglones** deben entrar (no 25 — límite real de `DetFact.FG1`,
  no del PDF) — la tabla arranca en Y=92mm y tiene que terminar antes
  del subtotal crudo en Y≈231mm, ~4.6mm por renglón.
- "Sin recuadros en detalle": la tabla de renglones no lleva grilla ni
  caja (réplica de `TableBorder = tbNone` real) — sólo una línea bajo
  el encabezado. Los recuadros de Totales (Subtotal/Descuento/Subtotal/
  IVA/[Perc.IIBB]/TOTAL) SÍ son reales (`DibujaRect` literal en el
  código) y se mantienen.
- El pie (subtotal crudo/"Son pesos"/recuadros de Totales/CAE/QR) baja
  a sus posiciones Y reales (231-284mm) en vez de quedar comprimido
  arriba — eso es lo que de paso libera el espacio para los 30
  renglones.
- El Despacho SÍ se ve como columna aparte (visualmente — internamente
  el legacy lo logra rellenando "Detalle" a ancho fijo en una fuente
  monoespaciada y pegando el código atrás, no con una columna de tabla
  real; acá se logra el mismo resultado visual dibujando el Despacho en
  su propia posición X, más simple en ReportLab).
- "Condición de Venta" es un campo real del cliente que faltaba.
- Nota de Crédito imprime "Aplicada a: Factura Nº... del..." cuando
  corresponde.
- AFIP puede aprobar un CAE CON advertencias (`Observaciones`, ya
  capturado en `ResultadoCAE.motivo` — `afip.py:554-557` — pero antes
  sólo se usaba si el comprobante era RECHAZADO); el legacy la imprime
  igual aunque haya CAE aprobado ("Se asignó CAE pero con advertencias").
- Factura/NC/ND imprimen 2 páginas completas (ORIGINAL/DUPLICADO,
  `EmiFact.frm` variable `Copias`); Recibo/Nota de Crédito Interna
  imprimen 2 copias apiladas en UNA sola página A4 (técnica real
  `MasLin`, `EmiRec.frm Sub IMPRIME2en1`/`NCInterna.frm Sub ImprimeNCI`).
- Nombre de archivo real confirmado (`EmiFact.frm:1736`):
  `TIPO & "-" & LaLetra & "-" & PtoVtaCpbte & "-" & NROFAC` — Factura
  usa el título decorado con espacios (" F A C T U R A", línea 1280),
  NC/ND usan el título plano ("NOTA DE CRÉDITO"/"NOTA DE DÉBITO").
- Carpeta real de guardado (pedido del usuario, 2026-08-20):
  `D:\\Comprobantes\\<año>\\<mes>\\<día>\\`, ahí van TODOS los tipos
  (Factura real/provisoria, Recibo, Cotización, NC, etc.).

Validado renderizando los PDF generados a imagen (`pymupdf`, sólo
herramienta de QA local) contra datos reales de las muestras del
usuario — no sólo por tests de contenido de texto.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Optional

from reportlab.graphics import renderPDF
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import PageBreak, SimpleDocTemplate, Table, TableStyle

from .decimals import format_decimal
from .fechas import formatear_fecha_corta, formatear_fecha_dd_mmm, formatear_fecha_impresion
from .repository import ETIQUETAS_TIPO_CTASCTE
from .services import AplicacionPago, PagoCheque, PagoRetencion, RenglonEmision, TotalFactura

# Texto corto de Condición de IVA para imprimir — mismas opciones que
# `CIVA_OPCIONES` (`cliente_detalle_dialog.py`, `Combo1` del legacy en
# CabFact/CabRec), sin la nota editorial de esa lista ("Inscripto", no
# "IVA Responsable Inscripto" — el legacy antepone sólo "IVA " a mano).
CONDICIONES_IVA = {
    1: "Inscripto",
    2: "Resp.No Insc.",
    3: "Cons.Final",
    4: "Exento",
    5: "Monotributo",
}

DIR_SALIDA_DEFAULT = Path("migration") / "pdf_output"
# El logo REAL que usa el legacy en los 3 `.frm` de impresión (`Picture1`,
# cargado desde `EmiFact.frx`/`EmiRec.frx`/`NCInterna.frx`, mismo tamaño
# de control en los 3: 3405x1080 twips ≈ 60x19mm) es `assets/LOGO2.jpg`
# — corregido 2026-08-20, `Logo-Alestel.png` es un logo distinto (con
# ícono circular de fondo) que sólo usa `MainMenuWindow`, no la
# impresión de comprobantes.
LOGO_ALESTEL_PATH = Path(__file__).resolve().parent.parent / "assets" / "LOGO2.jpg"

# --- Fuentes reales del legacy (2026-08-20, pedido del usuario: "debe
# quedar igual al legacy") -----------------------------------------------
# ReportLab sólo trae embebidas las 14 fuentes estándar de PDF (Helvetica,
# Times, Courier, ...) — "Arial Black"/"Lucida Console" no están entre
# ellas y hay que registrarlas a mano desde un .ttf real antes de poder
# usarlas con `setFont()`. Se embeben en `assets/fonts/` (copiadas de
# `C:\Windows\Fonts\ariblk.ttf` y del `lucon.ttf` que ya traía el propio
# repo) para que viajen con el proyecto — y con el `.exe` empaquetado —
# sin depender de que la PC de destino tenga esas fuentes instaladas.
#
# `FUENTE_TITULO`/`FUENTE_FIJA` son el punto único de verdad: si el .ttf
# no está presente (ej. checkout parcial) caen solas a Helvetica-Bold/
# Courier, las tipografías métricamente más parecidas, en vez de romper
# la generación del PDF.
_ARIAL_BLACK_PATH = Path(__file__).resolve().parent.parent / "assets" / "fonts" / "ariblk.ttf"
_LUCIDA_CONSOLE_PATH = Path(__file__).resolve().parent.parent / "assets" / "fonts" / "lucon.ttf"
_ARIAL_PATH = Path(__file__).resolve().parent.parent / "assets" / "fonts" / "arial.ttf"
_ARIAL_BOLD_PATH = Path(__file__).resolve().parent.parent / "assets" / "fonts" / "arialbd.ttf"

FUENTE_TITULO = 'Helvetica-Bold'  # membrete grande ("ARIAL BLACK" real del legacy: S.R.L., FACTURA, Letra, "O R I G I N A L", etc.)
FUENTE_FIJA = 'Courier'  # bloques de ancho fijo (CUIT/IB, Son pesos..., CAE/QR)
# El legacy también usa "Arial" (NO "Arial Black") para el resto del
# membrete de Recibo/NC Interna (dirección, teléfono) y para toda la
# fila de Totales de Factura (`EmiFact.frm Escribe "Arial", ...`, ver
# `_dibujar_pie`) — es una fuente distinta de FUENTE_TITULO, más fina.
# `FUENTE_NEGRITA` es su versión bold real (el legacy la arma con el
# flag `negrita` de `Escribe` sobre la MISMA "Arial", no cambiando a
# Arial Black) — antes esta fila caía en `FUENTE_TITULO` (Arial Black),
# visualmente mucho más pesada/ancha que un Arial en negrita real.
FUENTE_NORMAL = 'Helvetica'
FUENTE_NEGRITA = 'Helvetica-Bold'

if _ARIAL_BLACK_PATH.exists():
    pdfmetrics.registerFont(TTFont("Arial Black", str(_ARIAL_BLACK_PATH)))
    FUENTE_TITULO = "Arial Black"

if _LUCIDA_CONSOLE_PATH.exists():
    pdfmetrics.registerFont(TTFont("Lucida Console", str(_LUCIDA_CONSOLE_PATH)))
    FUENTE_FIJA = "Lucida Console"

if _ARIAL_PATH.exists():
    pdfmetrics.registerFont(TTFont("Arial", str(_ARIAL_PATH)))
    FUENTE_NORMAL = "Arial"

if _ARIAL_BOLD_PATH.exists():
    pdfmetrics.registerFont(TTFont("Arial Bold", str(_ARIAL_BOLD_PATH)))
    FUENTE_NEGRITA = "Arial Bold"


def _directorio_comprobantes(fecha: date) -> Path:
    """Carpeta real donde se archivan los comprobantes emitidos (pedido
    del usuario, 2026-08-20): `D:\\Comprobantes\\<año>\\<mes>\\<día>\\`,
    organizados por la FECHA del comprobante — ahí van todos los tipos
    (Factura real/provisoria, Recibo, Cotización, NC, etc., "etc.")."""
    return Path("D:/Comprobantes") / f"{fecha.year:04d}" / f"{fecha.month:02d}" / f"{fecha.day:02d}"


# ---------------------------------------------------------------------------
# Monto en letras — réplica EXACTA de `FCMENU.bas Function MontoEscrito`
# (línea 539), typos reales incluidos a propósito ("DIESISEIS" en vez de
# "DIECISÉIS", etc.) — el pedido del usuario es que el PDF quede IGUAL
# al legacy, no una versión corregida.
# ---------------------------------------------------------------------------

_UNIDADES = [
    "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ",
    "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO",
    "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO",
    "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE",
]
_DECENAS = ["DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
_CENTENAS = [
    "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS",
    "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS",
]


def monto_en_palabras(valor: Decimal) -> str:
    """Réplica de `MontoEscrito()` — "UN MILLON DOSCIENTOS TREINTA Y
    CUATRO MIL con 56/100", sin el "PESOS"/paréntesis (esa parte del
    legacy estaba comentada, ver línea 591 de `FCMENU.bas`; lo que
    realmente se imprime es sólo esto con "Son pesos"/"Son Pesos:"
    antepuesto por el llamador)."""
    valor = Decimal(valor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    entero = int(valor)
    centavos = int((valor - entero) * 100)

    resultado = ""
    cantidad = entero
    numero_bloques = 1
    while True:
        primer_digito = segundo_digito = tercer_digito = 0
        bloque = ""
        bloque_cero = 0
        for i in range(1, 4):
            digito = cantidad % 10
            if digito != 0:
                if i == 1:
                    bloque = " " + _UNIDADES[digito - 1]
                    primer_digito = digito
                elif i == 2:
                    if digito <= 2:
                        bloque = " " + _UNIDADES[digito * 10 + primer_digito - 1]
                    else:
                        bloque = " " + _DECENAS[digito - 1] + (" Y" if primer_digito != 0 else "") + bloque
                    segundo_digito = digito
                elif i == 3:
                    centena = "CIEN" if (digito == 1 and primer_digito == 0 and segundo_digito == 0) else _CENTENAS[digito - 1]
                    bloque = " " + centena + bloque
                    tercer_digito = digito
            else:
                bloque_cero += 1
            cantidad //= 10
            if cantidad == 0:
                break

        if numero_bloques == 1:
            resultado = bloque
        elif numero_bloques == 2:
            resultado = bloque + ("" if bloque_cero == 3 else " MIL") + resultado
        elif numero_bloques == 3:
            sufijo = " MILLON" if (primer_digito == 1 and segundo_digito == 0 and tercer_digito == 0) else " MILLONES"
            resultado = bloque + sufijo + resultado
        # Bloques 4+ (miles de millones): el legacy no los suma a ningún
        # lado (el `Select Case` sólo tiene casos 1/2/3) — se replica tal
        # cual, aunque sea un límite real del sistema original.

        numero_bloques += 1
        if cantidad == 0:
            break

    texto = resultado.strip() or "CERO"
    return f"{texto} con {centavos:02d}/100"


def _dibujar_logo(c: canvas.Canvas, x: float, y_top: float, ancho: float = 60 * mm) -> float:
    """Dibuja el logo real (`assets/LOGO2.jpg`, el que usa el legacy en
    la impresión — ver `LOGO_ALESTEL_PATH`). `ancho` default 60mm
    replica el tamaño real del control `Picture1` en los 3 `.frm`
    (3405 twips ≈ 60mm). Devuelve el alto real dibujado (para acomodar
    el texto de al lado). Si el archivo no está (ej. checkout parcial),
    no rompe — sigue sin logo."""
    if not LOGO_ALESTEL_PATH.exists():
        return 0
    lector = ImageReader(str(LOGO_ALESTEL_PATH))
    ancho_nativo, alto_nativo = lector.getSize()
    alto = ancho * (alto_nativo / ancho_nativo) if ancho_nativo else ancho * 0.3
    c.drawImage(lector, x, y_top - alto, width=ancho, height=alto, mask="auto")
    return alto


@dataclass
class DatosFacturaPDF:
    """Todo lo que necesita el PDF — ya resuelto por el llamador
    (`FacturadorWindow`), este módulo sólo dibuja.

    `cae`/`cae_vencimiento`/`qr_url` son `None` para el **boceto**
    (pedido del usuario, 2026-08-18: "al apretar emitir, muestre un
    boceto en pdf de la factura... Al apretar grabar, pregunte si se
    quiere grabar" — recién ahí, si confirma, se pide el CAE a AFIP).
    Como el CAE/QR sólo existen DESPUÉS de que AFIP aprueba, el boceto
    se genera y se muestra ANTES de tocar AFIP para nada — se dibuja un
    sello "BORRADOR" en su lugar, ver `generar_pdf_factura`."""

    letra: str
    punto_venta: int
    numero: int
    fecha: date
    cliente_codigo: int
    cliente_nombre: str
    cliente_cuit: str
    cliente_civa: int
    cliente_domicilio: str
    renglones: list[RenglonEmision]
    total: TotalFactura
    # Texto "%Descuento" por renglón (mismo orden que `renglones`, ver
    # `FacturadorWindow._descuentos_por_renglon`) — réplica de la
    # columna "% Descuento" que arma `EmiFact.frm` (líneas 1044-1052)
    # concatenando `DtoXClte.DTO1..DTO5` con "+" (ej. "10+5"). Lista
    # vacía (default) para quien todavía no la calcula.
    descuentos_renglones: list[str] = field(default_factory=list)
    cae: Optional[str] = None
    cae_vencimiento: Optional[date] = None
    qr_url: Optional[str] = None
    empresa_nombre: str = "ALESTEL SRL"
    empresa_cuit: str = "33703467909"
    en_dolares: bool = False
    # Cotización del dólar vigente al momento de emitir — sólo se usa
    # (impresa en la leyenda legal de `en_dolares`, ver `LeyenDol5` en
    # `EmiFact.frm:403`/`_dibujar_pagina_factura`) cuando `en_dolares`
    # es True; irrelevante en pesos.
    cotizacion: Decimal = Decimal("1")
    # Cotización (`CabFact.frm TipoFac=4`) reusa este mismo layout — ver
    # `CotizacionVentaService`/`CotizacionVentaWindow`. Letra "X" fija,
    # jamás pide CAE (nunca `Graba()`), documento sin validez fiscal.
    es_cotizacion: bool = False
    # Título impreso en el encabezado y usado en el nombre de archivo —
    # réplica real de `EmiFact.frm:1278-1285`: sólo Factura usa el
    # título decorado con espacios (" F A C T U R A"); NC/ND usan texto
    # plano ("NOTA DE CRÉDITO"/"NOTA DE DÉBITO").
    titulo_comprobante: str = "F A C T U R A"
    # -- Datos del cliente que faltaban (agregados 2026-08-20, hallazgo
    # real comparando contra la Factura de muestra) --
    cliente_localidad: str = ""
    cliente_cp: str = ""
    # Nombre YA RESUELTO de la Provincia (ej. "Bs. As.") — réplica de
    # `EmiFact.frm:1371` (`PCIA(Asc(CltePCIA)-65)`, ver
    # `provincias.nombre_provincia()`), el llamador ya lo resuelve.
    cliente_provincia: str = ""
    condicion_venta: str = ""
    # Observaciones/advertencias de AFIP aunque el CAE haya sido
    # APROBADO (`ResultadoCAE.motivo`, `afip.py:554-557` ya las
    # captura pero antes sólo se mostraban en el rechazo) — réplica de
    # `EmiFact.frm:1722-1725` ("Se asignó CAE pero con advertencias.
    # Motivo: ...").
    observaciones_afip: str = ""
    # Sólo Nota de Crédito: a qué Factura se está aplicando —
    # `(punto_venta, numero, fecha)` — réplica de `EmiFact.frm:1356-1362`
    # ("Aplicada a: Factura Nº... del..."). `None` si no aplica (ND, o
    # NC que no imprime esa leyenda).
    aplicada_a: Optional[tuple[int, int, date]] = None


def _nombre_archivo(datos: DatosFacturaPDF) -> str:
    """Réplica real de `EmiFact.frm:1736`: `TIPO & "-" & LaLetra & "-"
    & PtoVtaCpbte & "-" & NROFAC` — el número NO lleva ceros a la
    izquierda en el nombre de archivo (a diferencia del texto impreso
    adentro del PDF, que sí)."""
    if datos.es_cotizacion:
        return f"COTIZACIÓN-{datos.punto_venta:04d}-{datos.numero}.pdf"
    sufijo = "_BORRADOR" if datos.cae is None else ""
    return f"{datos.titulo_comprobante.strip()}-{datos.letra}-{datos.punto_venta:04d}-{datos.numero}{sufijo}.pdf"


def generar_pdf_factura(datos: DatosFacturaPDF, directorio_salida: Optional[Path] = None) -> Path:
    """Genera el PDF y lo guarda en disco (`directorio_salida`, por
    defecto `D:\\Comprobantes\\<año>\\<mes>\\<día>\\`, ver
    `_directorio_comprobantes` — pedido del usuario 2026-08-20) —
    réplica de "Resguardo y Gestión" del CLAUDE.md. Devuelve la ruta del
    archivo creado.

    2 páginas completas ORIGINAL/DUPLICADO para un comprobante real
    (réplica de `EmiFact.frm` variable `Copias`) — el boceto (sin CAE
    todavía) y la Cotización quedan en 1 sola página (no tiene sentido
    imprimir "duplicado" de algo que no se grabó/no es fiscal)."""
    directorio = Path(directorio_salida) if directorio_salida is not None else _directorio_comprobantes(datos.fecha)
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / _nombre_archivo(datos)

    c = canvas.Canvas(str(ruta), pagesize=A4)
    ancho, alto = A4

    es_boceto = datos.cae is None and not datos.es_cotizacion
    copias = [None] if (es_boceto or datos.es_cotizacion) else ["O R I G I N A L", "D U P L I C A D O"]
    for etiqueta_copia in copias:
        _dibujar_pagina_factura(c, datos, ancho, alto, etiqueta_copia, es_boceto)
        c.showPage()
    c.save()
    return ruta


def _dibujar_pagina_factura(
    c: canvas.Canvas, datos: DatosFacturaPDF, ancho: float, alto: float,
    etiqueta_copia: Optional[str], es_boceto: bool,
) -> None:
    """Coordenadas ABSOLUTAS (mm desde el borde superior/izquierdo de la
    hoja), transcriptas literales de `EmiFact.frm Sub IMPRIME_Electronica`
    (líneas 1310-1737, confirmado con el usuario 2026-08-20 que ÉSE es
    el procedimiento real — no `Sub IMPRIME()`, que es código muerto/de
    impresión en papel). `Escribe(Font, Size, Negr, Ital, Lin, Col,
    Texto)` ancla cada texto en (Col, Lin) como esquina superior-
    izquierda (no hay parámetro de alineación real en `Escribe` — todo
    el sistema dibuja "en el cursor", como una máquina de escribir) —
    acá se replica igual, `drawString` en vez de `drawRightString` salvo
    donde el propio legacy alinea a la derecha adentro de una tabla real
    (columnas numéricas de "Detalle" y los recuadros de Totales)."""

    def pos(lin_mm: float, col_mm: float) -> tuple[float, float]:
        return col_mm * mm, alto - lin_mm * mm

    def y_de(lin_mm: float) -> float:
        return alto - lin_mm * mm

    # Líneas/recuadros muy finos (pedido del usuario, 2026-08-20 — "muy
    # finas", el 1pt default de ReportLab se veía grueso comparado con
    # el original) — aplica a todo `line()`/`rect()` de acá en adelante.
    c.setLineWidth(0.35)

    # --- Logo + "S.R.L." (`DrawPicture Picture1, "10mm", "5mm"`) -----------
    # `Picture1` mide 3405x1080 twips en el `.frm` (~60x19mm) pero
    # `DrawPicture` sin ancho/alto explícitos dibuja al tamaño NATIVO del
    # bitmap, no al del control — 60mm de ancho pisaba el resto del
    # membrete (probado y descartado, 2026-08-20); 50mm deja lugar para
    # que "S.R.L." no quede pegado.
    _dibujar_logo(c, 10 * mm, y_de(5), ancho=50 * mm)
    c.setFont(FUENTE_TITULO, 20)
    # Ajuste a ojo del usuario (2026-08-20, sobre la 1ª ronda): a la
    # altura del logo (antes más arriba) y 5mm a la izquierda (antes
    # X=70 real).
    c.drawString(*pos(20, 63), "S.R.L.")

    # --- Membrete (Arial 9/10/7 bold, X=18-20) -----------------------------
    c.setFont(FUENTE_TITULO, 9)
    c.drawString(*pos(23, 20), "CINTAS Y CORREAS DE TRANSMISIÓN")
    c.drawString(*pos(26.5, 20), "AUTOMOTORES - AGRÍCOLAS - INDUSTRIALES")
    c.drawString(*pos(30, 20), "GOMAS - ARTÍCULOS - ACCESORIOS")
    c.setFont(FUENTE_TITULO, 10)
    c.drawString(*pos(35, 20), "Av. Crovara 2948 - (1766) La Tablada (BA)")
    c.drawString(*pos(39, 18), "Tel/Fax (011) 4652-1040 / 2684 / 2689 / 3080")
    c.setFont(FUENTE_TITULO, 7)
    c.drawString(*pos(41.5, 30), "email: correas@alestel.com.ar")
    c.setFont(FUENTE_NORMAL, 11)
    c.drawString(*pos(47, 18), "I.V.A. RESPONSABLE INSCRIPTO")
    if etiqueta_copia:
        c.setFont(FUENTE_TITULO, 9)
        c.drawString(*pos(45, 102), etiqueta_copia)

    # --- CUIT/IB/Partida/Imp.Internos/Fec.Inicio (Lucida→Courier, X=140) --
    c.setFont(FUENTE_FIJA, 8)
    c.drawString(*pos(31, 140), f"CUIT: {datos.empresa_cuit[:2]}-{datos.empresa_cuit[2:10]}-{datos.empresa_cuit[10:]}")
    c.drawString(*pos(34, 140), "IB Conv.Mult.: 901 33-70346790-9")
    c.drawString(*pos(37, 140), "Partida Municipal: 1985")
    c.drawString(*pos(40, 140), "Imp.Internos: No Responsable")
    c.drawString(*pos(43, 140), "Fec.Inicio Act.: 01-09-1999")

    # --- Letra (bajada 10mm + recuadro) + Código AFIP -----------------------
    # Ajuste a ojo del usuario (2026-08-20): la Letra 10mm más abajo que
    # la posición real, y ahora DENTRO de un recuadro (el original la
    # tenía suelta en esta réplica; el legacy sí la encierra en su propio
    # `DibujaRect`, ver docstring del módulo).
    c.rect(105 * mm, y_de(31), 17 * mm, 17 * mm)
    c.setFont(FUENTE_TITULO, 30)
    c.drawString(*pos(26, 110), datos.letra)
    c.setFont(FUENTE_NORMAL, 7)
    codigo_afip_doc = {"F A C T U R A": "01", "NOTA DE CRÉDITO": "03", "NOTA DE DÉBITO": "02"}.get(
        datos.titulo_comprobante.strip(), "01" if datos.letra == "A" else "06"
    )
    c.drawString(*pos(30, 108), f"Código {codigo_afip_doc}")

    # --- Título + Número + Fecha ---------------------------------------
    c.setFont(FUENTE_TITULO, 20)
    titulo = "COTIZACIÓN" if datos.es_cotizacion else datos.titulo_comprobante
    # Ajuste a ojo del usuario: FACTURA 5mm más abajo y 5mm más a la
    # derecha que la posición real (Y=5,X=130).
    c.drawString(*pos(17, 140), titulo)
    c.setFont(FUENTE_NORMAL, 14)
    c.drawString(*pos(22, 150), f"{datos.punto_venta:04d}-{datos.numero:08d}")
    c.setFont(FUENTE_NORMAL, 10)
    c.drawString(*pos(27, 146), formatear_fecha_impresion(datos.fecha))
    if datos.en_dolares:
        c.drawString(*pos(31, 146), "Comprobante en DÓLARES")

    # --- Cliente + Condición de Venta, en un recuadro de lado a lado ------
    # Pedido del usuario, 2026-08-20 (2ª ronda de ajustes a ojo): antes
    # suelto, ahora en un recuadro que va de margen a margen.
    c.rect(7 * mm, y_de(77), (ancho - 14 * mm), (77 - 47) * mm)
    c.setFont(FUENTE_NORMAL, 10)
    c.drawString(*pos(53, 10), f"{datos.cliente_nombre}         ({datos.cliente_codigo})")
    c.drawString(*pos(58, 10), datos.cliente_domicilio)
    cp_loc_pcia = f"({datos.cliente_cp})  {datos.cliente_localidad}  {datos.cliente_provincia}".strip()
    c.drawString(*pos(63, 10), cp_loc_pcia)
    c.drawString(*pos(68, 10), f"IVA {CONDICIONES_IVA.get(datos.cliente_civa, '—')}     CUIT : {datos.cliente_cuit or 's/d'}")
    if datos.condicion_venta:
        c.drawString(*pos(73, 10), f"Condición de Venta : {datos.condicion_venta}")
    if datos.aplicada_a:
        # Sólo Nota de Crédito (`EmiFact.frm:1356-1362`), X=145-160.
        ptovta_ap, numero_ap, fecha_ap = datos.aplicada_a
        c.setFont(FUENTE_TITULO, 10)
        c.drawString(*pos(63, 160), "Aplicada a:")
        c.drawString(*pos(68, 145), f"Factura Nº {ptovta_ap:04d}-{numero_ap:08d}")
        c.drawString(*pos(73, 150), f"del {fecha_ap.strftime('%d/%m/%Y')}")

    # --- Tabla de renglones: títulos en un recuadro de lado a lado --------
    # Cant. | Detalle | Despacho | % Descuento | Precio Unit. | Importe.
    # Pedido del usuario (2ª ronda): sacar la línea divisoria inferior y
    # meter los títulos en un recuadro de margen a margen (mismo tamaño
    # 10pt bold que usa el legacy — antes 8pt, por eso se veían corridos
    # a la izquierda respecto de los valores de abajo). Los RENGLONES en
    # sí siguen "sin recuadros" (pedido de la 1ª ronda) — el recuadro es
    # sólo para la fila de títulos.
    c.rect(7 * mm, y_de(86), (ancho - 14 * mm), (86 - 78) * mm)
    c.setFont(FUENTE_TITULO, 10)
    c.drawString(*pos(83, 12), "Cant.")
    c.drawString(*pos(83, 56), "Detalle")
    c.drawString(*pos(83, 83), "Despacho")
    c.drawString(*pos(83, 109), "%Descuento")
    c.drawString(*pos(83, 145), "Precio Unit.")
    c.drawString(*pos(83, 178), "Importe")

    # 30 renglones reales (pedido del usuario — antes 25) tienen que
    # entrar entre Y=92 (arranque real de la tabla) y Y≈231 (subtotal
    # crudo) — 139mm de espacio ÷ 30 = ~4.6mm/renglón.
    ALTO_RENGLON_MM = 4.6
    c.setFont(FUENTE_NORMAL, 8)
    descuentos_renglones = datos.descuentos_renglones or []
    lin = 92.0
    for indice, renglon in enumerate(datos.renglones[:30]):
        cantidad = renglon.cantidad_unidades if renglon.cantidad_unidades > 0 else renglon.mtr
        desc_texto = descuentos_renglones[indice] if indice < len(descuentos_renglones) else ""

        c.drawRightString(*pos(lin, 26), format_decimal(cantidad))
        c.drawString(*pos(lin, 27), renglon.descripcion.strip()[:32])
        if renglon.nrodesp_elegido:
            c.drawString(*pos(lin, 83), renglon.nrodesp_elegido)
        c.drawRightString(*pos(lin, 142), desc_texto)
        c.drawRightString(*pos(lin, 165), format_decimal(renglon.precio_unitario))
        c.drawRightString(*pos(lin, 198), format_decimal(renglon.importe))
        lin += ALTO_RENGLON_MM

    # --- Leyenda legal "en Dólares" (réplica EmiFact.frm:1085-1094, sólo
    # si `en_dolares`) — 5 líneas fijas (Lucida Console) + 3 etiquetas
    # "U$S" (Arial bold) sobre las cajas de Subtotal/Subtotal neto/TOTAL
    # más abajo. Texto SIN tildes tal cual el legacy (charset ANSI de
    # VB6 de la época, no un error de tipeo nuestro — mismo criterio que
    # `monto_en_palabras`, "igual al legacy" incluye sus mayúsculas sin
    # acentuar).
    if datos.en_dolares:
        # Réplica fiel: en el legacy estas 5 líneas están en las mismas
        # posiciones ABSOLUTAS que el área de renglones (Y=92-231) — si
        # una Factura en dólares tuviera más de ~22 renglones cargados,
        # la leyenda se superpondría a los últimos, igual que en el
        # original (no es un bug nuestro, es una limitación real del
        # legacy que no corresponde "corregir" acá).
        c.setFont(FUENTE_FIJA, 7)
        c.drawString(*pos(195, 10), "LA PRESENTE FACTURA SE ABONARA EN DOLARES ESTADOUNIDENSES BILLETES O EN LA CANTIDAD DE PESOS")
        c.drawString(*pos(199, 10), "SUFICIENTES PARA CANCELAR LA SUMA TOTAL EN DOLARES  AL TIPO DE CAMBIO VENDEDOR DEL B.N.A. DEL")
        c.drawString(*pos(203, 10), "MERCADO OFICIAL QUE RIJA PARA LAS IMPORTACIONES DE BIENES (O AQUEL QUE EN EL FUTURO LO ")
        c.drawString(*pos(207, 10), "SUSTITUYA) DEL DIA ANTERIOR AL EFECTIVO PAGO.---------------------------------------------")
        c.drawString(*pos(211, 10), f"A LOS EFECTOS IMPOSITIVOS EL TIPO DE CAMBIO APLICADO ES u$s 1= $ {format_decimal(datos.cotizacion)}")
        c.setFont(FUENTE_NEGRITA, 8)
        for x_mm in (22, 72, 188):
            c.drawString(*pos(250.5, x_mm), "U$S")

    # --- Subtotal crudo (sin etiqueta, réplica EmiFact.frm:1637, Y=231) ---
    c.setFont(FUENTE_TITULO, 9)
    c.drawRightString(*pos(236, 195), format_decimal(datos.total.bruto))

    # --- Son pesos... (réplica MontoEscrito, Y=236/240/244, X=30) ---------
    c.setFont(FUENTE_FIJA, 7)
    prefijo_moneda = "Son dólares " if datos.en_dolares else "Son pesos "
    texto_letras = prefijo_moneda + monto_en_palabras(datos.total.total)
    for i, inicio in enumerate((0, 70, 140)):
        c.drawString(*pos(236 + i * 4, 30), texto_letras[inicio : inicio + 70])
        
    # Linea totales    
    c.line(7 * mm, 50 * mm, 198 * mm, 50 * mm )  

    # --- Totales en recuadros (réplica EmiFact.frm:1619-1637, Y=255/260) --
    # Negrita real por etiqueta (`Escribe` real: Subtotal/Descuento/IVA
    # NO son negrita — sólo el 2º "Subtotal" [neto] y "TOTAL" sí,
    # `EmiFact.frm:1630-1635`) — antes las 5 tenían el mismo estilo.
    etiquetas_totales = [
        ("Subtotal", datos.total.bruto, False, False),
        ("Descuento", datos.total.descuento, False, False),
        ("Subtotal", datos.total.neto_gravado, True, False),
        ("IVA Insc. 21%", datos.total.iva, False, False),
    ]
    if datos.total.percepcion_iibb > 0:
        etiquetas_totales.append(("Perc. IIBB", datos.total.percepcion_iibb, False, False))
    etiquetas_totales.append(("        TOTAL", datos.total.total, True, True))

    # Posiciones/anchos reales de las 6 cajas (`DibujaRect 260,7,268,37` /
    # `260,39,268,68` / `260,70,268,100` / `260,102,268,132` /
    # `260,134,268,164` [IIBB, condicional] / `260,166,268,202` [TOTAL,
    # sombreado, más ancha]) — la etiqueta va ARRIBA de la caja (Y=255,
    # afuera), el valor ADENTRO (Y≈265, una sola línea, caja de 8mm).
    x_cajas = [7, 39, 70, 102, 166] if len(etiquetas_totales) == 5 else [7, 39, 70, 102, 134, 166]
    y_caja_bottom = y_de(268)
    alto_caja = 8 * mm
    # `es_total` viaja YA CALCULADO en la propia tupla (2026-08-22, bug
    # real reportado por el usuario — "el total de la factura impresa
    # tampoco está en negrita"): antes se re-derivaba acá comparando
    # `etiqueta == "      TOTAL"` (11 espacios) contra el valor real
    # armado arriba con "        TOTAL" (13 espacios, línea de arriba)
    # — un typo de espacios que hacía que la comparación NUNCA diera
    # `True`, así que el importe del TOTAL jamás se dibujó en negrita
    # ni la caja se ensanchó/sombreó (confirmado extrayendo la fuente
    # real del PDF: el importe salía en `ArialMT` común, no
    # `Arial-BoldMT` — sólo la ETIQUETA "TOTAL" de arriba, que usa el
    # flag `negrita` — aparte, no roto — se veía bien, por eso a
    # simple vista parecía correcto).
    for (etiqueta, valor, negrita, es_total), x_mm in zip(etiquetas_totales, x_cajas):
        ancho_caja_mm = 36 if es_total else 30
        x0 = x_mm * mm
        if es_total:
            c.setFillColor(colors.HexColor("#e0e0e0"))
        c.rect(x0, y_caja_bottom, ancho_caja_mm * mm, alto_caja, fill=1 if es_total else 0, stroke=1)
        if es_total:
            c.setFillColor(colors.black)
        # Tamaño 10 real (`Escribe ...,10,...`) — antes 8, por eso los
        # títulos se veían más chicos/corridos a la izquierda que los
        # valores de abajo (2ª ronda de ajustes a ojo).
        c.setFont(FUENTE_NEGRITA if negrita else FUENTE_NORMAL, 10)
        c.drawString(*pos(258, x_mm + 2), etiqueta)
        # El TOTAL final a tamaño 10 (pedido del usuario, 2026-08-22 —
        # antes 9, igual que el resto de las cajas) — la caja de TOTAL ya
        # venía más ancha que las demás (36mm vs 30mm) así que a tamaño
        # 10 el importe sigue entrando sobrado (medido: hasta "$
        # 1.234.567,89" ocupa ~24mm), no hizo falta agrandarla más.
        c.setFont(FUENTE_NEGRITA if es_total else FUENTE_NORMAL, 10 if es_total else 9)
        c.drawRightString(*pos(266, x_mm + ancho_caja_mm - 2), format_decimal(valor))

    # --- CAE + Vencimiento + QR (Y=270/274/269, X=132/137/7) ---------------
    c.setFont(FUENTE_FIJA, 10)
    if datos.es_cotizacion:
        # Nunca lleva CAE/QR (TipoFac=4 jamás llama a ConectaAFIP/Graba
        # en el legacy) ni el sello "BORRADOR" — no es un boceto de un
        # comprobante fiscal pendiente, es un documento sin validez
        # fiscal por diseño.
        c.drawString(*pos(270, 40), "Documento sin validez fiscal — no es Factura.")
        c.drawString(*pos(275, 40), "Precios sujetos a modificación sin previo aviso.")
    elif datos.cae is not None and datos.cae_vencimiento is not None and datos.qr_url is not None:
        c.drawString(*pos(281, 40), f"C.A.E.: {datos.cae}")
        c.drawString(*pos(285, 40), f"Fec.Vto.: {datos.cae_vencimiento.strftime('%d/%m/%Y')}")
        if datos.observaciones_afip:
            # Réplica de EmiFact.frm:1722-1725 — AFIP puede aprobar CON
            # advertencias, no sólo rechazar.
            c.setFont(FUENTE_NORMAL, 7)
            c.drawString(*pos(290, 40), f"Se asignó CAE pero con advertencias. Motivo: {datos.observaciones_afip}")

        qr_widget = qr.QrCodeWidget(datos.qr_url)
        bounds = qr_widget.getBounds()
        ancho_qr = bounds[2] - bounds[0]
        alto_qr = bounds[3] - bounds[1]
        tamano = 25 * mm
        dibujo = Drawing(tamano, tamano, transform=[tamano / ancho_qr, 0, 0, tamano / alto_qr, 0, 0])
        dibujo.add(qr_widget)
        renderPDF.draw(dibujo, c, 7 * mm, y_de(269) - tamano)
    else:
        c.drawString(*pos(280, 40), "CAE: (pendiente — todavía no se solicitó a AFIP)")
        _dibujar_sello_borrador(c, ancho, alto)


def _dibujar_sello_borrador(c: canvas.Canvas, ancho: float, alto: float) -> None:
    """Sello diagonal "BORRADOR" — para que el boceto que se muestra
    antes de pedir el CAE (`FacturadorWindow._on_emitir`) nunca se
    confunda con el comprobante fiscal real."""
    c.saveState()
    c.setFillColorRGB(0.85, 0.2, 0.2, alpha=0.35)
    c.setFont(FUENTE_TITULO, 60)
    c.translate(ancho / 2, alto / 2)
    c.rotate(35)
    c.drawCentredString(0, 0, "BORRADOR")
    c.restoreState()


# ---------------------------------------------------------------------------
# PDF de Recibo — EmisionReciboService, sin CAE/QR (Recibo no es
# comprobante fiscal electrónico, ver docstring de EmisionReciboService).
#
# Réplica visual real (`EmiRec.frm Sub IMPRIME2en1`, 2026-08-20): 2
# copias ORIGINAL/DUPLICADO apiladas en UNA sola hoja A4 (técnica
# `MasLin`), no 2 páginas separadas ni media hoja — mismo "Documento no
# válido como Factura" que usa `NCInterna.frm` (comparten letterhead,
# ver `_dibujar_membrete_no_fiscal`).
# ---------------------------------------------------------------------------


@dataclass
class DatosReciboPDF:
    """Todo lo que necesita el PDF del Recibo — ya resuelto por el
    llamador (`ReciboWindow`), este módulo sólo dibuja."""

    numero: int
    fecha: date
    correlativo: int
    cliente_codigo: int
    cliente_nombre: str
    cliente_cuit: str
    aplicaciones: list[AplicacionPago]
    anticipo: Decimal
    importe_efectivo: Decimal
    cheques: list[PagoCheque]
    retenciones: list[PagoRetencion]
    total_pago: Decimal
    cliente_domicilio: str = ""
    cliente_cp: str = ""
    cliente_localidad: str = ""
    # Código de 1 letra CRUDO (no el nombre resuelto) — réplica fiel de
    # `EmiRec.frm:571` (`ClteLoc & " " & CltePCIA`, sin resolver el
    # array `PCIA()` como sí hace Factura) — mismo dato, 2 templates
    # distintos con fidelidad distinta, no es un error acá.
    cliente_provincia_codigo: str = ""
    cliente_civa: int = 0
    empresa_nombre: str = "ALESTEL SRL"
    empresa_cuit: str = "33703467909"


def _nombre_archivo_recibo(datos: DatosReciboPDF) -> str:
    """Réplica real: `Recibo-X-0001-{numero}.pdf` (mismo patrón
    `TIPO-LETRA-PTOVTA-NUMERO`, Letra "X" fija — Recibo nunca es
    fiscal, `LaLetra = "X"` en `EmiRec.frm`)."""
    return f"Recibo-X-0001-{datos.numero}.pdf"


def _dibujar_membrete_no_fiscal(
    c: canvas.Canvas, ancho: float, alto: float, y_top_mm: float, letra: str, titulo: str, numero_texto: str, fecha: date,
) -> None:
    """Letterhead compartido por Recibo/Nota de Crédito Interna/
    Cotización — réplica real (`EmiRec.frm`/`NCInterna.frm`, ambos
    idénticos en esta parte): logo + "S.R.L." + recuadro con Letra +
    "Documento no válido como Factura" + título+número + CUIT/IB/
    Partida. Coordenadas en mm DESDE ARRIBA (`y_top_mm`, réplica del
    sistema real `Escribe(Lin, Col)` del legacy, MasLin ya sumado por
    el llamador para la 2ª copia)."""

    # Coordenadas reales absolutas (mm desde el borde izq./superior de la
    # página, réplica literal de `EmiRec.frm Sub IMPRIME2en1` — el
    # sistema real `Escribe(Lin, Col)` ya mide así, sin un margen propio
    # aparte, `IMPRE.VP1.MarginLeft = "1mm"`).
    def pos(lin_mm: float, col_mm: float) -> tuple[float, float]:
        return col_mm * mm, alto - (y_top_mm + lin_mm) * mm

    def y_pt(lin_mm: float) -> float:
        return alto - (y_top_mm + lin_mm) * mm

    _dibujar_logo(c, 10 * mm, y_pt(5), ancho=50 * mm)
    c.setFont(FUENTE_TITULO, 20)
    c.drawString(*pos(19.5, 62), "S.R.L.")

    # Recuadro grande + 4 líneas finas que dividen — réplica real
    # (`EmiRec.frm`/`NCInterna.frm`: `DibujaRect 5, 7, 130, 200, 30, 30` +
    # `DibujaLine` en Lin 31/46/53/105, mismas coordenadas para ambos
    # documentos). Sistema real Lin=fila/Col=columna en mm desde arriba —
    # convertido acá con `pos()`/`y_pt()` (el ancho de línea 0.35pt "fino"
    # ya lo dejó puesto el llamador).
    
    c.rect(6 * mm, y_pt(129), 192 * mm, 125 * mm)
    c.line(6 * mm, y_pt(30), 198 * mm, y_pt(30))    # cabecera de títulos
    c.line(6 * mm, y_pt(46), 198 * mm, y_pt(46))    # tít. de detalle y cabecera
    c.line(6 * mm, y_pt(56), 198 * mm, y_pt(56))    # detalle de títulos
    c.line(6 * mm, y_pt(103), 198 * mm, y_pt(103))  # totales de detalle

    c.setFont(FUENTE_TITULO, 24)
    c.drawCentredString(105.5 * mm, y_pt(17), letra)

    # Tamaño de fuente adaptado al largo del título ("Recibo Oficial" vs
    # "Nota de Crédito Interna", bastante más largo) — número SIEMPRE en
    # su propia línea debajo (no al lado, que es lo que colisionaba con
    # títulos largos).
    tamano_titulo = 15 if len(titulo) <= 16 else 12
    c.setFont(FUENTE_TITULO, tamano_titulo)
    c.drawString(*pos(14, 140), titulo)
    c.setFont(FUENTE_NORMAL, 12)
    c.drawString(*pos(19.5, 146), numero_texto)
    c.setFont(FUENTE_NORMAL, 6.5)
    c.drawString(*pos(19, 98), "Documento no")
    c.drawString(*pos(22, 94), "válido como Factura")

    c.setFont(FUENTE_NORMAL, 9)
    c.drawRightString(ancho - 28 * mm, y_pt(8), f"Buenos Aires, {formatear_fecha_impresion(fecha)}")

    c.setFont(FUENTE_NORMAL, 9)
    c.drawString(*pos(23, 14), "Av. Crovara 2948 - (1766) La Tablada (BA)")
    c.drawString(*pos(27, 14), "Tel. (011) 4652-1040 / 2684 / 2689 / 3080")

    c.setFont(FUENTE_NORMAL, 7)
    c.drawString(*pos(23.5, 135), "Imp.Int.: No Resp.   I.B.: 901 33-70346790-9")
    c.drawString(*pos(27, 135), "Part.Munic.: 1985   CUIT: 33-70346790-9")
   
def generar_pdf_recibo(datos: DatosReciboPDF, directorio_salida: Optional[Path] = None) -> Path:
    """Genera el PDF del Recibo — 2 copias (Original/Duplicado)
    apiladas en UNA hoja A4 (réplica real, ver docstring de arriba).
    Devuelve la ruta.

    **Orden Vista Previa → Grabar** (decisión del usuario, 2026-08-15):
    este generador se llama ANTES de `EmisionReciboService.
    emitir_recibo()` (para la vista previa) — el número que lleva el
    PDF es el que TENDRÍA el comprobante, no necesariamente uno ya
    persistido. Ver `ReciboWindow._on_emitir`."""
    directorio = Path(directorio_salida) if directorio_salida is not None else _directorio_comprobantes(datos.fecha)
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / _nombre_archivo_recibo(datos)

    c = canvas.Canvas(str(ruta), pagesize=A4)
    ancho, alto = A4

    for indice_copia, etiqueta in enumerate(("Original", "Duplicado")):
        offset_mm = indice_copia * 140  # réplica real de MasLin=145
        _dibujar_copia_recibo(c, datos, ancho, alto, offset_mm, etiqueta)
    c.showPage()
    c.save()
    return ruta


def _truncar_a_ancho(texto: str, fuente: str, tamano: float, ancho_max_pt: float) -> str:
    """Recorta `texto` (con "…" al final) para que entre en `ancho_max_pt`
    puntos con esa fuente/tamaño — usado para el nombre del Banco en el
    Recibo (bug real, 2026-08-22: un nombre de Banco largo, ej. "AMERICAN
    EXPRESS BANK LTD. SOC", pisaba la columna Fecha de al lado)."""
    if stringWidth(texto, fuente, tamano) <= ancho_max_pt:
        return texto
    while texto and stringWidth(texto + "…", fuente, tamano) > ancho_max_pt:
        texto = texto[:-1]
    return (texto + "…") if texto else ""


def _dibujar_copia_recibo(
    c: canvas.Canvas, datos: DatosReciboPDF, ancho: float, alto: float, offset_mm: float, etiqueta_copia: str
) -> None:
    def y_pt(lin_mm: float) -> float:
        return alto - (offset_mm + lin_mm) * mm

    c.setLineWidth(0.1)  # más finas todavía, pedido del usuario 2026-08-22 (0.35 seguía viéndose grueso)
    _dibujar_membrete_no_fiscal(
        c, ancho, alto, offset_mm, "X", "Recibo Oficial", f"0001-{datos.numero:08d}", datos.fecha,
    )
    # Línea divisoria Comprobantes Cancelados / Valores Recibidos — sólo en
    # Recibo (NC Interna no tiene columna de "Valores Recibidos", réplica
    # real `DibujaLine 46, 95, 105, 95` de `EmiRec.frm`, ausente en
    # `NCInterna.frm`).
    c.line(95 * mm, y_pt(46), 95 * mm, y_pt(103))

    margen_x = 8 * mm
    y = y_pt(34)
    c.setFont(FUENTE_TITULO, 9)
    c.drawString(margen_x, y, f"Sr./es: ({datos.cliente_codigo}) {datos.cliente_nombre}")
    c.setFont(FUENTE_NORMAL, 9)
    y = y_pt(38.5)
    c.drawString(margen_x, y, datos.cliente_domicilio)
    y = y_pt(43)
    c.drawString(margen_x, y, f"({datos.cliente_cp}) {datos.cliente_localidad} {datos.cliente_provincia_codigo}".strip())
    c.setFont(FUENTE_NORMAL, 9)
    c.drawString(
        margen_x + 92 * mm, y,
        f"IVA: {CONDICIONES_IVA.get(datos.cliente_civa, '—')}     CUIT : {datos.cliente_cuit or 's/d'}",
    )

    # --- Comprobantes Cancelados / Valores Recibidos (2 columnas) --------
    y = y_pt(49.5)
    c.setFont(FUENTE_TITULO, 8)
    c.drawString(margen_x + 30, y, "Comprobantes Cancelados")
    c.drawString(margen_x + 125 * mm, y, "Valores Recibidos")
    y = y_pt(54)
    c.setFont(FUENTE_NORMAL, 7)
    c.drawString(margen_x, y, "Tipo")
    c.drawString(margen_x + 15 * mm, y, "Nro.")
    c.drawString(margen_x + 40 * mm, y, "Fecha")
    c.drawRightString(margen_x + 82 * mm, y, "Importe")
    c.drawString(margen_x + 90 * mm, y, "Nro./Tipo")
    c.drawString(margen_x + 115 * mm, y, "Bco./Detalle")
    c.drawString(margen_x + 161 * mm, y, "Fecha")
    c.drawRightString(margen_x + 186 * mm, y, "Importe")

    c.setFont(FUENTE_NORMAL, 7)
    y_izq = y - 8
    for aplicacion in datos.aplicaciones:
        y_izq -= 8
        if y_izq < y_pt(100):
            break
        comprobante = aplicacion.comprobante
        etiqueta = ETIQUETAS_TIPO_CTASCTE.get(comprobante.TIPO, "—")
        fecha_txt = formatear_fecha_dd_mmm(comprobante.FECHA) if comprobante.FECHA else "—"
        c.drawString(margen_x, y_izq, etiqueta)
        c.drawString(margen_x + 15 * mm, y_izq, str(comprobante.CPBTE))
        c.drawString(margen_x + 35 * mm, y_izq, fecha_txt)
        c.drawRightString(margen_x + 86 * mm, y_izq, format_decimal(aplicacion.importe_aplicado))
    if datos.anticipo > 0:
        y_izq -= 8
        c.drawString(margen_x, y_izq, "Pago a Cuenta (Anticipo)")
        c.drawRightString(margen_x + 86 * mm, y_izq, format_decimal(datos.anticipo))

    y_der = y - 8
    # 4to campo = fecha del cheque (`fecha_vencimiento` — réplica real,
    # `DetPago.frm:905` graba esa misma columna del FG1 con el comentario
    # explícito "Fec.Vto.", y `EmiRec.frm:666` la imprime tal cual) —
    # vacía para Efectivo/Retención, que no tienen una fecha propia.
    medios: list[tuple[str, str, str, Decimal]] = []
    if datos.importe_efectivo > 0:
        medios.append(("* Efectivo *", "", "", datos.importe_efectivo))
    for pago_cheque in datos.cheques:
        tipo_ch = "ECh" if pago_cheque.tipo_cheque == 1 else ""
        fecha_cheque = formatear_fecha_dd_mmm(pago_cheque.fecha_vencimiento) if pago_cheque.fecha_vencimiento else ""
        # `pago_cheque.banco` ya viene con el NOMBRE resuelto (no el
        # código) — bug real reportado por el usuario (2026-08-22):
        # imprimía el código ("Bco. xx"), resuelto por el llamador
        # (`ReciboWindow._cheques_para_pdf`, `pdf.py` no tiene acceso a
        # la tabla de Bancos).
        medios.append((
            f"{tipo_ch} {pago_cheque.nro_cheque}".strip(),
            f"Bco. {pago_cheque.banco}",
            fecha_cheque,
            pago_cheque.importe,
        ))
    for retencion in datos.retenciones:
        medios.append((retencion.concepto or f"Tipo {retencion.tipreg}", "", "", retencion.importe))
    for texto_izq, texto_medio, fecha_txt, importe in medios:
        y_der -= 8
        if y_der < y_pt(100):
            break
        # Nro./Tipo ajustado a la derecha (pedido del usuario, 2026-08-22)
        # — pegado contra el borde izquierdo de la columna "Bco./Detalle".
        # "Bco./Detalle" ensanchada a 46mm (era 35mm — un nombre de Banco
        # real, ej. "BANCO HIPOTECARIO S.A.", pisaba la Fecha de al lado)
        # a costa de la columna Fecha (le sobraba de punta a punta, una
        # fecha corta entra sobrada en 25mm) — igual, algunos Bancos
        # reales son más largos todavía (ej. "AMERICAN EXPRESS BANK LTD.
        # SOC", 30 caracteres) y siguen sin entrar: se recorta con "…" al
        # ancho real disponible como último recurso (`_truncar_a_ancho`).
        c.drawRightString(margen_x + 113 * mm, y_der, texto_izq)
        c.drawString(margen_x + 115 * mm, y_der, _truncar_a_ancho(texto_medio, FUENTE_NORMAL, 7, 44 * mm))
        c.drawString(margen_x + 161 * mm, y_der, fecha_txt)
        c.drawRightString(margen_x + 188 * mm, y_der, format_decimal(importe))

    # --- Totales + firma ---------------------------------------------------
    y = y_pt(101)
    c.setFont(FUENTE_NORMAL, 8)
    c.drawString(margen_x, y, f"Total Cancelado: {format_decimal(datos.total_pago)}")
    total_valores = sum(
        [datos.importe_efectivo] + [p.importe for p in datos.cheques] + [r.importe for r in datos.retenciones],
        Decimal("0"),
    )
    c.drawString(margen_x + 90 * mm, y, f"Total Valores: {format_decimal(total_valores)}")

    y = y_pt(107)
    c.setFont(FUENTE_TITULO, 10)
    c.drawString(margen_x, y, f"Total: $ {format_decimal(datos.total_pago)}")
    c.setFont(FUENTE_NORMAL, 8)
    c.drawString(margen_x + 40 * mm, y, f"Son Pesos: {monto_en_palabras(datos.total_pago)}")

    y = y_pt(120)
    c.drawString(ancho - margen_x - 48 * mm, y, "___________________")
    c.drawCentredString(ancho - margen_x - 32 * mm, y_pt(125), "Por Alestel SRL")

    c.setFont(FUENTE_NORMAL, 7)
    c.drawString(margen_x, y_pt(125), "Este pago no implica cancelación de intereses por pago de facturas fuera de término")

    c.setFont(FUENTE_NORMAL, 8)
    c.drawRightString(ancho - margen_x - 20, y_pt(133), etiqueta_copia)

    c.setFont(FUENTE_NORMAL, 6)
    c.drawString(margen_x, y_pt(133), f"Correlativo: {datos.correlativo}")


# ---------------------------------------------------------------------------
# PDF de Nota de Crédito Interna (`NCInterna.frm`) — documento SIN
# validez fiscal (Letra "X" fija, nunca pasa por AFIP, ver
# `EmisionNotaCreditoInternaService`), cancela toda la deuda pendiente
# de un Cliente de una sola vez. Mismo letterhead que Recibo (comparten
# `.frm`s casi idénticos en esta parte) — 2 copias apiladas en 1 hoja.
# ---------------------------------------------------------------------------


@dataclass
class DatosNotaCreditoInternaPDF:
    numero: int
    fecha: date
    cliente_codigo: int
    cliente_nombre: str
    cliente_cuit: str
    motivo_texto: str
    nota: str
    comprobantes_cancelados: list[AplicacionPago]
    total_cancelado: Decimal
    cliente_domicilio: str = ""
    cliente_cp: str = ""
    cliente_localidad: str = ""
    cliente_provincia_codigo: str = ""
    cliente_civa: int = 0
    empresa_nombre: str = "ALESTEL SRL"
    empresa_cuit: str = "33703467909"


def _nombre_archivo_nci(datos: DatosNotaCreditoInternaPDF) -> str:
    return f"NOTA DE CREDITO INTERNA-X-0001-{datos.numero}.pdf"


def generar_pdf_nota_credito_interna(
    datos: DatosNotaCreditoInternaPDF, directorio_salida: Optional[Path] = None
) -> Path:
    """Réplica de `NCInterna.frm Sub ImprimeNCI` — documento interno que
    cancela deuda, sin CAE/QR (no es un comprobante fiscal). 2 copias
    apiladas en 1 hoja A4, mismo letterhead que Recibo."""
    directorio = Path(directorio_salida) if directorio_salida is not None else _directorio_comprobantes(datos.fecha)
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / _nombre_archivo_nci(datos)

    c = canvas.Canvas(str(ruta), pagesize=A4)
    ancho, alto = A4

    for indice_copia, etiqueta in enumerate(("Original", "Duplicado")):
        offset_mm = indice_copia * 145
        _dibujar_copia_nci(c, datos, ancho, alto, offset_mm, etiqueta)
    c.showPage()
    c.save()
    return ruta


def _dibujar_copia_nci(
    c: canvas.Canvas, datos: DatosNotaCreditoInternaPDF, ancho: float, alto: float, offset_mm: float, etiqueta_copia: str
) -> None:
    def y_pt(lin_mm: float) -> float:
        return alto - (offset_mm + lin_mm) * mm

    c.setLineWidth(0.35)  # líneas/recuadros finos, pedido del usuario 2026-08-20
    _dibujar_membrete_no_fiscal(
        c, ancho, alto, offset_mm, "X", "Nota de Crédito Interna", f"0001-{datos.numero:06d}", datos.fecha,
    )

    margen_x = 8 * mm
    c.setFont(FUENTE_TITULO, 9)
    c.drawString(margen_x, y_pt(32), f"Sr./es: ({datos.cliente_codigo}) {datos.cliente_nombre}")
    c.setFont(FUENTE_NORMAL, 9)
    c.drawString(margen_x, y_pt(37), datos.cliente_domicilio)
    c.drawString(margen_x, y_pt(42), f"({datos.cliente_cp}) {datos.cliente_localidad} {datos.cliente_provincia_codigo}".strip())
    c.setFont(FUENTE_TITULO, 9)
    c.drawString(
        margen_x + 92 * mm, y_pt(42),
        f"IVA: {CONDICIONES_IVA.get(datos.cliente_civa, '—')}     CUIT : {datos.cliente_cuit or 's/d'}",
    )

    y = y_pt(46)
    c.setFont(FUENTE_TITULO, 8)
    c.drawString(margen_x, y, "Comprobantes Cancelados")
    y -= 10
    c.setFont(FUENTE_TITULO, 7)
    c.drawString(margen_x, y, "Tipo")
    c.drawString(margen_x + 20 * mm, y, "Nro.")
    c.drawString(margen_x + 45 * mm, y, "Fecha")
    c.drawRightString(margen_x + 130 * mm, y, "Importe")
    c.drawRightString(margen_x + 175 * mm, y, "Cancelado")

    c.setFont(FUENTE_NORMAL, 7.5)
    for aplicacion in datos.comprobantes_cancelados:
        y -= 8
        if y < y_pt(100):
            break
        comprobante = aplicacion.comprobante
        etiqueta = ETIQUETAS_TIPO_CTASCTE.get(comprobante.TIPO, "—")
        fecha_txt = formatear_fecha_dd_mmm(comprobante.FECHA) if comprobante.FECHA else "—"
        c.drawString(margen_x, y, etiqueta)
        c.drawString(margen_x + 20 * mm, y, str(comprobante.CPBTE))
        c.drawString(margen_x + 45 * mm, y, fecha_txt)
        c.drawRightString(margen_x + 130 * mm, y, format_decimal(comprobante.IMPTE or Decimal("0")))
        c.drawRightString(margen_x + 175 * mm, y, format_decimal(aplicacion.importe_aplicado))

    y = y_pt(101)
    c.setFont(FUENTE_TITULO, 10)
    c.drawString(margen_x, y, f"Total: $ {format_decimal(datos.total_cancelado)}")
    c.setFont(FUENTE_NORMAL, 8)
    c.drawString(margen_x + 45 * mm, y, f"Son Pesos: {monto_en_palabras(datos.total_cancelado)}")

    y = y_pt(112)
    c.drawString(margen_x, y, f"Motivo de la Nota de Crédito Interna: {datos.motivo_texto}")
    y -= 9
    lineas_nota = [datos.nota[i : i + 90] for i in range(0, len(datos.nota), 90)] or [""]
    for indice, linea in enumerate(lineas_nota):
        prefijo = "Nota: " if indice == 0 else "      "
        c.drawString(margen_x, y, f"{prefijo}{linea}")
        y -= 9

    y_firma = y_pt(120)
    c.drawString(ancho - margen_x - 45 * mm, y_firma, "___________________")
    c.drawCentredString(ancho - margen_x - 22 * mm, y_pt(125), "Confeccionó")

    c.setFont(FUENTE_TITULO, 8)
    c.drawRightString(ancho - margen_x, y_pt(132), etiqueta_copia)


# ---------------------------------------------------------------------------
# Listados varios (migración de Listados.frm) — un único generador
# genérico de tabla paginada, reusado por los ~13 reportes en vez de
# repetir el dibujo manual con `canvas` de arriba: acá la mayoría de las
# columnas del legacy no importan (fuentes/anchos del RTF de VSPrinter),
# lo único que hay que replicar fielmente son los datos/filtros.
#
# **Paginado manual** (pedido del usuario, 2026-08-20 — título+fecha en
# cada hoja, pie con nro. de hoja, "Viene/Transporte" entre hojas,
# totales en recuadro bajo su columna, cantidad de renglones impresos):
# antes se armaba una única `Table` gigante y se dejaba que Platypus
# paginara solo (`repeatRows=1`), lo que alcanzaba para repetir el
# encabezado de columnas pero no para dibujar título/fecha/pie en cada
# hoja (Platypus sólo dibuja un flowable una vez) ni para insertar
# "Viene"/"Transporte" exactamente en el corte de cada hoja (no hay forma
# de saber de antemano dónde va a cortar Platypus). Acá se arma una
# `Table` por cada bloque de filas que entra en una página (separadas por
# `PageBreak()`), y el encabezado/pie de página los dibuja `_NumberedCanvas`
# directo con `canvas`, en cada hoja.
# ---------------------------------------------------------------------------


def generar_pdf_listado(
    *,
    titulo: str,
    subtitulo: str,
    columnas: list[str],
    filas: list[list[str]],
    columnas_derecha: tuple[int, ...] = (),
    pie: Optional[list[tuple[str, str, int]]] = None,
    columna_transporte: Optional[int] = None,
    valores_transporte: Optional[list[Decimal]] = None,
    apaisado: bool = True,
    nombre_archivo: str = "listado.pdf",
    directorio_salida: Optional[Path] = None,
) -> Path:
    """Genera un PDF tabular genérico — reusado por todos los listados de
    `migration.services.ListadosService`. En cada hoja: "ALESTEL SRL" +
    fecha (`formatear_fecha_corta`) arriba, título+subtítulo centrados,
    "Hoja N de M" centrado abajo de todo.

    `columnas_derecha`: índices de columna a alinear a la derecha
    (importes/cantidades), mismo criterio que `TablaBusqueda` en la UI.

    `pie`: lista de `(etiqueta, valor, columna)` para los totales finales
    — `valor` ya formateado tal cual se quiere mostrar (con "$" si es un
    importe, igual que antes; no todos los totales son un importe puro,
    ej. "ND Rechazadas" de Comisiones combina cantidad + importe en un
    solo texto). Cada uno se dibuja en un recuadro bajo SU columna real,
    al pie de la tabla de la ÚLTIMA hoja (antes era una tablita aparte,
    sin relación con la columna de importe).

    `columna_transporte`/`valores_transporte` (opcionales — "cuando
    corresponda": reportes sin una columna de importe corrida con
    sentido de negocio, ej. listado de Clientes, no los pasan):
    `valores_transporte` es paralelo a `filas` (un `Decimal` por fila).
    Si se pasan ambos, cada hoja (salvo la primera) arranca con
    "Viene: $X" (acumulado de hojas previas) y cada hoja (salvo la
    última) termina con "Transporte: $Y" (acumulado hasta ahí) — el
    total de la última hoja coincide con el total general.

    Al final de la última hoja se agrega también, sola, la cantidad de
    renglones impresos.
    """
    directorio = Path(directorio_salida) if directorio_salida is not None else DIR_SALIDA_DEFAULT
    directorio.mkdir(parents=True, exist_ok=True)
    ruta = directorio / nombre_archivo

    tamano_pagina = landscape(A4) if apaisado else A4
    ancho_pagina, alto_pagina = tamano_pagina
    margen = 12 * mm

    # Alto reservado arriba/abajo para el encabezado/pie que dibuja
    # `_NumberedCanvas` en cada hoja — pedido del usuario (2026-08-20,
    # "agregar más renglones... que empiece 2mm después de los títulos y
    # hasta 2mm antes del nro de hoja"): antes había un colchón mucho más
    # generoso ahí (60pt/24pt) sin pedido explícito, que le sacaba
    # renglones de más a cada hoja — ahora es sólo lo que ocupa el texto
    # (empresa/fecha + título + subtítulo arriba; "Hoja N de M" abajo)
    # más el aire pedido antes de que arranque/termine la tabla de datos.
    # El de abajo se ensanchó a 5mm el 2026-08-21 (pedido explícito: "que
    # quede a 5mm de 'Hoja...'" — 2mm quedaba demasiado pegado); el de
    # arriba sigue en 2mm, no se tocó.
    GAP_ENCABEZADO = 2 * mm
    GAP_PIE = 5 * mm
    alto_texto_encabezado = 40  # empresa/fecha + título + subtítulo (ver `_dibujar_marco`)
    alto_texto_pie = 10  # una línea de 8pt ("Hoja N de M")
    alto_encabezado = alto_texto_encabezado + GAP_ENCABEZADO
    alto_pie = alto_texto_pie + GAP_PIE
    top_margin = margen + alto_encabezado
    bottom_margin = margen + alto_pie

    num_columnas = len(columnas)
    ancho_util = ancho_pagina - 2 * margen

    def _estilo_base() -> list[tuple]:
        # Ahorro de tonner (pedido del usuario, 2026-08-20): sin relleno
        # de color en los títulos (antes una barra sólida oscura en toda
        # la fila de encabezado) — en cambio, títulos en negrita más
        # grande que el cuerpo, centrados en cada columna, con una sola
        # línea debajo para separarlos visualmente. Padding de celda más
        # angosto (2pt en vez del default de Platypus, 3pt) para que
        # entren más renglones por hoja ("lo máximo posible", ídem).
        estilo = [
            ("FONTSIZE", (0, 1), (-1, -1), 7.5),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#bbbbbb")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
        for col in columnas_derecha:
            estilo.append(("ALIGN", (col, 0), (col, -1), "RIGHT"))
        # Encabezado: va DESPUÉS del `ALIGN` de `columnas_derecha` de
        # arriba a propósito, para pisar el "RIGHT" de esa fila con
        # "CENTER" (los comandos de `TableStyle` se aplican en orden, el
        # último gana sobre el mismo rango).
        estilo.append(("FONTNAME", (0, 0), (-1, 0), FUENTE_TITULO))
        estilo.append(("FONTSIZE", (0, 0), (-1, 0), 9))
        estilo.append(("ALIGN", (0, 0), (-1, 0), "CENTER"))
        estilo.append(("LINEBELOW", (0, 0), (-1, 0), 1, colors.black))
        return estilo

    # --- Cuántas filas de datos entran por página -----------------------
    # Medido de verdad con `Table.wrap()` (no un alto de fila "a ojo") —
    # bug real encontrado probando con 79 filas (2026-08-20): un valor
    # aproximado dejaba cada bloque un pelo más grande de lo que entraba
    # de verdad, y Platypus terminaba partiendo CADA bloque en 2 páginas
    # reales (una con casi todo el bloque, otra con 1 sola fila + la fila
    # de Transporte) — funcionaba (no se perdía nada) pero desperdiciaba
    # una hoja extra en cada corte. Se arman 2 tablas de prueba con el
    # mismo estilo real (sólo encabezado, y encabezado + 1 fila) y se
    # resta para saber el alto real de una fila de datos.
    _tabla_solo_header = Table([columnas], repeatRows=1)
    _tabla_solo_header.setStyle(TableStyle(_estilo_base()))
    _, _alto_header = _tabla_solo_header.wrap(ancho_util, alto_pagina)

    _tabla_una_fila = Table([columnas, [""] * num_columnas], repeatRows=1)
    _tabla_una_fila.setStyle(TableStyle(_estilo_base()))
    _, _alto_header_mas_fila = _tabla_una_fila.wrap(ancho_util, alto_pagina)

    alto_fila = max(_alto_header_mas_fila - _alto_header, 1)
    alto_util_datos = alto_pagina - top_margin - bottom_margin
    # Colchón para "Viene"/`pie`/el renglón de conteo, calculado — no un
    # número fijo "por las dudas" (pedido del usuario, 2026-08-21: "queda
    # mucho espacio desaprovechado al final de la hoja"; un colchón fijo
    # de 6 renglones reservaba de más en la enorme mayoría de los
    # listados reales, que no usan `columna_transporte` — "Viene" NUNCA
    # aparece ahí, y muchos ni siquiera tienen `pie`). "Viene" y
    # "Transporte" nunca están juntos en la misma hoja (uno es de
    # arranque, el otro de cierre), así que contarlos como si pudieran
    # coincidir era además doble margen de más. Si el cálculo se queda
    # corto en algún caso límite, Platypus parte la tabla puntual sola
    # (red de seguridad ya existente, no se pierde nada).
    FILAS_RESERVADAS = (1 if columna_transporte is not None else 0) + len(pie or []) + 1
    filas_por_pagina = max(1, int((alto_util_datos - _alto_header) / alto_fila) - FILAS_RESERVADAS)

    def _fila_total(etiqueta: str, valor: str, columna: int) -> list[str]:
        # `valor` viaja ya formateado tal cual se quiere mostrar (con "$"
        # si es un importe, como ya hacían los call sites existentes) —
        # no se le antepone nada acá: algunos totales no son un importe
        # puro (ej. "ND Rechazadas" de Comisiones combina cantidad +
        # importe en un solo texto).
        fila = [""] * num_columnas
        if columna <= 0:
            fila[0] = f"{etiqueta} {valor}"
        else:
            fila[0] = etiqueta
            fila[columna] = valor
        return fila

    def _estilizar_fila_total(estilo: list[tuple], fila_idx: int, columna: int) -> None:
        estilo.append(("FONTNAME", (0, fila_idx), (-1, fila_idx), FUENTE_TITULO))
        if columna > 0:
            estilo.append(("SPAN", (0, fila_idx), (columna - 1, fila_idx)))
            estilo.append(("BOX", (columna, fila_idx), (columna, fila_idx), 0.8, colors.black))
            estilo.append(("ALIGN", (columna, fila_idx), (columna, fila_idx), "RIGHT"))
        else:
            estilo.append(("SPAN", (0, fila_idx), (-1, fila_idx)))
            estilo.append(("BOX", (0, fila_idx), (0, fila_idx), 0.8, colors.black))

    # --- Armado de las páginas ------------------------------------------
    total_filas = len(filas)
    num_paginas = -(-total_filas // filas_por_pagina) if total_filas else 1
    elementos: list = []
    acumulado_transporte = Decimal("0")

    for num_pagina in range(num_paginas):
        es_primera = num_pagina == 0
        es_ultima = num_pagina == num_paginas - 1
        desde = num_pagina * filas_por_pagina
        hasta = min(desde + filas_por_pagina, total_filas)
        bloque = filas[desde:hasta]

        datos_tabla = [columnas]
        estilo = _estilo_base()
        filas_zebra: list[int] = []  # índices dentro de ESTA tabla, sólo filas de datos reales

        if not es_primera and columna_transporte is not None:
            fila_idx = len(datos_tabla)
            datos_tabla.append(_fila_total("Viene:", f"$ {format_decimal(acumulado_transporte)}", columna_transporte))
            _estilizar_fila_total(estilo, fila_idx, columna_transporte)

        for fila in bloque:
            filas_zebra.append(len(datos_tabla))
            datos_tabla.append(fila)

        if valores_transporte is not None and columna_transporte is not None:
            acumulado_transporte += sum(valores_transporte[desde:hasta], Decimal("0"))

        if not es_ultima and columna_transporte is not None:
            fila_idx = len(datos_tabla)
            datos_tabla.append(_fila_total("Transporte:", f"$ {format_decimal(acumulado_transporte)}", columna_transporte))
            _estilizar_fila_total(estilo, fila_idx, columna_transporte)

        if es_ultima:
            for etiqueta, valor, columna in pie or []:
                fila_idx = len(datos_tabla)
                datos_tabla.append(_fila_total(f"{etiqueta}:", valor, columna))
                _estilizar_fila_total(estilo, fila_idx, columna)

            fila_idx = len(datos_tabla)
            # Bug real encontrado probando (2026-08-20): "renglón"/
            # "impreso" no pluralizan con el mismo sufijo ("renglones",
            # no "renglónes"; "impresos", no "impresoes") — un solo
            # sufijo compartido quedaba mal en los dos.
            texto_conteo = "1 renglón impreso" if total_filas == 1 else f"{total_filas} renglones impresos"
            fila_conteo = [""] * num_columnas
            fila_conteo[0] = texto_conteo
            datos_tabla.append(fila_conteo)
            estilo.append(("FONTNAME", (0, fila_idx), (-1, fila_idx), "Helvetica-Oblique"))
            estilo.append(("SPAN", (0, fila_idx), (-1, fila_idx)))

        # Zebra striping sólo sobre filas de datos reales — no sobre el
        # encabezado ni las de total (que ya tienen su propio estilo).
        for i, idx in enumerate(filas_zebra):
            if i % 2 == 1:
                estilo.append(("BACKGROUND", (0, idx), (-1, idx), colors.HexColor("#f2f2f2")))

        tabla = Table(datos_tabla, repeatRows=1)
        tabla.setStyle(TableStyle(estilo))
        elementos.append(tabla)
        if not es_ultima:
            elementos.append(PageBreak())

    class _NumberedCanvas(canvas.Canvas):
        """Difiere el dibujo del pie de página (necesita el total de
        páginas, que sólo se sabe al terminar) hasta `save()` — patrón
        estándar de ReportLab para numerar "Hoja N de M"."""

        def __init__(self, *args, **kwargs):
            canvas.Canvas.__init__(self, *args, **kwargs)
            self._paginas_guardadas: list[dict] = []

        def showPage(self) -> None:  # noqa: N802 (override de reportlab)
            self._paginas_guardadas.append(dict(self.__dict__))
            self._startPage()

        def save(self) -> None:  # noqa: N802 (override de reportlab)
            total_paginas = len(self._paginas_guardadas)
            for estado in self._paginas_guardadas:
                self.__dict__.update(estado)
                self._dibujar_marco(total_paginas)
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)

        def _dibujar_marco(self, total_paginas: int) -> None:
            self.saveState()
            y_tope = alto_pagina - margen
            self.setFont(FUENTE_TITULO, 11)
            self.drawString(margen, y_tope - 8, "ALESTEL SRL")
            self.setFont("Helvetica", 9)
            self.drawRightString(ancho_pagina - margen, y_tope - 8, formatear_fecha_corta(date.today()))
            self.setFont(FUENTE_TITULO, 13)
            self.drawCentredString(ancho_pagina / 2, y_tope - 26, titulo)
            self.setFont("Helvetica", 8.5)
            self.drawCentredString(ancho_pagina / 2, y_tope - 38, subtitulo)
            self.setFont("Helvetica", 8)
            # Alineado con `alto_texto_pie` de más arriba — la tabla
            # termina 2mm por encima de este texto (ver `alto_pie`).
            self.drawCentredString(ancho_pagina / 2, margen + 4, f"Hoja {self.getPageNumber()} de {total_paginas}")
            self.restoreState()

    doc = SimpleDocTemplate(
        str(ruta), pagesize=tamano_pagina,
        leftMargin=margen, rightMargin=margen, topMargin=top_margin, bottomMargin=bottom_margin,
        title=titulo,
    )
    doc.build(elementos, canvasmaker=_NumberedCanvas)
    return ruta
