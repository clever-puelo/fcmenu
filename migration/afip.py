"""Integración AFIP (WSAA + WSFEv1) para Factura Electrónica.

Migración de `EmiFact.frm Sub ConectaAFIP()` (líneas 2469-2708) y
`Sub Obtener_QR()` (líneas 2774-2851). Alcance acotado a lo ya confirmado
con el usuario: sólo Factura, Letra A/B (`CODIGOS_AFIP`), y **sin
credenciales de Homologación todavía** — el usuario decidió avanzar con
una interfaz clara (`NumeracionYCAEProvider`) mockeable/stub por ahora,
conectando las credenciales reales más adelante.

Reemplaza el patrón legacy `CreateObject("WSAA")`/`CreateObject("WSFE")`
(interfaz COM de PyAfipWs) por un cliente SOAP nativo (`zeep` +
`cryptography` para firmar el TRA en CMS/PKCS7) — decisión del usuario
(2026-08-16) de no depender de `pyafipws`: WSAA/WSFEv1 son sólo dos
servicios SOAP estándar, no hace falta el wrapper, y el paquete de
PyPI de `pyafipws` falla al instalar en este entorno (ver docstring de
`AfipWSFEv1Cliente`). Sin ActiveX ni Shell externo — ver CLAUDE.md
sección "Facturación Electrónica" (actualizada para reflejar esta
decisión).

**Hallazgo que corrige una nota previa de memoria**: `Punto_Vta` en el
legacy NO se hardcodea por Letra (A/B) como se había anotado antes —
línea 2617: `If CodFactAFIP = 1 Or CodFactAFIP = 6 Then Punto_Vta = 3
Else Punto_Vta = 4`. `CodFactAFIP` 1 y 6 son Factura A y Factura B
respectivamente (códigos WSFE reales, ver `CODIGOS_AFIP`) — así que la
regla real es **Punto de Venta 3 para Factura (A o B), Punto de Venta 4
para NC/ND (A o B)**, no "3 para A, 4 para B". Como el alcance actual es
sólo Factura, en la práctica el Punto de Venta a usar hoy es siempre 3 —
pero ese valor concreto (cuál Punto de Venta está habilitado en AFIP
para Factura Electrónica de esta empresa) hay que confirmarlo con
credenciales reales antes de ir a producción, no asumirlo de nuevo acá.
"""

from __future__ import annotations

import base64
from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal
from typing import Optional

# Códigos de comprobante WSFEv1 reales (confirmados contra EmiFact.frm:721-733,
# que ya los arma correctamente a partir de Letra+TipoFac). Sólo A/B de
# Factura están en alcance hoy; NC/ND quedan documentados para cuando se
# migre esa fase.
CODIGOS_AFIP = {
    ("A", 1): 1,  # Factura A
    ("A", 2): 3,  # Nota de Crédito A
    ("A", 3): 2,  # Nota de Débito A
    ("B", 1): 6,  # Factura B
    ("B", 2): 8,  # Nota de Crédito B
    ("B", 3): 7,  # Nota de Débito B
}

# Punto de Venta AFIP por tipo de comprobante — ver hallazgo en el
# docstring del módulo. Sólo FACTURA está en alcance; NC_ND queda
# documentado para cuando se migre esa fase.
PUNTO_VENTA_FACTURA = 3
PUNTO_VENTA_NC_ND = 4

TIPO_DOC_CUIT = 80  # WSFEv1 Tipo_Doc=80 (legacy siempre usa este, ver docstring de solicitar_cae)


@dataclass
class ResultadoCAE:
    """Respuesta de `WSFE.Aut()` ya normalizada (`CAE`/`Vencimiento`/
    `Resultado`/`Motivo` en el legacy)."""

    cae: str
    vencimiento: date
    resultado: str  # "A" (aprobado) / "R" (rechazado)
    motivo: str = ""

    @property
    def aprobado(self) -> bool:
        # EmiFact.frm:2668: `CAE = "NULL" Or WSFE.Resultado <> "A"` -> rechazo.
        return bool(self.cae) and self.cae != "NULL" and self.resultado == "A"


class NumeracionYCAEProvider(ABC):
    """Interfaz que separa la resolución de numeración/CAE (AFIP) de la
    persistencia del comprobante (`EmisionFacturaService`, paso 3) — así
    `EmisionFacturaService.emitir_factura()` sigue recibiendo el número y
    la letra ya resueltos, sin depender de si el proveedor es el stub de
    desarrollo o el cliente real de AFIP."""

    @abstractmethod
    def ultimo_comprobante(self, punto_venta: int, tipo_cbte: int) -> int:
        """Réplica de `WSFE.RecuperaLastCMP(PtoVta, TipoCbte)` — el
        próximo número a emitir es este valor + 1 (`EmiFact.frm:2623,
        2634`)."""

    @abstractmethod
    def solicitar_cae(
        self,
        tipo_cbte: int,
        punto_venta: int,
        cbte_nro: int,
        cuit_receptor: str,
        importe_neto: Decimal,
        importe_iva: Decimal,
        importe_iibb: Decimal,
        importe_total: Decimal,
        fecha_cbte: date,
    ) -> ResultadoCAE:
        """Réplica de `WSFE.Aut(...)` (`EmiFact.frm:2650-2654`). El
        `Nro_Doc` del receptor siempre viaja como CUIT (`Tipo_Doc=80`,
        `TIPO_DOC_CUIT`) — igual que el legacy, que nunca contempla
        Consumidor Final sin CUIT (código 99) ni DNI (código 96); si se
        necesita eso, es una ampliación de alcance a confirmar, no algo
        que este método deba inventar."""


class AfipWSFEv1Stub(NumeracionYCAEProvider):
    """Implementación FALSA para desarrollo — el usuario confirmó avanzar
    sin credenciales de Homologación por ahora. Genera un CAE determinístico
    (no es un CAE real, nunca lo va a aceptar AFIP) para poder construir y
    probar `FacturadorWindow`/`EmisionFacturaService` de punta a punta sin
    depender de la conexión real. **No usar en producción** — reemplazar
    por `AfipWSFEv1Cliente` cuando haya CUIT/certificado/clave reales."""

    def __init__(self):
        self._contadores: dict[tuple[int, int], int] = {}

    def ultimo_comprobante(self, punto_venta: int, tipo_cbte: int) -> int:
        return self._contadores.get((punto_venta, tipo_cbte), 0)

    def solicitar_cae(
        self,
        tipo_cbte: int,
        punto_venta: int,
        cbte_nro: int,
        cuit_receptor: str,
        importe_neto: Decimal,
        importe_iva: Decimal,
        importe_iibb: Decimal,
        importe_total: Decimal,
        fecha_cbte: date,
    ) -> ResultadoCAE:
        if not cuit_receptor or not cuit_receptor.strip():
            raise ValueError("No se puede solicitar CAE sin CUIT del receptor.")

        clave = (punto_venta, tipo_cbte)
        self._contadores[clave] = max(self._contadores.get(clave, 0), cbte_nro)

        # El CAE real de AFIP es SIEMPRE numérico (14 dígitos) — acá se
        # arma uno igual de numérico (no un prefijo "STUB" con letras)
        # para que sea intercambiable con `generar_qr_afip()` y con
        # cualquier otro código que asuma el formato real, tal como se
        # usaría un CAE real. Sigue siendo obviamente falso (no lo va a
        # reconocer AFIP), sólo compatible en forma.
        cae_falso = f"{fecha_cbte:%Y%m%d}{punto_venta:04d}{cbte_nro % 100:02d}"
        # AFIP: el CAE vence a los 10 días corridos desde la emisión.
        return ResultadoCAE(cae=cae_falso, vencimiento=fecha_cbte + timedelta(days=10), resultado="A", motivo="")


WSAA_WSDL = {
    True: "https://wsaahomo.afip.gov.ar/ws/services/LoginCms?WSDL",
    False: "https://wsaa.afip.gov.ar/ws/services/LoginCms?WSDL",
}
WSFEV1_WSDL = {
    True: "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL",
    False: "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL",
}


# ---------------------------------------------------------------------------
# Entorno de facturación electrónica — parámetro EXTERNO a la app, mismo
# criterio que el legacy (que lo tomaba de un argumento de línea de comandos
# `/TF:`, ver `FCMENU.bas` líneas 301-304 y 470-486, y lo mostraba en el
# cartel `Label6` de `FCMENU.frm`/`Inicio.frm`). Acá NO se usa un argumento de
# arranque (decisión del usuario, 2026-08-16) — se lee de la variable de
# entorno `FCMENU_AFIP_ENTORNO` (ver `.env.example`), mismo mecanismo que ya
# se usa para `DATABASE_URL` (`migration/db.py`): nada que editar en el .py
# para pasar de prueba a producción.
#
# A diferencia del legacy (que sólo conocía Homologación/Producción, porque
# siempre hablaba con AFIP de verdad), acá hay un tercer estado real: STUB
# (`AfipWSFEv1Stub` arriba, sin conexión ninguna a AFIP, CAE falso) — es el
# default mientras no haya credenciales de Homologación reales cargadas. De
# cara al operador (cartel de `MainMenuWindow`) STUB y HOMOLOGACION se
# muestran igual ("Homologación / Prueba"): la diferencia entre "sin
# conexión" y "conectado a los servidores de prueba de AFIP" es un detalle
# de implementación, no algo que el operador necesite distinguir.
ENTORNOS_AFIP_VALIDOS = ("STUB", "HOMOLOGACION", "PRODUCCION")


def entorno_afip() -> str:
    """Lee `FCMENU_AFIP_ENTORNO` del entorno del proceso. Default `STUB`
    (ver arriba). Cualquier valor no reconocido cae también a `STUB` —
    antes preferir "sin conexión" por error que arriesgar apuntar a
    Producción por una variable mal tipeada."""
    import os

    valor = os.environ.get("FCMENU_AFIP_ENTORNO", "STUB").strip().upper()
    return valor if valor in ENTORNOS_AFIP_VALIDOS else "STUB"


def etiqueta_entorno_afip(entorno: str | None = None) -> str:
    """Texto del cartel de estado de `MainMenuWindow` — mismos textos
    literales que usaba `Label6` en el legacy."""
    entorno = entorno if entorno is not None else entorno_afip()
    return "PRODUCCIÓN" if entorno == "PRODUCCION" else "Homologación / Prueba"

# AlicIva.Id — códigos de alícuota de IVA reales de WSFEv1. Sólo 21%
# (Id=5) está en alcance hoy (única alícuota que usa el sistema, ver
# FacturaService.alicuota_iva_inscripto — Parametro.IVAINS real es 21).
ALICUOTA_IVA_21_ID = 5

# Margen de seguridad para renovar el token antes de que venza (AFIP lo
# emite válido por 12hs — WSAA.CreateTRA original del legacy, comentario
# "se puede usar el mismo token y sign por 6 horas").
_MARGEN_RENOVACION_TOKEN = timedelta(minutes=10)
_TZ_ARGENTINA = timezone(timedelta(hours=-3))


class AfipWSFEv1Cliente(NumeracionYCAEProvider):
    """Adaptador real sobre WSAA/WSFEv1 vía SOAP nativo (`zeep` +
    `cryptography` para firmar el TRA en CMS/PKCS7) — reemplaza tanto el
    `CreateObject("WSAA")`/`CreateObject("WSFE")` COM del legacy como la
    dependencia de `pyafipws` (decisión del usuario, 2026-08-16: la
    librería no es necesaria — WSAA/WSFEv1 son sólo dos servicios SOAP —
    y su paquete de PyPI falla al instalar en este entorno, ver más
    abajo). `cryptography`/`zeep` están en `requirements.txt`, instalan
    limpio.

    **Verificado contra el WSDL y el servicio REAL de Homologación de
    AFIP** (sin credenciales todavía, pero la conectividad SOAP en sí
    sí se probó): `FEDummy()` contra
    `https://wswhomo.afip.gov.ar/wsfev1/service.asmx?WSDL` respondió
    `AppServer/DbServer/AuthServer: OK` en vivo, y una llamada de prueba
    a `FECompUltimoAutorizado` con credenciales inválidas devolvió el
    error real de AFIP (`Errors.Err: [{Code: 600, Msg: "ValidacionDeToken:
    No valido token..."}]`) — confirma que los nombres de operación y la
    forma de los tipos (`FEAuthRequest`, `FECAERequest`, etc., ver
    `_solicitud_cae()`) son correctos. **Lo único que falta para
    funcionar de punta a punta es un certificado/clave reales dados de
    alta en AFIP** (`WSAA.CallWSAA` rechaza cualquier CMS firmado con un
    certificado no reconocido) — no se puede probar más sin eso.

    **`CondicionIVAReceptorId`** (`FECAEDetRequest`, campo real del WSDL
    vigente): quedó fuera del legacy porque es un requisito de AFIP
    posterior (RG 5616/2024) a la última vez que se tocó `EmiFact.frm`.
    Es opcional a nivel XSD pero probablemente exigido por AFIP en la
    práctica para ciertos casos — se expone como parámetro opcional acá
    (`condicion_iva_receptor_id`) en vez de inventar un valor; falta
    confirmar el mapeo real Cliente.CIVA -> este código cuando haya
    credenciales para probar contra AFIP real.
    """

    def __init__(self, cuit_emisor: str, certificado_path: str, clave_privada_path: str, homologacion: bool = True):
        self.cuit_emisor = cuit_emisor
        self.certificado_path = certificado_path
        self.clave_privada_path = clave_privada_path
        self.homologacion = homologacion
        self._token: Optional[str] = None
        self._sign: Optional[str] = None
        self._token_vence: Optional[datetime] = None
        self._wsaa_client = None
        self._wsfe_client = None

    # ------------------------------------------------------------------
    # WSAA — autenticación (TRA firmado en CMS, ver docstring de la clase)
    # ------------------------------------------------------------------
    def _token_vigente(self) -> bool:
        if self._token is None or self._sign is None or self._token_vence is None:
            return False
        return datetime.now(_TZ_ARGENTINA) < (self._token_vence - _MARGEN_RENOVACION_TOKEN)

    def _generar_tra(self) -> bytes:
        ahora = datetime.now(_TZ_ARGENTINA)
        unique_id = int(ahora.timestamp())
        generation_time = (ahora - timedelta(minutes=5)).isoformat(timespec="seconds")
        expiration_time = (ahora + timedelta(minutes=10)).isoformat(timespec="seconds")
        tra = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<loginTicketRequest version="1.0">'
            "<header>"
            f"<uniqueId>{unique_id}</uniqueId>"
            f"<generationTime>{generation_time}</generationTime>"
            f"<expirationTime>{expiration_time}</expirationTime>"
            "</header>"
            "<service>wsfe</service>"
            "</loginTicketRequest>"
        )
        return tra.encode("utf-8")

    def _firmar_tra(self, tra: bytes) -> str:
        """Firma el TRA como CMS/PKCS7 SignedData (adjunto, con el
        certificado incluido — lo que AFIP espera en `loginCms`),
        codificado en base64. Réplica de `WSAA.SignTRA()` del legacy,
        pero con `cryptography` en vez de PyAfipWs/OpenSSL externo."""
        from cryptography.hazmat.primitives import hashes, serialization
        from cryptography.hazmat.primitives.serialization import pkcs7
        from cryptography import x509

        with open(self.certificado_path, "rb") as f:
            certificado = x509.load_pem_x509_certificate(f.read())
        with open(self.clave_privada_path, "rb") as f:
            clave_privada = serialization.load_pem_private_key(f.read(), password=None)

        cms = (
            pkcs7.PKCS7SignatureBuilder()
            .set_data(tra)
            .add_signer(certificado, clave_privada, hashes.SHA256())
            .sign(serialization.Encoding.DER, [pkcs7.PKCS7Options.Binary])
        )
        return base64.b64encode(cms).decode("ascii")

    def _autenticar(self) -> None:
        from zeep import Client
        from zeep.exceptions import Fault

        if self._wsaa_client is None:
            self._wsaa_client = Client(WSAA_WSDL[self.homologacion])

        tra = self._generar_tra()
        cms_b64 = self._firmar_tra(tra)

        try:
            respuesta_xml = self._wsaa_client.service.loginCms(in0=cms_b64)
        except Fault as exc:
            # Ej. real verificado contra Homologación: "Certificado no
            # emitido por AC de confianza" (certificado no reconocido por
            # AFIP) — se relanza con mensaje claro en vez de dejar pasar
            # la excepción cruda de zeep.
            raise RuntimeError(f"WSAA rechazó la autenticación: {exc}") from exc

        import xml.etree.ElementTree as ET

        raiz = ET.fromstring(respuesta_xml)
        token = raiz.findtext("./credentials/token")
        sign = raiz.findtext("./credentials/sign")
        expiration_time = raiz.findtext("./header/expirationTime")
        if not token or not sign:
            raise RuntimeError(f"WSAA no devolvió token/sign válidos. Respuesta: {respuesta_xml}")

        self._token, self._sign = token, sign
        self._token_vence = datetime.fromisoformat(expiration_time) if expiration_time else None

    # ------------------------------------------------------------------
    # WSFEv1 — numeración y CAE
    # ------------------------------------------------------------------
    def _cliente_wsfe(self):
        from zeep import Client

        if not self._token_vigente():
            self._autenticar()
        if self._wsfe_client is None:
            self._wsfe_client = Client(WSFEV1_WSDL[self.homologacion])
        return self._wsfe_client

    def _auth_request(self) -> dict:
        return {"Token": self._token, "Sign": self._sign, "Cuit": int(self.cuit_emisor.replace("-", ""))}

    @staticmethod
    def _errores_de(respuesta) -> str:
        errores = getattr(respuesta, "Errors", None)
        if not errores or not getattr(errores, "Err", None):
            return ""
        return "; ".join(f"[{e.Code}] {e.Msg}" for e in errores.Err)

    @staticmethod
    def _llamar(nombre: str, operacion, **kwargs):
        """Invoca una operación SOAP de WSFEv1, traduciendo un `Fault` de
        `zeep` (rechazo a nivel SOAP, ej. token vencido) a un error claro
        — los rechazos de negocio (CAE denegado, etc.) vienen aparte,
        dentro de la respuesta normal (`Errors`, ver `_errores_de`)."""
        from zeep.exceptions import Fault

        try:
            return operacion(**kwargs)
        except Fault as exc:
            raise RuntimeError(f"{nombre} falló a nivel SOAP: {exc}") from exc

    def ultimo_comprobante(self, punto_venta: int, tipo_cbte: int) -> int:
        wsfe = self._cliente_wsfe()
        respuesta = self._llamar(
            "FECompUltimoAutorizado",
            wsfe.service.FECompUltimoAutorizado,
            Auth=self._auth_request(),
            PtoVta=punto_venta,
            CbteTipo=tipo_cbte,
        )
        errores = self._errores_de(respuesta)
        if errores:
            raise RuntimeError(f"FECompUltimoAutorizado falló: {errores}")
        return int(respuesta.CbteNro or 0)

    def solicitar_cae(
        self,
        tipo_cbte: int,
        punto_venta: int,
        cbte_nro: int,
        cuit_receptor: str,
        importe_neto: Decimal,
        importe_iva: Decimal,
        importe_iibb: Decimal,
        importe_total: Decimal,
        fecha_cbte: date,
        condicion_iva_receptor_id: Optional[int] = None,
    ) -> ResultadoCAE:
        if not cuit_receptor or not cuit_receptor.strip():
            raise ValueError("No se puede solicitar CAE sin CUIT del receptor.")

        wsfe = self._cliente_wsfe()
        fecha_txt = fecha_cbte.strftime("%Y%m%d")

        detalle = {
            "Concepto": 1,  # Productos
            "DocTipo": TIPO_DOC_CUIT,
            "DocNro": int(cuit_receptor.replace("-", "")),
            "CbteDesde": cbte_nro,
            "CbteHasta": cbte_nro,
            "CbteFch": fecha_txt,
            "ImpTotal": float(importe_total),
            "ImpTotConc": 0.00,
            "ImpNeto": float(importe_neto),
            "ImpOpEx": 0.00,
            "ImpTrib": float(importe_iibb),
            "ImpIVA": float(importe_iva),
            "FchVtoPago": fecha_txt,
            "MonId": "PES",
            "MonCotiz": 1,
        }
        if importe_iva > 0:
            detalle["Iva"] = {
                "AlicIva": [
                    {"Id": ALICUOTA_IVA_21_ID, "BaseImp": float(importe_neto), "Importe": float(importe_iva)}
                ]
            }
        if condicion_iva_receptor_id is not None:
            detalle["CondicionIVAReceptorId"] = condicion_iva_receptor_id

        solicitud = {
            "FeCabReq": {"CantReg": 1, "PtoVta": punto_venta, "CbteTipo": tipo_cbte},
            "FeDetReq": {"FECAEDetRequest": [detalle]},
        }

        respuesta = self._llamar(
            "FECAESolicitar", wsfe.service.FECAESolicitar, Auth=self._auth_request(), FeCAEReq=solicitud
        )
        errores = self._errores_de(respuesta)
        if errores:
            raise RuntimeError(f"FECAESolicitar falló: {errores}")

        detalle_resp = respuesta.FeDetResp.FECAEDetResponse[0]
        observaciones = getattr(detalle_resp, "Observaciones", None)
        motivo = ""
        if observaciones and getattr(observaciones, "Obs", None):
            motivo = "; ".join(f"[{o.Code}] {o.Msg}" for o in observaciones.Obs)

        return ResultadoCAE(
            cae=detalle_resp.CAE or "",
            vencimiento=_parsear_fecha_afip(detalle_resp.CAEFchVto) or fecha_cbte,
            resultado=respuesta.FeCabResp.Resultado or "",
            motivo=motivo,
        )


def _parsear_fecha_afip(valor: Optional[str]) -> Optional[date]:
    """WSFEv1 devuelve fechas como texto 'yyyymmdd'."""
    if not valor:
        return None
    return date(int(valor[0:4]), int(valor[4:6]), int(valor[6:8]))


def codigo_afip(letra: str, tipo_fac: int) -> int:
    """Código de comprobante WSFEv1 real por Letra+TipoFac (réplica de
    `EmiFact.frm:721-733`). `tipo_fac`: 1=Factura, 2=NC, 3=ND (mismos
    índices que `CabFact.frm Option1`)."""
    clave = (letra, tipo_fac)
    if clave not in CODIGOS_AFIP:
        raise ValueError(f"Combinación Letra={letra!r}/TipoFac={tipo_fac!r} no soportada.")
    return CODIGOS_AFIP[clave]


def punto_venta_por_tipo(tipo_fac: int) -> int:
    """Punto de Venta AFIP por tipo de comprobante (ver hallazgo en el
    docstring del módulo) — 3 para Factura, 4 para NC/ND. Sólo Factura
    está en alcance hoy."""
    return PUNTO_VENTA_FACTURA if tipo_fac == 1 else PUNTO_VENTA_NC_ND


def generar_qr_afip(
    cuit_emisor: str,
    punto_venta: int,
    tipo_cbte: int,
    nro_cbte: int,
    importe_total: Decimal,
    tipo_doc_receptor: int,
    nro_doc_receptor: str,
    cae: str,
    fecha_cbte: Optional[date] = None,
    moneda: str = "PES",
    cotizacion: Decimal = Decimal("1"),
) -> str:
    """Genera la URL del QR de AFIP según la especificación real (RG
    4892) — reemplaza `Obtener_QR()` (`EmiFact.frm:2774-2851`), que tenía
    un bug real en producción: un `GoTo SaltaFactura` salteaba el armado
    del JSON real y usaba un string hardcodeado con el QR de OTRA factura
    (otro CUIT, de 2021) para TODAS las facturas — cada factura emitida
    hoy imprime un QR ajeno. Acá se arma el JSON real siempre, en Python,
    sin depender de un .exe externo (`Shell("qrcode.exe ...")` en el
    legacy, eliminado por directiva del CLAUDE.md).

    El formato del JSON (campos y su orden) está confirmado empíricamente
    contra el ejemplo real que el legacy tenía hardcodeado en
    `Texto1` (`EmiFact.frm:2785`, un QR real de AFIP de 2021, CUIT
    33708755309) — decodificado: `{"ver":1,"fecha":"2021-03-17",
    "cuit":33708755309,"ptoVta":5,"tipoCmp":1,"nroCmp":1511,
    "importe":38478.00,"moneda":"PES","ctz":1.000000,"tipoDocRec":80,
    "nroDocRec":305227806,... "tipoCodAut":"E","codAut":...}`. El bloque
    de código real que el `GoTo` salteaba tenía además sus propios
    errores de formato (comillas faltantes en `"moneda"`/`"tipoCodAut"`,
    `ctz` ausente, `importe` multiplicado x100) — no se portó ese bloque,
    se armó de nuevo contra el ejemplo real.
    """
    fecha_cbte = fecha_cbte or date.today()
    cuit_emisor_num = int(cuit_emisor.replace("-", ""))
    nro_doc_num = int(nro_doc_receptor.replace("-", ""))
    cae_num = int(cae)

    payload = (
        "{"
        f'"ver":1,'
        f'"fecha":"{fecha_cbte:%Y-%m-%d}",'
        f'"cuit":{cuit_emisor_num},'
        f'"ptoVta":{punto_venta},'
        f'"tipoCmp":{tipo_cbte},'
        f'"nroCmp":{nro_cbte},'
        f'"importe":{Decimal(importe_total):.2f},'
        f'"moneda":"{moneda}",'
        f'"ctz":{Decimal(cotizacion):.6f},'
        f'"tipoDocRec":{tipo_doc_receptor},'
        f'"nroDocRec":{nro_doc_num},'
        f'"tipoCodAut":"E",'
        f'"codAut":{cae_num}'
        "}"
    )
    b64 = base64.b64encode(payload.encode("utf-8")).decode("ascii")
    return f"https://www.afip.gob.ar/fe/qr/?p={b64}"
