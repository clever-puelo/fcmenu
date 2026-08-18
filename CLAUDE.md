# Contexto del Proyecto: Migración ERP VB6 a Python

## 1. Visión General
Este proyecto consiste en auditar, refactorizar y migrar un ERP legacy de Facturación y Cuentas Corrientes desarrollado en **Visual Basic 6.0 (VB6)** hacia una aplicación moderna en **Python**.

- **Proyecto Principal VB6:** `FCMENU.vbp`
- **Base de Datos Origen:** `Fcmenu.mdb` (MS-Access vía ADODB).
- **Base de Datos Destino:** PostgreSQL (con ORM SQLAlchemy en Python).
- **Entorno del Cliente:** Operadores muy estrictos con la usabilidad, velocidad de carga y atajos de teclado.

---

## 2. Reglas de Oro para Claude Code (CONSULTA Y VALIDACIÓN)

1. **Dudas, Inconsistencias y Código Muerto:**
   - Si al revisar el código de VB6 encuentras parches contradictorios, lógica ambigua, consultas SQL ineficientes o variables no usadas, **DETENTE y pregunta al usuario** antes de asumir una solución.
   - Si detectas diferencias en lógica de redondeos, cálculo de IVA, o saldos de cuentas corrientes, debes reportarlo de inmediato para su validación.

2. **Cero Suposiciones en Reglas de Negocio:**
   - La lógica de imputación de pagos, facturación y cuentas corrientes debe extraerse del VB6 y confirmarse paso a paso.

3. **Estrategia por Fases:**
   - No traduzcas todo el proyecto de golpe. Avanza **formulario por formulario** (`.frm`) y **módulo por módulo** (`.bas`/`.cls`).

---

## 3. Mapeo de Tecnologías (VB6 Legacy -> Python Target)

### Core y Base de Datos
- **Base de Datos:** Migración de `Fcmenu.mdb` (MS-Access) a **PostgreSQL**.
- **Acceso a Datos:** Reemplazar llamadas ADODB directas por **SQLAlchemy** (ORM / Core).
- **Tipos de Datos Monetarios:** En Python se debe usar **SIEMPRE** `decimal.Decimal` para importes, IVA, retenciones y saldos. Prohibido usar `float`.

### Interfaz Gráfica (GUI) y Controles Especiales
- **Framework GUI:** **PyQt6** (o PySide6).
- **ComponentOne FlexGrid / CodeJock / DBI-Tech:** Reemplazar por `QTableView` + `QAbstractTableModel` de PyQt6. Deben permitir edición rápida en celda, navegación fluida con flechas/Enter y coloreado condicional.
- **Navegación Teclado:** El evento `Enter` debe comportarse como `Tab` (avance de campo). Configurar teclas de función (`F1` a `F12`) para búsquedas y acciones rápidas según el comportamiento original de VB6.

### Facturación Electrónica (AFIP / ARCA)
- **Comprobantes:** Facturas A, B (ocasionales) y Remitos. 100% Factura Electrónica (sin controlador fiscal).
- **Servicios AFIP:** `WSAA` (Autenticación) y `WSFEv1` (Facturación).
- **Librería en Python:** cliente SOAP nativo propio (`migration/afip.py`), **sin PyAfipWs** (decisión confirmada 2026-08-16). WSAA/WSFEv1 son sólo dos servicios SOAP estándar — no hace falta el wrapper de PyAfipWs, y su paquete de PyPI (`pyafipws`) falla al instalar en este entorno (pip rechaza el `.tar.gz` por un symlink que escapa del directorio de extracción, protección real de pip, no un problema nuestro). Se usa `cryptography` para firmar el TRA en CMS/PKCS7 y `zeep` como cliente SOAP — ambos instalan limpio y ya están validados contra el WSDL/servicio REAL de Homologación de AFIP (`FEDummy` respondió en vivo; una autenticación de prueba con certificado autofirmado fue rechazada por AFIP con el error real esperado, confirmando que el circuito TRA→CMS→SOAP funciona de punta a punta). Si en algún momento se reevalúa volver a PyAfipWs, instalarlo desde GitHub (`pip install git+https://github.com/PyAr/pyafipws`), no desde PyPI.
- **Reemplazo de Chilkat / ActiveX:** Eliminar cualquier invocación COM o DLL externa de ChilKat y usar las librerías nativas de Python (`zeep`, `requests`, `cryptography`).
- **Resguardo y Gestión:** Generación de PDF automáticos para guardar en disco y emitir por impresoras Láser.

### Impresión y Reportes
- **Facturas / Recibos / Reportes Láser:** Generación de PDF mediante **ReportLab** (o Jinja2 + WeasyPrint) e impresión posterior.
- **Informes Matriciales:** Soporte para salida de texto plano directas a puerto LPT/USB o secuencias ESC/P si se requiere velocidad matricial.

---

## 4. Flujo de Trabajo Sugerido para Claude Code

Cuando se te indique trabajar sobre un módulo de `FCMENU.vbp`:
1. **Analizar:** Leer el archivo `.frm` o `.bas` correspondiente.
2. **Documentar:** Listar las tablas involucradas de `Fcmenu.mdb`, eventos de interfaz, controles OCX usados y llamadas a la base de datos o AFIP.
3. **Modelar BD:** Si la tabla asociada no existe en PostgreSQL, proponer la clase de SQLAlchemy equivalente.
4. **Construir UI Python:** Generar la vista PyQt6 replicando el layout original.
5. **Implementar Lógica:** Trasladar la lógica del evento VB6 al controlador/servicio de Python.
6. **Revisar con el Usuario:** Mostrar diferencias o preguntar dudas detectadas en el código legacy.