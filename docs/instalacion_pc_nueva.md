# FcMenu II — Instalación y primera migración en la PC nueva

Guía para levantar PostgreSQL, migrar los datos reales desde `fcmenu.mdb` y
dejar la app PyQt6 lista para probar, en una PC que hoy no tiene nada de
esto instalado. Todos los valores concretos (versión de Postgres, locale,
timezone, nombres de rol/base) son los que ya están funcionando en esta PC
de desarrollo — replicarlos evita sorpresas de compatibilidad.

**Antes de instalar nada**, dos hallazgos de esta sesión que conviene tener
presentes (detalle completo al final, sección 7):

- El programa de migración (`migrar_completo.py`, hoy empaquetado como
  `migrar_fcmenu.exe`) estaba **roto** desde el 2026-08-20 por un import a
  un archivo que se había borrado — ya corregido, ver sección 7.
- Hay DOS `.exe` — `migrar_fcmenu.exe` (sólo migra datos) y `FcMenuII.exe`
  (la app completa, sección 5). Ninguno de los dos se pudo probar contra
  un Postgres real de punta a punta en esta sesión (no hay acá el
  password de superusuario de `postgres`, a propósito no queda guardado en
  ningún archivo) — `FcMenuII.exe` sí se probó abriendo la ventana real
  contra una base SQLite de prueba, ver sección 5.4. La primera corrida
  real en la PC nueva hay que mirarla con atención en los dos casos.

---

## 1. PostgreSQL

### 1.1. Instalar

Descargar el instalador de PostgreSQL para Windows (el de EDB,
`postgresql.org/download/windows`) — **versión 18**, misma que esta PC
(`psql --version` acá da `18.6`; cualquier 18.x sirve).

Durante la instalación:

| Paso del instalador | Valor a usar |
|---|---|
| Componentes | Server + Command Line Tools (pgAdmin 4 es opcional, no lo usa la app) |
| Puerto | `5432` (default) |
| Configuración regional (locale) | **Spanish, Argentina** — importante, ver nota abajo |
| Password del superusuario `postgres` | El que elijas — **anotalo en un lugar seguro tuyo, no en el repo ni en ningún chat**. Sólo hace falta para crear el rol/base la primera vez (`PG_ADMIN_URL`, sección 2). |

**Por qué el locale importa:** la base real de esta PC (`fcmenu_dev`) quedó
con collation `Spanish_Argentina.1252` (se ve con `SELECT datcollate FROM
pg_database WHERE datname='fcmenu_dev';`). Si el instalador de la PC nueva
usa otro locale, la base que cree `migrar_fcmenu.exe` va a ordenar
texto/acentos distinto — no rompe nada funcionalmente (la app no depende de
`ORDER BY` de SQL para nada crítico, ordena en Python donde importa), pero
mejor mantenerlo igual por las dudas.

### 1.2. Verificar después de instalar

Con `psql -U postgres -h localhost`:

```sql
SHOW timezone;        -- debería dar America/Argentina/Buenos_Aires (u os equivalente) — si no, ver nota abajo
SHOW port;             -- 5432
SELECT version();      -- PostgreSQL 18.x
```

Si `timezone` no sale correcto (el instalador lo autodetecta de la
configuración regional de Windows, normalmente ya viene bien si la PC está
en huso horario Argentina), editar
`C:\Program Files\PostgreSQL\18\data\postgresql.conf`:

```
timezone = 'America/Buenos_Aires'
```

y reiniciar el servicio (`services.msc` → `postgresql-x64-18` → Reiniciar).

### 1.3. `pg_hba.conf` — no hace falta tocarlo

El instalador de Windows ya deja esto por default en
`C:\Program Files\PostgreSQL\18\data\pg_hba.conf` (igual en esta PC, sin
cambios manuales):

```
local   all   all                 scram-sha-256
host    all   all   127.0.0.1/32  scram-sha-256
host    all   all   ::1/128       scram-sha-256
```

Con eso alcanza para conectar desde la propia PC con usuario/password
(que es como se conecta la app, vía `DATABASE_URL`). No hace falta abrir el
puerto a la red ni cambiar `listen_addresses` salvo que la app vaya a
correr en una PC y conectarse a Postgres en OTRA (no es el caso hoy).

### 1.4. Rol de aplicación y base — los crea el propio migrador, no a mano

**No hace falta crear el rol `fcmenu_app` ni la base a mano con `psql`** —
`migrar_fcmenu.exe` (sección 4) hace `DROP ROLE/DATABASE IF EXISTS` +
`CREATE ROLE/DATABASE` + `alembic upgrade head` solo, la primera vez que se
corre (a menos que se le pase `--skip-regen`). Sólo hace falta tener listo
el superusuario `postgres` (paso 1.1) para dárselo como `PG_ADMIN_URL`
(sección 2).

---

## 2. Variables de entorno

Dos formas de setearlas en Windows: **por sesión** (PowerShell, se pierden
al cerrar la ventana — sirven para probar) o **permanentes** (Panel de
Control → Sistema → Configuración avanzada del sistema → Variables de
entorno, o `setx` desde PowerShell — quedan para toda sesión futura,
inclusive para el usuario de Windows que corra la app día a día).

### 2.1. Imprescindibles para la app y para el migrador

| Variable | Uso | Ejemplo |
|---|---|---|
| `DATABASE_URL` | Conexión de la app (rol `fcmenu_app`) — la usan tanto la UI PyQt6 como el migrador (ahí es donde el migrador crea/recrea la base) | `postgresql+psycopg2://fcmenu_app:ELEGIR_PASSWORD@localhost:5432/fcmenu_dev` |

Por sesión (PowerShell):
```powershell
$env:DATABASE_URL = "postgresql+psycopg2://fcmenu_app:ELEGIR_PASSWORD@localhost:5432/fcmenu_dev"
```

Permanente:
```powershell
setx DATABASE_URL "postgresql+psycopg2://fcmenu_app:ELEGIR_PASSWORD@localhost:5432/fcmenu_dev"
```
(`setx` no afecta la ventana de PowerShell donde lo corriste — abrí una
nueva para que tome efecto.)

Si `DATABASE_URL` no está seteada, la app cae sola a un SQLite local
(`fcmenu.sqlite3`, sirve para probar la UI sin Postgres, pero **no** es lo
que se quiere para probar la migración real).

### 2.2. Sólo para correr `migrar_fcmenu.exe` (no la usa la app en uso normal)

| Variable | Uso | Ejemplo |
|---|---|---|
| `PG_ADMIN_URL` | Conexión de superusuario, para que el migrador pueda `DROP`/`CREATE` el rol `fcmenu_app` y la base | `postgresql+psycopg2://postgres:TU_PASSWORD_DE_POSTGRES@localhost:5432/postgres` |

### 2.3. Facturación Electrónica AFIP — dejar en `STUB` por ahora

| Variable | Uso | Valor recomendado para las primeras pruebas |
|---|---|---|
| `FCMENU_AFIP_ENTORNO` | `STUB` (sin conexión real a AFIP, CAE falso — default si no está seteada) / `HOMOLOGACION` (servidores de prueba reales) / `PRODUCCION` | `STUB`, hasta que se decida cablear el cliente real (ver memoria `afip_certificados`) |
| `FCMENU_AFIP_CUIT` | Sólo si `HOMOLOGACION`/`PRODUCCION` — CUIT del emisor sin guiones | `33703467909` |
| `FCMENU_AFIP_CERT_PATH` / `FCMENU_AFIP_KEY_PATH` | Sólo si `HOMOLOGACION`/`PRODUCCION` — rutas al certificado/clave reales de ESE entorno (son pares distintos Homologación/Producción, nunca el mismo certificado sirve para los dos) | No hace falta copiar nada a la PC nueva todavía si se queda en `STUB` |

### 2.4. Automáticas, no hay que setearlas

`USERNAME` — variable estándar de Windows (siempre presente), la app la usa
para estampar "cargado por" en Factura/Recibo/NC/Stock. No requiere
ninguna acción.

---

## 3. Requisitos de la PC

| Componente | ¿Para qué? | Notas |
|---|---|---|
| **PostgreSQL 18** | Base de datos | Sección 1 |
| **Microsoft Access Database Engine 2016 Redistributable (64-bit)** | Que `migrar_fcmenu.exe` pueda leer `fcmenu.mdb` directo por ODBC | Buscar "Access Database Engine 2016 Redistributable" en el sitio de Microsoft, descargar `AccessDatabaseEngine_X64.exe`. **Tiene que ser la de 64 bits** (misma arquitectura que el `.exe`) salvo que la PC ya tenga Office de 32 bits instalado — en ese caso usar la de 32 bits y correr todo (Python/exe) también de 32 bits. No hace falta para usar la app en el día a día, sólo para migrar. |
| **Python 3.14** (sólo si vas a correr la app PyQt6 desde código fuente para probarla, no sólo migrar datos) | UI de Facturación/Cta.Cte. | Ver sección 5 |
| Fuentes Windows (Arial, Arial Black, Lucida Console) | PDFs de Factura/Recibo/NC | **No hace falta instalar nada a mano** — las 4 tipografías que usa `migration/pdf.py` viajan embebidas como archivos `.ttf` en `assets/fonts/` dentro del propio repo, se registran solas al generar el PDF. |

---

## 4. El migrador: `migrar_fcmenu.exe`

### 4.1. Qué es y dónde está

Un único ejecutable standalone (no necesita Python instalado en la PC que
lo corre) que reemplaza al viejo circuito de 2 pasos
(`export_access_to_csv.ps1` en PowerShell de 32-bit + `load_csv_to_postgres.py`)
por un solo comando que lee `fcmenu.mdb` directo vía ODBC.

En esta PC de desarrollo: `c:\FcMenu\dist\migrar_fcmenu.exe` (~32 MB).
Copiarlo a la PC nueva junto con el `fcmenu.mdb` que se quiera migrar (no
hace falta copiar el resto del repo para correrlo — es standalone).

### 4.2. Uso

Con `DATABASE_URL` y `PG_ADMIN_URL` ya seteadas (sección 2):

```powershell
# Primera prueba — liviana: Fcestad1/MovStock vacías, resto recortado a 10 años
.\migrar_fcmenu.exe --mdb C:\ruta\fcmenu.mdb --modo prueba --anios 10

# Corte real completo — TODAS las filas, sin recortes
.\migrar_fcmenu.exe --mdb C:\ruta\fcmenu.mdb --modo completo
```

Sin `--yes`, pide confirmación explícita ("Escribí SI para confirmar")
antes de borrar/recrear la base — es un `DROP DATABASE`, léelo con
atención la primera vez. Al final imprime un conteo de filas por tabla
(Postgres vs. Access) y, en `--modo completo`, un cruce de saldos de
Cliente (`DEUDA` cacheada vs. recalculada desde `Ctasctes`) para detectar
cualquier diferencia real antes de dar la migración por buena.

`--skip-regen` (no borra/recrea la base, sólo recarga datos sobre la
existente) y `--tablas X Y` (subconjunto) también están disponibles si
hace falta re-cargar sólo una parte — ver `--help`.

### 4.3. Cómo se generó / cómo reconstruirlo

No había un `.spec` de PyInstaller en el repo — se armó en esta sesión
(`migrar_fcmenu.spec`, en la raíz). Para reconstruir el `.exe` (ej. después
de un cambio de código), desde la raíz del repo con el `.venv` activo:

```powershell
.venv\Scripts\pip install pyinstaller   # si no está ya
.venv\Scripts\pyinstaller migrar_fcmenu.spec
```

El resultado queda en `dist\migrar_fcmenu.exe`. `build\` y `dist\` ya están
en `.gitignore` — no se versiona el binario, sólo el `.spec` (la receta).

### 4.4. Lo que se pudo y no se pudo verificar en esta sesión

**Verificado:**
- El `.exe` arranca y `--help` muestra las opciones correctas.
- Los 10 archivos que Alembic necesita en disco (`alembic.ini` + `env.py` +
  `script.py.mako` + las 8 migraciones de `migration/alembic/versions/`)
  quedaron empaquetados en la ruta relativa correcta adentro del `.exe`
  (extraje un build intermedio sin comprimir para confirmarlo archivo por
  archivo — Alembic los necesita como archivos reales en disco, no
  alcanza con que el código Python esté "importable" adentro del `.exe`).
- Los 357 tests de `pytest` siguen en verde con el código refactorizado
  (ver sección 7).

**NO verificado** (requiere un superusuario de Postgres que no está
disponible en este entorno, a propósito):
- Una corrida real completa `--mdb fcmenu.mdb --modo prueba` contra un
  Postgres real, de punta a punta (`DROP`/`CREATE` de rol+base, `alembic
  upgrade head`, carga de las 20 tablas, generación de fantasmas,
  recreación de FKs). El circuito equivalente **corriendo desde código
  fuente** (no el `.exe`) sí está extensamente validado en sesiones
  anteriores (ver memoria `cutover_migration_tool`) — lo nuevo y no
  probado es específicamente el empaquetado como `.exe` en sí.

**Recomendación:** la primera corrida en la PC nueva, hacerla con
`--modo prueba --anios 10` (liviana, rápida de revisar) antes que
`--modo completo`, y mirar el resumen final de conteos con atención.

---

## 5. La app completa: `FcMenuII.exe`

### 5.1. Qué es y dónde está

El ERP completo (Facturador, Recibos, NC, Cta.Cte., Stock, Listados, etc.
— reemplaza a `FCMENU.frm`/`FCMENU.vbp` enteros) empaquetado como un único
`.exe` standalone, igual de autocontenido que `migrar_fcmenu.exe` — no
necesita Python instalado en la PC que lo corre.

En esta PC de desarrollo: `c:\FcMenu\dist\FcMenuII.exe` (~128 MB — mucho
más grande que el migrador porque además de PyQt6 completo empaqueta
`reportlab`/`cryptography`/`zeep` para los PDF y AFIP). Copiarlo a la PC
nueva; sólo necesita `DATABASE_URL` seteada (sección 2.1) apuntando a la
base que ya migró `migrar_fcmenu.exe` — nada más.

### 5.2. Uso

```powershell
# Con DATABASE_URL ya seteada (sección 2.1):
.\FcMenuII.exe
```

Arranca directo en `MainMenuWindow` — mismo comportamiento que corriendo
`.venv\Scripts\python.exe -m migration.ui.main_menu` desde código fuente.

**Es la versión "de consola" a propósito** (con la ventana negra de fondo
detrás de la app) — pensada para esta primera ronda de pruebas en la PC
nueva: si algo falla al conectar a Postgres o cualquier otro error, el
traceback real queda visible ahí en vez de perderse en silencio. Una vez
confirmado que todo anda bien, se puede regenerar una versión "limpia"
sin esa consola (un solo flag distinto al reconstruir, `console=False` en
`fcmenu.spec` — avisar si se quiere esa versión).

### 5.3. Cómo se generó / cómo reconstruirlo

`fcmenu.spec` (raíz del repo) + `run_app.py` (wrapper de entrada — hace
falta porque `migration/ui/main_menu.py` usa imports relativos, que
rompen si PyInstaller lo toma directo como script top-level; `run_app.py`
sólo hace `from migration.ui.main_menu import main; main()` desde
AFUERA del paquete, con lo que los imports relativos internos vuelven a
funcionar normal). Reconstruir, desde la raíz del repo con el `.venv`
activo:

```powershell
.venv\Scripts\pyinstaller fcmenu.spec
```

Resultado en `dist\FcMenuII.exe`.

### 5.4. Verificado en esta sesión

Corrida real (no sólo `--help`, a diferencia del primer intento con el
migrador): el `.exe` **abre la ventana real de `MainMenuWindow`** ("FcMenu
II — Facturación y Cuenta Corriente (Electrónica)", confirmado con el
título de la ventana en el proceso vivo), con una base SQLite de prueba
recién creada (`alembic`/`init_db()`, no la real). Encontrado y corregido
en el camino:
- El primer intento de build rompía al arrancar (`ImportError: attempted
  relative import with no known parent package`) — arreglado con
  `run_app.py` (ver 5.3).
- `AfipWSFEv1Cliente._archivo_cache_ta()` (`migration/afip.py`) resolvía
  el cache del Ticket de Acceso de WSAA con una ruta relativa al código
  (`Path(__file__)...`), que adentro de un `.exe` congelado apunta a una
  carpeta TEMPORAL que se borra al cerrar el programa — inútil para un
  cache que tiene que sobrevivir reinicios. Corregido para usar la
  carpeta del propio `.exe` cuando está congelado (sólo importa si más
  adelante se cablea el cliente real de AFIP — hoy sigue en `STUB` por
  default, no lo usa).

**NO verificado**: contra un Postgres real (mismo motivo que el
migrador — no hay superusuario disponible en este entorno), ni el
circuito de AFIP real (sigue en `STUB` por default), ni impresión física
en una impresora Láser real. La UI en sí (ventanas, PDF, cálculos) está
extensamente probada en sesiones anteriores corriendo desde código
fuente — lo nuevo acá es específicamente el empaquetado como `.exe`.

---

## 6. Checklist resumido

1. [ ] Instalar PostgreSQL 18 (locale Spanish-Argentina, anotar password de `postgres`).
2. [ ] Verificar `SHOW timezone;` = Buenos Aires (corregir en `postgresql.conf` si no).
3. [ ] Instalar Microsoft Access Database Engine 2016 Redistributable (64-bit).
4. [ ] Copiar `dist\migrar_fcmenu.exe` + `dist\FcMenuII.exe` + `fcmenu.mdb` a la PC nueva.
5. [ ] Setear `DATABASE_URL` y `PG_ADMIN_URL` (sección 2).
6. [ ] Correr `migrar_fcmenu.exe --mdb <ruta> --modo prueba --anios 10`, revisar el resumen.
7. [ ] Si todo cierra, repetir con `--modo completo`.
8. [ ] Correr `FcMenuII.exe` (sección 5) para probar la app completa contra los datos ya migrados.

---

## 7. Hallazgo real: el migrador estaba roto — ya corregido

Mientras se armaba esta guía y el `.exe`, `migrar_fcmenu.exe --help`
tiraba `ModuleNotFoundError: No module named
'migration.etl.load_csv_to_postgres'` al primer intento de import.

**Causa:** el commit `0ebc3be` ("Delete migration/etl directory",
2026-08-20) borró todo `migration/etl/` — correcto para los CSV con datos
reales de clientes (`migration/etl/data/*.csv`, la razón real del commit,
ver hallazgo de seguridad en memoria de sesión), pero de paso también borró
`load_csv_to_postgres.py` y `export_access_to_csv.ps1`, que son **código**,
no datos. El problema: `load_mdb_to_postgres.py` (el módulo nuevo,
agregado el mismo día) todavía importaba de ahí `TABLAS`, `ConfigTabla`,
el manejo de FKs y la generación de Clientes/Artículos "fantasma" — código
compartido entre el camino viejo (CSV) y el nuevo (ODBC directo) que nunca
se movió a un lugar propio. Resultado: **el migrador quedó
100% roto desde ese commit**, y nadie lo notó porque no se había vuelto a
correr hasta ahora.

**Arreglado:**
- Se restauró `migration/etl/load_csv_to_postgres.py` y
  `migration/etl/export_access_to_csv.ps1` desde el historial de git
  (`git show 0ebc3be~1:...`) — sólo el código, **no** los CSV de datos
  reales (esos siguen afuera y protegidos por `.gitignore`, la exclusión
  sigue vigente).
- Se refactorizó `migrar_completo.py`: en vez de lanzar un subproceso
  (`sys.executable` + la ruta a otro archivo `.py`) para llamar a
  `load_mdb_to_postgres.py`/`load_csv_to_postgres.py`, ahora los llama
  como funciones Python directas (`cargar(...)`, nueva en ambos módulos).
  El subproceso no sólo era más lento (² procesos Python, ² conexiones a
  Postgres) — **no puede funcionar dentro de un `.exe` empaquetado**, que
  no tiene un intérprete Python ni archivos `.py` sueltos para invocar así.
  Este cambio fue, a la vez, la corrección del bug Y el requisito para
  poder generar el `.exe` en primer lugar.

Verificado: 357/357 tests en verde, `migrar_completo.py`/`load_mdb_to_postgres.py`/
`load_csv_to_postgres.py` importan y corren `--help` sin error, y el `.exe`
generado con el código corregido arranca limpio.

**No commiteado todavía** — queda en el working tree del repo para que lo
revises antes, junto con el resto de los cambios de esta sesión (fixes de
cotización y fuentes de PDF, ver conversación).
