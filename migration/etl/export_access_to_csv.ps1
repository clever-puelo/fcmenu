# Exporta las tablas ya mapeadas en migration/models.py desde fcmenu.mdb
# a CSV, uno por tabla, para cargarlas después a PostgreSQL con
# load_csv_to_postgres.py.
#
# DEBE correrse con PowerShell de 32 bits — el motor DAO de Access solo
# está registrado en COM de 32 bits en este entorno:
#   /c/Windows/SysWOW64/WindowsPowerShell/v1.0/powershell.exe -File export_access_to_csv.ps1
#
# Campos de tipo Fecha (dbDate, type=8) se formatean ISO (yyyy-MM-dd) para
# que el loader de Python los parsee sin ambigüedad. Cotizacion.FECHA es la
# única excepción real: es TEXT "dd/mm/yyyy" en el origen (confirmado vía
# DAO), no una fecha — se copia tal cual, tal como está mapeada en
# models.py (ver CotizacionRepository).

param(
    # Lista opcional de tablas a exportar (por defecto, las 14 completas).
    # Ej.: -Tablas Articulo,Fctabla1 para re-exportar solo un subconjunto.
    [string[]]$Tablas
)

$ErrorActionPreference = "Stop"
$mdbPath = "c:\FcMenu\fcmenu.mdb"
$outDir = "c:\FcMenu\migration\etl\data"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$engine = New-Object -ComObject DAO.DBEngine.36
$db = $engine.OpenDatabase($mdbPath)

# Nombre de tabla en Access -> nombre de archivo CSV de salida.
# (todas coinciden con __tablename__ de models.py)
$tablas = if ($Tablas) { $Tablas } else { @(
    "Clientes", "Articulo", "STOCK", "Fctabla1", "Parametro", "Ctasctes",
    "Imputacion", "FcivaVta", "Fcestad1", "NOTACLTE", "Proveed", "Bancos",
    "Cheques", "Cotizacion", "DtoxClte", "Despachos", "MovStock", "Totales",
    "Efectivo", "MovimVS"
) }

$invariant = [System.Globalization.CultureInfo]::InvariantCulture

function Format-Valor($valor, $esFecha, $esFechaTexto) {
    if ($null -eq $valor -or [DBNull]::Value.Equals($valor)) {
        return $null
    }
    if ($esFecha -and -not $esFechaTexto) {
        return ([DateTime]$valor).ToString("yyyy-MM-dd", $invariant)
    }
    # OJO: la cultura de Windows en este entorno es es-AR (coma decimal).
    # Sin forzar InvariantCulture, Currency/Double salen como "0,0000" en
    # vez de "0.0000" y el loader de Python (Decimal()) rompe. Se fuerza
    # explícitamente formato invariante (punto decimal) para todo valor
    # numérico; el resto (texto) usa ToString() normal.
    if ($valor -is [double] -or $valor -is [decimal] -or $valor -is [single]) {
        return $valor.ToString($invariant)
    }
    # dbBoolean (Articulo.MCALIS, Fctabla1.MCA1-3) llega como [bool] y
    # .ToString() da "True"/"False" — se normaliza a "1"/"0" para que
    # coincida con el tipo Integer del modelo SQLAlchemy.
    if ($valor -is [bool]) {
        return $(if ($valor) { "1" } else { "0" })
    }
    return $valor.ToString()
}

foreach ($nombreTabla in $tablas) {
    Write-Host "Exportando $nombreTabla ..."
    $td = $db.TableDefs($nombreTabla)

    # dbDate = 8. Cotizacion.FECHA es la única columna FECHA que en
    # realidad es texto (type=10) — no entra en este set aunque se llame
    # igual que las demás.
    $camposFecha = @{}
    foreach ($f in $td.Fields) {
        if ($f.Type -eq 8) { $camposFecha[$f.Name] = $true }
    }

    $rs = $db.OpenRecordset("SELECT * FROM [$nombreTabla]", 4) # dbOpenSnapshot = 4, forward-only-friendly
    $nombresCampos = @()
    foreach ($f in $rs.Fields) { $nombresCampos += $f.Name }

    $filas = New-Object System.Collections.Generic.List[PSObject]
    $contador = 0
    while (-not $rs.EOF) {
        $fila = [ordered]@{}
        foreach ($nombreCampo in $nombresCampos) {
            $valor = $rs.Fields.Item($nombreCampo).Value
            $esFecha = $camposFecha.ContainsKey($nombreCampo)
            $esFechaTexto = ($nombreTabla -eq "Cotizacion" -and $nombreCampo -eq "FECHA")
            $fila[$nombreCampo] = Format-Valor $valor $esFecha $esFechaTexto
        }
        $filas.Add([PSCustomObject]$fila)
        $rs.MoveNext()
        $contador++
        if ($contador % 10000 -eq 0) { Write-Host "  ... $contador filas" }
    }
    $rs.Close()

    $outFile = Join-Path $outDir "$nombreTabla.csv"
    $filas | Export-Csv -Path $outFile -NoTypeInformation -Encoding UTF8
    Write-Host "  -> $outFile ($contador filas)"
}

$db.Close()
Write-Host "Listo."
