$engine = New-Object -ComObject DAO.DBEngine.36
$db = $engine.OpenDatabase('c:\FcMenu\fcmenu.mdb')
$tables = @('Clientes','Articulo','STOCK','Fctabla1','Parametro','Ctasctes','FcivaVta','Fcestad1','NOTACLTE','Proveed','Bancos','Cheques','Cotizacion')
foreach ($name in $tables) {
    $td = $db.TableDefs[$name]
    if ($null -ne $td) {
        Write-Host "## $name"
        foreach ($f in $td.Fields) {
            Write-Host ($f.Name + ' : ' + $f.Type)
        }
        Write-Host ''
    }
}
$db.Close()
