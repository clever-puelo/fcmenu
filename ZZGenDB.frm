VERSION 5.00
Begin VB.Form ZZGenDB 
   Caption         =   "Generador de Base de Datos"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Ejecutar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   930
      TabIndex        =   2
      Top             =   3390
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3450
      TabIndex        =   1
      Top             =   3630
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4860
      TabIndex        =   0
      Top             =   3630
      Width           =   1230
   End
End
Attribute VB_Name = "ZZGenDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const Jet10 = 1
Const Jet11 = 2
Const Jet20 = 3
Const Jet3x = 4
Const Jet4x = 5

Private Sub Form_Load()

If (Params(3) <> "A" And Params(3) <> "S") Or _
    Params(4) = "" Then
    MsgBox "Faltan Parámetros o estan mal Ingresados", vbCritical, "FCMENU - Error Grave"
    End
    Exit Sub
End If

End Sub

Private Sub Command1_Click()
    Dim cat As ADOX.Catalog
    Dim tbl As ADOX.Table
    '
 '   On Error GoTo ErrCrearTabla
    '
    Set cat = New ADOX.Catalog
    Set tbl = New ADOX.Table

If TipoDB = "A" Then
 ' Si existe la base de datos, preguntar si se quiere borrarla
    If Len(Dir$(Params(4) & "FCMENU.MDB")) Then
        If MsgBox("La base de datos ya existe." & vbCrLf & _
                  "Desea Borrarla?", _
                    vbQuestion + vbYesNo + vbDefaultButton2 _
            ) = vbNo Then
            GoTo AlResto
        Else
            Kill (Params(4) & "FCMENU.MDB")
        End If
    End If
    
'Crea una Base Access2000
    CreateNewMDB Params(4) & "FCMENU.MDB", Jet4x
End If

AlResto:

Err = CreateNewKey(HKEY_CURRENT_USER, "FCMENU\Datos\DSN")
Err = CreateNewKey(HKEY_CURRENT_USER, "FCMENU\Datos\Backup")
Err = CreateNewKey(HKEY_CURRENT_USER, "FCMENU\Datos\Impresora")

Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos", "Driver", Params(3), REG_SZ)

Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Nombre", "FCMENU", REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Descrip", "Base de Datos de Facturación", REG_SZ)

If Params(3) = "A" Then
    Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Carpeta", Params(4), REG_SZ)
Else
    Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Servidor", Params(4), REG_SZ)
End If

Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Nombre", "Fmanual", REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Descrip", "Facturas Manuales", REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Carpeta", "e:\", REG_SZ)

Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Impresora", "1", "LPT1:", REG_SZ)


' SaveSetting appname:="FCMENU", Section:="Datos", Key:="Driver", setting:="A"

Inicio.Show
    '
    ' Abrir el catálogo
    cat.ActiveConnection = _
            "Provider=Microsoft.Jet.OLEDB.4.0;" & ";" & _
            "Data Source=" & (Params(4) & "FCMENU.MDB") & ";"
    '
    ' Crear la nueva tabla
'    With tbl
'        .Name = "Prueba"
'        .Columns.Append "ID", adInteger
'    End With
    '
    ' Añadir la nueva tabla a la base de datos
'    cat.Tables.Append tbl


End
    
End Sub

Private Sub Command2_Click()

End


End Sub

Private Sub Command3_Click()
    ' Crear una base de datos usando los datoos indicados en:       (28/Sep/01)
    ' txtNombreBase     El nombre (y path) de la base de datos
    ' cboProvider       El proveedor de la base de datos
    '
    Dim cat As ADOX.Catalog
    '
    ' Si se produce un error: interceptarlo
    On Error GoTo ErrCrearBase
    '
    Set cat = New ADOX.Catalog
    '
    ' Si existe la base de datos, preguntar si queremos borrarla
    If Len(Dir$(Params(4) & "FCMENU.MDB")) Then
        If MsgBox("La base de datos ya existe." & vbCrLf & _
                  "¿Desea Borrarla?", _
                    vbQuestion + vbYesNo + vbDefaultButton2 _
            ) = vbNo Then
            Exit Sub
        Else
            Kill (Params(4) & "FCMENU.MDB")
        End If
    End If
    '
    ' Crear la base de datos
    cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & (Params(4) & "FCMENU.MDB") & ";"
    '
    MsgBox "Base de datos creada satisfactoriamente."
    '
    ' Para no "colarnos" en la rutina de error
    Exit Sub
    '
ErrCrearBase:
    ' Mostrar el mensaje de error
    MsgBox "Error al crear la base de datos:" & vbCrLf & _
            Err.Number & " " & Err.Description, _
            vbExclamation, "Error al crear la base de datos"
    Err.Clear
End Sub

Sub CreateNewMDB(FileName, Format)
  Dim Catalog
  
  Set Catalog = CreateObject("ADOX.Catalog")
  Catalog.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & _
     "Jet OLEDB:Engine Type=" & Format & _
    ";Data Source=" & FileName
End Sub
Private Sub cmdCrearTabla_Click()

    ' Crear una tabla en la base de datos indicada                  (28/Sep/01)
    Dim cat As ADOX.Catalog
    Dim tbl As ADOX.Table
    '
    On Error GoTo ErrCrearTabla
    '
    Set cat = New ADOX.Catalog
    Set tbl = New ADOX.Table
    '
    ' Abrir el catálogo
    cat.ActiveConnection = _
            "Provider=Microsoft.Jet.OLEDB.4.0;" & ";" & _
            "Data Source=" & (Params(4) & "FCMENU.MDB") & ";"
    '
    ' Crear la nueva tabla
    With tbl
        .Name = "Clientes"
        ' Crear los campos y añadirlos a la tabla.
        ' Esto hay que hacerlo antes de añadir la tabla a la colección de tablas
        '
        .Columns.Append "ID", adInteger
            ' Para Access 2000
            .Columns.Append "Nombre", adVarWChar, 50        ' Una cadena de 50 caracteres
            .Columns.Append "email", adVarWChar, 100
            .Columns.Append "Telefono", adVarWChar
            .Columns.Append "Observaciones", adLongVarWChar ' Una cadena larga, (Memo)
        
        .Columns("Nombre").Attributes = adColNullable       ' Permite contener nulos
        .Columns("email").Attributes = adColNullable
        .Columns("Telefono").Attributes = adColNullable
        .Columns("Observaciones").Attributes = adColNullable
    End With
    '
    ' Añadir la nueva tabla a la base de datos
    cat.Tables.Append tbl
    '
SalirCrearTabla:
    Set tbl = Nothing
    Set cat = Nothing
    '
    Exit Sub
    '
ErrCrearTabla:
    ' Mostrar el mensaje de error
    MsgBox "Error al crear la tabla:" & vbCrLf & _
            Err.Number & " " & Err.Description, _
            vbExclamation, "Error al crear la tabla"
    Err.Clear
    Resume SalirCrearTabla
End Sub
