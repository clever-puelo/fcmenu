VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ABMItem 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambios a los Items"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8235
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox List3 
      BackColor       =   &H00FEFADE&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "ABMItem.frx":0000
      Left            =   1710
      List            =   "ABMItem.frx":000D
      TabIndex        =   30
      Text            =   "1-Importe"
      Top             =   1395
      Width           =   1710
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Alta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Dar de Alta un Cliente nuevo"
      Top             =   4095
      Width           =   1035
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Baja"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2235
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Buscar un Cliente"
      Top             =   4095
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Notas"
      Height          =   315
      Left            =   180
      TabIndex        =   12
      Top             =   4095
      Width           =   1300
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   6720
      TabIndex        =   8
      Top             =   4095
      Width           =   1300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   315
      Left            =   5310
      TabIndex        =   7
      Top             =   4095
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Datos del Ítem"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3990
      Left            =   105
      TabIndex        =   13
      Top             =   0
      Width           =   7965
      Begin VB.ComboBox List4 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "ABMItem.frx":0036
         Left            =   1605
         List            =   "ABMItem.frx":0046
         TabIndex        =   31
         Text            =   "1-Peso"
         Top             =   1710
         Width           =   1710
      End
      Begin VB.ComboBox List2 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "ABMItem.frx":006C
         Left            =   1605
         List            =   "ABMItem.frx":0076
         TabIndex        =   29
         Text            =   "1-Sin Elaborar"
         Top             =   1080
         Width           =   1710
      End
      Begin VB.ComboBox List1 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "ABMItem.frx":0097
         Left            =   1605
         List            =   "ABMItem.frx":00AD
         TabIndex        =   28
         Text            =   "1-Materia Prima"
         Top             =   765
         Width           =   1710
      End
      Begin VB.TextBox Text108 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   6
         ToolTipText     =   "Stock de Reposición (Items sin elaborar)"
         Top             =   3525
         Width           =   1980
      End
      Begin VB.TextBox Text107 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Stock Máximo (Items sin elaborar)"
         Top             =   3225
         Width           =   1980
      End
      Begin VB.TextBox Text106 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Stock Mínimo (Items sin elaborar)"
         Top             =   2925
         Width           =   1965
      End
      Begin VB.TextBox Text105 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Precio Unitario (En Items sin elaborar)"
         Top             =   2625
         Width           =   1950
      End
      Begin VB.TextBox Text104 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   2
         ToolTipText     =   "Cantidad a preparar (el items elaborados)"
         Top             =   2325
         Width           =   1950
      End
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         MaxLength       =   6
         TabIndex        =   1
         ToolTipText     =   "Unidad de Medida"
         Top             =   2025
         Width           =   885
      End
      Begin VB.TextBox Text102 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         MaxLength       =   50
         TabIndex        =   0
         ToolTipText     =   "Razón Social del Cliente"
         Top             =   480
         Width           =   4515
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1605
         MaxLength       =   5
         TabIndex        =   9
         ToolTipText     =   "Código de Cliente"
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Stock Rep.     :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   26
         Top             =   3570
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Stock Máx.     :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   3270
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "Stock Mín.      :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   24
         Top             =   2970
         Width           =   1470
      End
      Begin VB.Label Label3 
         Caption         =   "Nivel              :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   23
         Top             =   870
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Cant. a Prep. :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   22
         ToolTipText     =   "Cantidad a Preparar"
         Top             =   2370
         Width           =   1395
      End
      Begin VB.Label Label109 
         Caption         =   "Precio                  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   21
         Top             =   2670
         Width           =   1470
      End
      Begin VB.Label Label108 
         Caption         =   "Unid.Medida  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   20
         Top             =   2070
         Width           =   1395
      End
      Begin VB.Label Label107b 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5265
         TabIndex        =   19
         Top             =   1920
         Width           =   2505
      End
      Begin VB.Label Label106 
         Caption         =   "Moneda               :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   18
         Top             =   1770
         Width           =   1455
      End
      Begin VB.Label Label105 
         Caption         =   "Impte./Porc.      :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   17
         Top             =   1470
         Width           =   1440
      End
      Begin VB.Label Label104 
         Caption         =   "Tipo                     :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   1170
         Width           =   1410
      End
      Begin VB.Label Label102 
         Caption         =   "Descripción   :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   15
         Top             =   570
         Width           =   1395
      End
      Begin VB.Label Label101 
         Caption         =   "Código           :  "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   14
         Top             =   285
         Width           =   1425
      End
   End
   Begin MSAdodcLib.Adodc CST1 
      Height          =   330
      Left            =   3630
      Top             =   4725
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=CSDAT"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "CSDAT"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "CST1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "ABMItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mensaje As String
Dim CAMPOS      As String
Dim s$
Dim CANAL As String * 2

Private Sub Form_Load()
Dim I1%

Me.Move 0, 0, 8300, 4850
   
   CSMenu.Toolbar2.Buttons(3).Value = tbrPressed

Command3.Visible = False
Command5.Visible = False


Select Case TipoMov

Case 1

    Blanquea
    ABMItem.Caption = "Alta de un Item"

Case 2, 3
    
    ABMItem.Caption = "Cambios a un Item"

    Command3.Visible = True
    Command5.Visible = True
    
    CST1.RecordSource = "SELECT  * FROM Items WHERE COD = " & CodITEM
    CST1.Refresh
    If CST1.Recordset.EOF = True Then Exit Sub

    Text101.Enabled = False
    Text101.Text = CST1.Recordset!COD
    Text102.Text = CST1.Recordset!DESCRI
    Text103.Text = CST1.Recordset!UMedi
    Text104.Text = CST1.Recordset!CANT
    Text105.Text = CST1.Recordset!Prec
    Text106.Text = CST1.Recordset!stmin
    Text107.Text = CST1.Recordset!stmax
    Text108.Text = CST1.Recordset!densid
    List1.ListIndex = CST1.Recordset!NIVEL - 1
    List2.ListIndex = CST1.Recordset!TIPO - 1
    List3.ListIndex = CST1.Recordset!porc - 1
    List4.ListIndex = CST1.Recordset!moneda - 1
        
End Select

VerSiVa

If TipoMov = 2 Then
                ABMItem.Caption = "Baja de un Item"
                Text101.Enabled = False
                Text102.Enabled = False
                Text103.Enabled = False
                Text104.Enabled = False
                Text105.Enabled = False
                Text106.Enabled = False
                Text107.Enabled = False
                Text108.Enabled = False
                List1.Enabled = False
                List2.Enabled = False
                List3.Enabled = False
                List4.Enabled = False
                
                Command2.Enabled = True
                Command2.Caption = "Eliminar"
End If

' If NotasITEM Then NOTAITEM.Show


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF2 Then
    BusItems.Show
    Unload Me
        End If


'If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 39 Then SendKeys "{TAB}"
'If KeyCode = 38 Or KeyCode = 37 Then SendKeys "+{TAB}"
If KeyCode = 27 Then End

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

Grabacion
Unload Me

CONSULTA.DoVer1

End Sub

Private Sub Command3_Click()

NOTAITEM.Show

End Sub

Private Sub Command4_Click()

TipoMov = 2
Form_Load

End Sub

Private Sub Command5_Click()

Command3.Visible = False
Command5.Visible = False

    Blanquea

End Sub


Private Sub Form_Unload(Cancel As Integer)

    CSMenu.Toolbar2.Buttons(3).Value = tbrUnpressed

End Sub

Private Sub List2_ItemCheck(Item As Integer)

VerSiVa

End Sub

Private Sub List2_LostFocus()

VerSiVa

End Sub

'*************************************
'   Ingresa Código de Item
'*************************************

Private Sub Text101_GotFocus()

Text101.SelStart = 0
Text101.SelLength = Len(Text101.Text)

End Sub

Private Sub Text102_GotFocus()

Text102.SelStart = 0
Text102.SelLength = Len(Text102.Text)

End Sub

Private Sub Text104_GotFocus()

Text104.SelStart = 0
Text104.SelLength = Len(Text104.Text)

End Sub

Private Sub Text105_GotFocus()

Text105.SelStart = 0
Text105.SelLength = Len(Text105.Text)

End Sub

Private Sub Text101_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode <> vbKeyReturn Then Exit Sub
    
    CST1.RecordSource = "SELECT * FROM Items WHERE CODIGO = " & Val(Text101.Text)
    CST1.Refresh

If Not CST1.Recordset.EOF Then
  vbMsgBoxTitle = "Error de Ingreso"
  vbMsgBoxText = "  " & vbCrLf & "Código  YA  Existe  "
  vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
   MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle
   Text101.SetFocus
   Exit Sub
    End If

End Sub


'*************************************
'   Ingresa Cod.Postal
'*************************************

Private Sub Text106_GotFocus()

Text106.SelStart = 0
Text106.SelLength = Len(Text106.Text)

End Sub

'*************************************
'   Ingresa Provincia
'*************************************

Private Sub Text107_GotFocus()

Text107.SelStart = 0
Text107.SelLength = Len(Text107.Text)

End Sub

Private Sub Text108_KeyPress(KeyAscii As Integer)

        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
            KeyAscii = KeyAscii + 32
            End If

End Sub

Private Sub Grabacion()
Dim I1%, I2%, I3%

                Me.Hide

Select Case TipoMov
        Case 1
                Mensaje = "Desea Grabar ?"
        Case 2
                Mensaje = "Desea Eliminar ?"
        Case 3
                Mensaje = "Desea Cambiar ? "
End Select

Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "ABM de ITEMS")

If Respuesta = vbNo Then Exit Sub
If Respuesta = vbCancel Then
         Exit Sub
         Me.Show
         End If

If TipoMov = 2 Then
        CST1.Recordset.Delete
        Exit Sub
        End If
        
     
'******************************************
'           Graba en Items
'******************************************

CST1.RecordSource = "SELECT  * FROM ITEMS WHERE COD = " & Val(Text101.Text)
CST1.Refresh

    If CST1.Recordset.EOF = True Then
        CST1.Recordset.AddNew
        
        CST1.Recordset!COD = Val(Text101.Text)
        CST1.Recordset!DESCRI = ""
    End If
    
    CST1.Recordset!DESCRI = Text102.Text
    CST1.Recordset!UMedi = Text103.Text
    CST1.Recordset!CANT = Val(Replace(Text104.Text, ",", "."))
    CST1.Recordset!Prec = Val(Replace(Text105.Text, ",", "."))
    CST1.Recordset!stmin = Val(Replace(Text106.Text, ",", "."))
    CST1.Recordset!stmax = Val(Replace(Text107.Text, ",", "."))
    CST1.Recordset!densid = Val(Replace(Text108.Text, ",", "."))

CST1.Recordset!NIVEL = Val(Mid(List1.Text, 1, 1))
CST1.Recordset!TIPO = Val(Mid(List2.Text, 1, 1))
CST1.Recordset!porc = Val(Mid(List3.Text, 1, 2))
CST1.Recordset!moneda = Val(Mid(List4.Text, 1, 2))
CST1.Recordset!baja = 0

CST1.Recordset!FACTUAL = Date
CST1.Recordset!USUARIO = Mid$(Red_Usuario, 1, 6)

CST1.Recordset.Update

End Sub

Sub MuestraError()
          
        vbMsgBoxTitle = "Error de Ingreso"
        vbMsgBoxText = "  " & vbCrLf & "Código  NO  Existe  "
        vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
        MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle

End Sub

Sub Blanquea()
    
    CST1.RecordSource = "SELECT  TOP 1 COD FROM Items WHERE COD < 99900 " _
                                & "ORDER BY COD DESC"
    
    CST1.Refresh
    If CST1.Recordset.EOF = True Then
'        CST1.Recordset!CODIGO = 0
    End If
    If CST1.Recordset.RecordCount = 0 Then
        Text101.Text = 1
    Else
        Text101.Text = CST1.Recordset!COD + 1
    End If
    
    Text102.Text = ""
    Text103.Text = ""
    Text104.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text107.Text = ""
    Text108.Text = ""
    List1.ListIndex = 0
    List2.ListIndex = 0
    List3.ListIndex = 0
    List4.ListIndex = 0

Text101.Enabled = True
Label1.Enabled = False
Label109.Enabled = False

Text102.SetFocus

End Sub


Private Sub VerSiVa()

If Val(Mid(List2.Text, 1, 1)) = 1 Then
        Label1.Enabled = False
        Text104.Enabled = False
        Label109.Enabled = True
        Text105.Enabled = True
        Label4.Enabled = True
        Text106.Enabled = True
        Label2.Enabled = True
        Text107.Enabled = True
        Label5.Enabled = True
        Text108.Enabled = True
Else
        Label109.Enabled = False
        Text105.Enabled = False
        Label4.Enabled = False
        Text106.Enabled = False
        Label2.Enabled = False
        Text107.Enabled = False
        Label5.Enabled = False
        Text108.Enabled = False
        Label1.Enabled = True
        Text104.Enabled = True
End If

End Sub
