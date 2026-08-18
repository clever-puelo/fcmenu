VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CargaCC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga de Cuenta Corriente"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.Frame V 
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
      Height          =   3645
      Left            =   420
      TabIndex        =   14
      Top             =   60
      Width           =   6405
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1590
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "0001"
         ToolTipText     =   "Punto de Venta o Prefijo"
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00F9FADC&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "CargaCC.frx":0000
         Left            =   1590
         List            =   "CargaCC.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1140
         Width           =   675
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00F9FADC&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "CargaCC.frx":001A
         Left            =   1590
         List            =   "CargaCC.frx":0033
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   1665
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Borrar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4770
         TabIndex        =   27
         ToolTipText     =   "Elimina definitivamente el Registro de la Cuenta Corriente"
         Top             =   1980
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox Text107 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2610
         MaxLength       =   30
         TabIndex        =   8
         ToolTipText     =   "Tercera Imputación"
         Top             =   2100
         Width           =   435
      End
      Begin VB.TextBox Text106 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2100
         MaxLength       =   30
         TabIndex        =   7
         ToolTipText     =   "Segunda Imputación"
         Top             =   2100
         Width           =   435
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         MaxLength       =   5
         TabIndex        =   0
         ToolTipText     =   "Código de Cliente"
         Top             =   225
         Width           =   1035
      End
      Begin VB.TextBox Text104 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2370
         MaxLength       =   8
         TabIndex        =   5
         ToolTipText     =   "Nro. del Comprobante"
         Top             =   1440
         Width           =   1245
      End
      Begin VB.TextBox Text108 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1605
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   9
         ToolTipText     =   "Resto a Pagar"
         Top             =   2370
         Width           =   2100
      End
      Begin VB.TextBox Text109 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1605
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   10
         ToolTipText     =   "Importe Total del Comprobante"
         Top             =   2640
         Width           =   2100
      End
      Begin VB.TextBox Text105 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1605
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "Primera Imputación"
         Top             =   2100
         Width           =   435
      End
      Begin MSComCtl2.DTPicker FECHA 
         Height          =   330
         Left            =   1590
         TabIndex        =   1
         Top             =   510
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "ddd d MMM yyy"
         Format          =   63045635
         CurrentDate     =   36877
      End
      Begin MSComCtl2.DTPicker FECVTO 
         Height          =   330
         Left            =   1590
         TabIndex        =   11
         ToolTipText     =   "Fecha de Vencimiento"
         Top             =   2940
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "ddd d MMM yyy"
         Format          =   63045635
         CurrentDate     =   36877
      End
      Begin VB.Line Line1 
         X1              =   330
         X2              =   6075
         Y1              =   1875
         Y2              =   1875
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   330
         X2              =   6075
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2100
         TabIndex        =   26
         Top             =   900
         Width           =   2715
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2670
         TabIndex        =   25
         Top             =   270
         Width           =   3525
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Cpbte. :  "
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
         Left            =   150
         TabIndex        =   24
         Top             =   915
         Width           =   1425
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha :"
         Height          =   240
         Left            =   180
         TabIndex        =   23
         Top             =   555
         Width           =   1320
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente :  "
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
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label102 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro. Cpbte.:"
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
         TabIndex        =   21
         Top             =   1470
         Width           =   1395
      End
      Begin VB.Label Label104 
         Alignment       =   1  'Right Justify
         Caption         =   "Resto :"
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
         Left            =   165
         TabIndex        =   20
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label105 
         Alignment       =   1  'Right Justify
         Caption         =   "Importe :"
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
         Left            =   165
         TabIndex        =   19
         Top             =   2655
         Width           =   1335
      End
      Begin VB.Label Label106 
         Alignment       =   1  'Right Justify
         Caption         =   "Fec. Vto. :"
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
         Left            =   165
         TabIndex        =   18
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Letra :  "
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
         Left            =   165
         TabIndex        =   17
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Imputaciones :"
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
         Left            =   75
         TabIndex        =   16
         Top             =   2130
         Width           =   1485
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   4110
         TabIndex        =   15
         Top             =   3030
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3975
      TabIndex        =   12
      Top             =   3810
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5400
      TabIndex        =   13
      Top             =   3810
      Width           =   1320
   End
End
Attribute VB_Name = "CargaCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mensaje As String

Dim dbCCTE As New ADODB.Connection, RgCCTE As New ADODB.Recordset
Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset

Private Sub Form_Load()
Dim I1%

Me.Move 0, 0, 7230, 4600

dbCCTE.ConnectionString = BDatos1
dbCCTE.Open

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

Combo1.ListIndex = 1
Combo2.ListIndex = 0
Text103.Text = "0001"

FECHA.Value = Date
FECVTO.Value = Date

Command2.Enabled = False

If CodCLTE <> 0 Then
        TIPOMov = 3
    Else
        TIPOMov = 1
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
'If KeyCode = 40 Then SendKeys "{TAB}"
'If KeyCode = 38 Then SendKeys "+{TAB}"
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbCCTE.Close
dbCLTE.Close

End Sub

Private Sub Command1_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Command2_Click()
 
  vbMsgBoxTitle = " Carga de Datos a la Tabla "
  vbMsgBoxText = " Desea continuar ?  "
  vbMsgBoxResp = vbYesNoCancel + vbExclamation + vbApplicationModal + vbDefaultButton1
 
  vbResponse = MsgBox(vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle)
 
  Select Case vbResponse
         Case vbYes
              GoTo Grabacion
 
         Case vbNo
              GoTo Salir1
 
         Case vbCancel
              GoTo Salir2
  End Select
 
Grabacion:

'-------------------------------------------------------------

miSQL2 = "SELECT * FROM CtasCtes WHERE CLTE = " & Text101.Text & _
         " AND LETRA = '" & _
         Combo2.List(Combo2.ListIndex) & "' AND TIPO = " & _
         Mid(Combo1.List(Combo1.ListIndex), 1, 1) & " AND CPBTE = " & _
         Val(Text104.Text)
RgCCTE.Open miSQL2, dbCCTE, adOpenDynamic, adLockPessimistic

If RgCCTE.EOF Then
    RgCCTE.AddNew
    RgCCTE!CLTE = Val(Text101.Text)
    RgCCTE!FECHA = FECHA.Value
    RgCCTE!TIPO = Mid(Combo1.List(Combo1.ListIndex), 1, 1)
    RgCCTE!Letra = Combo2.List(Combo2.ListIndex)
    RgCCTE!prefijo = Text103.Text
    RgCCTE!Cpbte = Text104.Text
    RgCCTE!imput4 = 0
    RgCCTE!imput5 = 0
    RgCCTE!imput6 = 0
    RgCCTE!moti = 0
    RgCCTE!tipo9 = 0
    RgCCTE!CVTA = 0
    RgCCTE!bon = 0
    RgCCTE!USUAR = "Arregl"
    End If
    
    RgCCTE!imput1 = Val(Text105.Text)
    RgCCTE!imput2 = Val(Text106.Text)
    RgCCTE!imput3 = Val(Text107.Text)
    RgCCTE!DEBE = CCur(Text108.Text)
    RgCCTE!Impte = CCur(Text109.Text)
    RgCCTE!FECVTO = FECVTO.Value

    RgCCTE.Update

RgCCTE.Close

Salir1:

Text101.SetFocus

Salir2:
   
End Sub

Private Sub Command3_Click()

  vbMsgBoxTitle = " Eliminar Datos de la Tabla "
  vbMsgBoxText = " Desea continuar ?  "
  vbMsgBoxResp = vbYesNoCancel + vbExclamation + vbApplicationModal + vbDefaultButton1
 
  vbResponse = MsgBox(vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle)
 
  Select Case vbResponse
         Case vbYes
              GoTo Grabacion
 
         Case vbNo
              GoTo Salir1
 
         Case vbCancel
              GoTo Salir2
  End Select
 
Grabacion:

miSQL2 = "SELECT * FROM CtasCtes WHERE CLTE = " & Text101.Text & _
         " AND LETRA = '" & _
         Combo2.List(Combo2.ListIndex) & "' AND TIPO = " & _
         Mid(Combo1.List(Combo1.ListIndex), 1, 1) & " AND CPBTE = " & _
         Val(Text104.Text)
RgCCTE.Open miSQL2, dbCCTE, adOpenDynamic, adLockPessimistic

If RgCCTE.EOF Then Exit Sub

RgCCTE.Delete

RgCCTE.Close

Salir1:

Text101.SetFocus

Salir2:

End Sub


Private Sub Text101_GotFocus()

Text101.SelStart = 0
Text101.SelLength = Len(Text101.Text)

Text104.Text = ""
Text105.Text = ""
Text106.Text = ""
Text107.Text = ""
Text108.Text = ""
Text109.Text = ""

End Sub





Private Sub Text104_GotFocus()

Text104.SelStart = 0
Text104.SelLength = Len(Text104.Text)

End Sub

Private Sub Text105_GotFocus()

Text105.SelStart = 0
Text105.SelLength = Len(Text105.Text)

End Sub

Private Sub Text106_GotFocus()

Text106.SelStart = 0
Text106.SelLength = Len(Text106.Text)

End Sub

Private Sub Text107_GotFocus()

Text107.SelStart = 0
Text107.SelLength = Len(Text107.Text)

End Sub

Private Sub Text108_GotFocus()

Text108.SelStart = 0
Text108.SelLength = Len(Text108.Text)

End Sub

Private Sub Text109_GotFocus()

Text109.SelStart = 0
Text109.SelLength = Len(Text109.Text)

End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0
    
miSQL2 = "SELECT NOMB FROM Clientes WHERE CODIGO = " & Val(Text101.Text)
RgCLTE.Open miSQL2, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgCLTE.EOF Then
  vbMsgBoxTitle = "Error de Ingreso"
  vbMsgBoxText = "  " & vbCrLf & "Cliente  NO  Existe  "
  vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
   MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle
   RgCLTE.Close
   Text101.SetFocus
   Exit Sub
    End If
    
Label3.Caption = RgCLTE!Nomb

RgCLTE.Close

End Sub

Private Sub Text104_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0
    
End Sub

Private Sub Text109_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub


Private Sub text104_lostfocus()

'Validar

miSQL2 = "SELECT * FROM CtasCtes WHERE CLTE = " & Text101.Text & _
         "  AND TIPO = " & Mid(Combo1.List(Combo1.ListIndex), 1, 1) & _
         " AND CPBTE = " & Val(Text104.Text)
RgCCTE.Open miSQL2, dbCCTE, adOpenForwardOnly, adLockReadOnly
'Debug.Print miSQL2
If Not RgCCTE.EOF Then
    Label4.Caption = "": Label4.Visible = False
    Text105.Text = RgCCTE!imput1
    Text106.Text = RgCCTE!imput2
    Text107.Text = RgCCTE!imput3
    Text108.Text = RgCCTE!DEBE
    Text109.Text = RgCCTE!Impte
    FECVTO.Value = RgCCTE!FECVTO
    Command3.Visible = True
   GoTo Salirse
    End If
   
    Command3.Visible = False
    
   Label4.Caption = " N U E V O": Label4.Visible = True
    
Text105 = 0
Text106 = 0
Text107 = 0
Text108 = 0
Text109 = 0
FECVTO = Date

Salirse:
   
Text105.SetFocus

RgCCTE.Close
End Sub

Private Sub text105_lostfocus()

Validar

End Sub

Private Sub text106_lostfocus()

Validar

End Sub

Private Sub text107_lostfocus()

Validar

End Sub

Private Sub text108_lostfocus()

Validar

End Sub

Private Sub text109_lostfocus()

Validar

End Sub

'--------------------------------------------
'    Rutina de Validación y consistencia
'--------------------------------------------

Private Sub Validar()

If TIPOMov <> 2 Then Command2.Enabled = False

If Text104.Text = "" Or Text105.Text = "" Or Text106.Text = "" Or _
   Text107.Text = "" Then
    Exit Sub
End If

Command2.Enabled = True
    
End Sub
