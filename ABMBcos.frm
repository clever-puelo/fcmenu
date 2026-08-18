VERSION 5.00
Begin VB.Form ABMBcos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   A B M  a Bancos"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
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
      Left            =   510
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Dar de Alta un Cliente nuevo"
      Top             =   2985
      Width           =   1035
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cambio"
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
      Left            =   1650
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Buscar un Cliente"
      Top             =   2985
      Width           =   1035
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
      Left            =   5085
      TabIndex        =   11
      Top             =   2985
      Width           =   1320
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
      Left            =   3660
      TabIndex        =   10
      Top             =   2985
      Width           =   1320
   End
   Begin VB.Frame Frame1 
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
      Height          =   2835
      Left            =   75
      TabIndex        =   14
      Top             =   75
      Width           =   6375
      Begin VB.TextBox Text109 
         BackColor       =   &H00FEFADE&
         Enabled         =   0   'False
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
         Left            =   4320
         LinkTimeout     =   5
         MaxLength       =   3
         TabIndex        =   9
         ToolTipText     =   "Clearing"
         Top             =   2385
         Width           =   525
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
         Left            =   1575
         MaxLength       =   30
         TabIndex        =   3
         ToolTipText     =   "Nombre de la Sucursal"
         Top             =   1035
         Width           =   4515
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Propio          :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   8
         ToolTipText     =   "Si es Banco Propio"
         Top             =   2385
         Width           =   1665
      End
      Begin VB.TextBox Text102 
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
         TabIndex        =   1
         ToolTipText     =   "Código de Sucursal"
         Top             =   495
         Width           =   1035
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
         Height          =   285
         Left            =   5580
         LinkTimeout     =   5
         MaxLength       =   2
         TabIndex        =   7
         ToolTipText     =   "Código de Provincia"
         Top             =   1845
         Visible         =   0   'False
         Width           =   525
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
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   10
         TabIndex        =   6
         ToolTipText     =   "Código Postal"
         Top             =   1845
         Width           =   1020
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
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Localidad del Banco/Sucursal"
         Top             =   1575
         Width           =   4530
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
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Dirección del Banco/Sucursal"
         Top             =   1305
         Width           =   4530
      End
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
         Left            =   1575
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Razón Social del Banco"
         Top             =   765
         Width           =   4515
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
         ToolTipText     =   "Código de Banco"
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   3600
         TabIndex        =   24
         Top             =   360
         Width           =   2565
      End
      Begin VB.Label Label3 
         Caption         =   "Clearing  :"
         Enabled         =   0   'False
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
         Left            =   3420
         TabIndex        =   23
         Top             =   2385
         Width           =   930
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   255
         X2              =   6000
         Y1              =   2270
         Y2              =   2270
      End
      Begin VB.Line Line1 
         X1              =   255
         X2              =   6000
         Y1              =   2250
         Y2              =   2250
      End
      Begin VB.Label Label2 
         Caption         =   "Nomb.Suc.    :"
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
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Cód. Suc.      :  "
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
         Top             =   540
         Width           =   1425
      End
      Begin VB.Label Label107a 
         Caption         =   "Provincia          :"
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
         Left            =   4230
         TabIndex        =   20
         Top             =   1890
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label106 
         Caption         =   "Cod. Postal        :"
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
         TabIndex        =   19
         Top             =   1890
         Width           =   1335
      End
      Begin VB.Label Label105 
         Caption         =   "Localidad           :"
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
         Top             =   1605
         Width           =   1335
      End
      Begin VB.Label Label104 
         Caption         =   "Dirección           :"
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
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label102 
         Caption         =   "Razón Social :"
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
         TabIndex        =   16
         Top             =   810
         Width           =   1395
      End
      Begin VB.Label Label101 
         Caption         =   "Código          :  "
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
         Top             =   270
         Width           =   1425
      End
   End
End
Attribute VB_Name = "ABMBcos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mensaje As String

Dim dbCYB As New ADODB.Connection, RgCYB As New ADODB.Recordset

Private Sub Form_Load()
Dim i1%

Me.Move 0, 0, 6650, 3800

dbCYB.ConnectionString = BDatos1
dbCYB.Open

Command2.Enabled = False
Command4.Visible = False
Command5.Visible = False

'FCMENU.Toolbar2.Buttons(7).Value = tbrPressed

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

Private Sub Form_Terminate()

ParaBcos = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

'FCMENU.Toolbar2.Buttons(7).Value = tbrUnpressed
dbCYB.Close

End Sub

Private Sub Command1_Click()

Unload Me

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

miSQL2 = "SELECT * FROM Bancos WHERE COD = " & Val(Text101.Text)
RgCYB.Open miSQL2, dbCYB, adOpenDynamic, adLockPessimistic

If RgCYB.EOF Then
    RgCYB.AddNew
    RgCYB.Fields!COD = Val(Text101.Text)
    RgCYB.Fields!SUC = 0
    RgCYB.Fields!NOMBRE = ""
    RgCYB.Fields!NOMSUC = ""
    RgCYB.Fields!Dir = ""
    RgCYB.Fields!Loc = ""
'    RGCYB.Fields!Prov = ""
    RgCYB.Fields!cp = ""
    RgCYB.Fields!propio = 0
    RgCYB.Fields!clearing = 0
    End If
    
    RgCYB.Fields!COD = Val(Text101.Text)
    RgCYB.Fields!SUC = Val(Text102.Text)
    RgCYB.Fields!NOMBRE = Text103.Text
    RgCYB.Fields!NOMSUC = Text104.Text
    RgCYB.Fields!Dir = Text105.Text
    RgCYB.Fields!Loc = Text106.Text
'    RGCYB.Fields!PCIA = Text108.Text
    RgCYB.Fields!cp = Text107.Text
    RgCYB.Fields!propio = Check1.Value
    RgCYB.Fields!clearing = Val(Text109.Text)

    RgCYB.Update

   If ParaBcos Then
        DetPago.Label5.Caption = Text103.Text
        DetPago.DTPicker1.SetFocus
   End If

RgCYB.Close

Salir1:

Unload Me

Salir2:
   
End Sub

Private Sub Command4_Click()
 
    Label4.Caption = "B A J A"
 
  vbMsgBoxTitle = " Elimina Datos de la Tabla "
  vbMsgBoxText = " Desea Eliminar el Registro ?  "
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

   RgCYB.Delete
    
Salir1:

Unload Me

Salir2:

End Sub

Private Sub Text101_GotFocus()

Text101.SelStart = 0
Text101.SelLength = Len(Text101.Text)

Command4.Visible = False

Text102.Text = ""
Text103.Text = ""
Text104.Text = ""
Text105.Text = ""
Text106.Text = ""
Text107.Text = ""
'Text108.Text = ""
Text109.Text = ""

Check1.Value = 0

End Sub

Private Sub Text102_GotFocus()

Text102.SelStart = 0
Text102.SelLength = Len(Text102.Text)

End Sub

Private Sub Text103_GotFocus()

Text103.SelStart = 0
Text103.SelLength = Len(Text103.Text)

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

Private Sub Check1_Click()

Label3.Enabled = False
Text109.Enabled = False

If Check1.Value = 1 Then
    Label3.Enabled = True
    Text109.Enabled = True
End If

Validar

End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0
    
miSQL2 = "SELECT * FROM Bancos WHERE COD = " & Val(Text101.Text)
RgCYB.Open miSQL2, dbCYB, adOpenDynamic, adLockPessimistic

If Not RgCYB.EOF And ParaBcos Then
  vbMsgBoxTitle = "Error de Ingreso"
  vbMsgBoxText = "  " & vbCrLf & "Código  YA  Existe  "
  vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
   MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle
   RgCYB.Close
   Text101.SetFocus
   Exit Sub
    End If

If RgCYB.EOF Then
    Label4.Caption = "A L T A"
    Exit Sub
Else
    Label4.Caption = "MODIFICACIÓN"
    End If
    
Text102.Text = RgCYB!SUC
Text103.Text = RgCYB!NOMBRE

If RgCYB!Dir <> Null Then
    Text104.Text = RgCYB!NOMSUC
End If

If RgCYB!Dir <> Null Then
    Text105.Text = RgCYB!Dir
End If

If RgCYB!Loc <> Null Then
    Text106.Text = RgCYB!Loc
End If

If RgCYB!cp <> Null Then
    Text107.Text = RgCYB!cp
End If

' Text108.Text = RGCYB!codsuc
If RgCYB!clearing <> Null Then
    Text109.Text = RgCYB!clearing
End If

If RgCYB!propio <> Null Then
    Check1.Value = RgCYB!propio
End If

RgCYB.Close

Command4.Visible = True

End Sub

Private Sub Text102_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text109_KeyPress(KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub text102_lostfocus()

Validar

End Sub

Private Sub text103_lostfocus()

Validar

End Sub

Private Sub text104_lostfocus()

Validar

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

If Text102.Text = "" Then Text102.Text = "0"

If Text103.Text = "" Or Text104.Text = "" Or Text105.Text = "" Or Text106.Text = "" Or _
   Text107.Text = "" Then
    Exit Sub
End If

If Check1.Value = 1 And Val(Text109.Text) < 24 Then Exit Sub

Command2.Enabled = True
    
End Sub


