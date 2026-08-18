VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Cotizac 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga Cotización del Dólar"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
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
   ScaleHeight     =   4065
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   450
      Top             =   3840
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Height          =   330
      Left            =   1680
      TabIndex        =   6
      Top             =   3540
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   330
      Left            =   2940
      TabIndex        =   7
      Top             =   3540
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Height          =   1635
      Left            =   180
      TabIndex        =   10
      Top             =   1830
      Width           =   3915
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   1830
         MaxLength       =   12
         TabIndex        =   5
         Text            =   "0"
         Top             =   1170
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   1830
         MaxLength       =   12
         TabIndex        =   4
         Text            =   "0"
         Top             =   870
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1830
         MaxLength       =   12
         TabIndex        =   3
         Text            =   "0"
         Top             =   570
         Width           =   1425
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1830
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "0"
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dólar 'D' :"
         Height          =   240
         Index           =   4
         Left            =   720
         TabIndex        =   15
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dólar 'C' :"
         Height          =   240
         Index           =   3
         Left            =   720
         TabIndex        =   14
         Top             =   900
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dólar 'B' :"
         Height          =   240
         Index           =   2
         Left            =   720
         TabIndex        =   13
         Top             =   600
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dólar 'A' :"
         Height          =   240
         Index           =   1
         Left            =   720
         TabIndex        =   12
         Top             =   300
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   180
      TabIndex        =   9
      Top             =   870
      Width           =   3915
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1620
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "0"
         Top             =   420
         Width           =   1425
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dólar :"
         Height          =   240
         Index           =   0
         Left            =   510
         TabIndex        =   11
         Top             =   450
         Width           =   1110
      End
   End
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   330
      Left            =   1950
      TabIndex        =   0
      Top             =   180
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
      Format          =   159449091
      CurrentDate     =   36877
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   840
      TabIndex        =   16
      Top             =   600
      Width           =   2790
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha :"
      Height          =   240
      Left            =   -30
      TabIndex        =   8
      Top             =   225
      Width           =   1950
   End
End
Attribute VB_Name = "Cotizac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UnaVez As Boolean, Mensaje As String, EsFecha As String
'Dim Separa As String

Dim RgCOT As New ADODB.Recordset

Private Sub FECHA_CloseUp()

CargaFecha

Text1(0).SetFocus

End Sub

Private Sub FECHA_LostFocus()

CargaFecha

End Sub

Private Sub Form_Activate()

If UnaVez Then Exit Sub

UnaVez = True

CargaFecha

End Sub

Private Sub Form_Load()

Me.Move 0, 0, 4410, 4425

FECHA.Value = Date

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 39 Then SendKeys "{TAB}"
If KeyCode = 38 Or KeyCode = 37 Then SendKeys "+{TAB}"
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Command1_Click()

Grabacion

End Sub

Private Sub Command2_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Form_Unload(Cancel As Integer)

UnaVez = False

End Sub

Private Sub text1_GotFocus(Index As Integer)

Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)

End Sub

Private Sub text1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Text1_LostFocus(Index As Integer)

Text1(Index).Text = Format(Text1(Index).Text, "$ ##,##0.00 ")

End Sub

Private Sub Grabacion()
Dim I1%

Mensaje = "Desea Grabar ?"
Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "Cotización del Dólar")

If Respuesta = vbNo Then
    GoTo AlFinal
    End If
    
If Respuesta = vbCancel Then
         Exit Sub
         End If

dbTABL.BeginTrans

'EsFecha = Separa & FECHA.Value & Separa
'laLOGIK1 = Format(FECHA.Day, "00") & "/" & Format(FECHA.Month, "00") & "/" & FECHA.Year
laLOGIK1 = Format(FECHA.Year, "0000") & "/" & Format(FECHA.Month, "00") & "/" & Format(FECHA.Day, "00")

miSQL1 = "SELECT top 1 * FROM Cotizacion WHERE FECHA = " & Separa & laLOGIK1 & Separa
RgCOT.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic
    
'    RgCOT.Close

If RgCOT.EOF Then
    RgCOT.AddNew
    RgCOT!FECHA = FECHA.Value
    End If

    RgCOT!dolar = CCur(Text1(0).Text)
    RgCOT!dolara = CCur(Text1(1).Text)
    RgCOT!dolarb = CCur(Text1(2).Text)
    RgCOT!dolarc = CCur(Text1(3).Text)
    RgCOT!dolard = CCur(Text1(4).Text)
    RgCOT!dolarE = 0
    
    RgCOT!FECACT = Date
    RgCOT!USUARIO = Mid$(Red_Usuario, 1, 6)

    RgCOT.Update
    RgCOT.Close

'EsFecha = Separa & "01/01/1900" & Separa
miSQL1 = "SELECT * FROM Cotizacion WHERE fecha = " & Separa & "1900/01/01" & Separa
RgCOT.Open miSQL1, dbTABL, adOpenDynamic, adLockOptimistic

If RgCOT.EOF Then
    Exit Sub
End If
    
    LaCotiz = CCur(Text1(0).Text)

    RgCOT!dolar = CCur(Text1(0).Text)
    RgCOT!dolara = CCur(Text1(1).Text)
    RgCOT!dolarb = CCur(Text1(2).Text)
    RgCOT!dolarc = CCur(Text1(3).Text)
    RgCOT!dolard = CCur(Text1(4).Text)
    RgCOT!dolarE = 0
    
    RgCOT!FECACT = Date
    RgCOT!USUARIO = Mid$(Red_Usuario, 1, 6)

    RgCOT.Update
    RgCOT.Close

dbTABL.CommitTrans

AlFinal:

If FECHA.Value = Date Then
    FCMENU.StatusBar1.Panels.item(3).Text = "  1 Dólar = " & Format(Text1(0).Text, "$ ##,##0.00")
End If

Unload Me

'For i1 = 0 To 4
'Text1(i1).Text = 0
'Next i1

'FECHA.SetFocus
'CargaFecha

End Sub

Private Sub CargaFecha()

Label5.Caption = ""
EsFecha = Separa & FECHA.Value & Separa

laLOGIK1 = Format(FECHA.Year, "0000") & "/" & Format(FECHA.Month, "00") & "/" & Format(FECHA.Day, "00")
miSQL1 = "SELECT * FROM Cotizacion WHERE fecha = " & Separa & laLOGIK1 & Separa
RgCOT.Open miSQL1, dbTABL, adOpenForwardOnly, adLockReadOnly

If Not RgCOT.EOF Then
    GoTo CargaDatos
    End If
    
    RgCOT.Close

Label5.Caption = "N U E V O"

EsFecha = Separa & "1900/01/01" & Separa
miSQL1 = "SELECT * FROM Cotizacion WHERE fecha = " & Separa & "1900/01/01" & Separa
RgCOT.Open miSQL1, dbTABL, adOpenForwardOnly, adLockReadOnly

  If RgCOT.EOF Then
    MsgBox "No hay cotización Anterior, Avise", vbCritical + vbOKOnly, "Error Crítico"
    RgCOT.Close
    GrabaPrimero
    Unload Me
    Exit Sub
    End If
    
CargaDatos:
    
Text1(0).Text = Format(RgCOT!dolar, "$ ##,##0.00 ")
Text1(1).Text = Format(RgCOT!dolara, "$ ##,##0.00 ")
Text1(2).Text = Format(RgCOT!dolarb, "$ ##,##0.00 ")
Text1(3).Text = Format(RgCOT!dolarc, "$ ##,##0.00 ")
Text1(4).Text = Format(RgCOT!dolard, "$ ##,##0.00 ")
    
RgCOT.Close

End Sub

Private Sub GrabaPrimero()

dbTABL.BeginTrans

'EsFecha = Separa & FECHA.Value & Separa
laLOGIK1 = "1900/01/01"
miSQL1 = "SELECT top 1 * FROM Cotizacion WHERE FECHA = " & Separa & laLOGIK1 & Separa
RgCOT.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic
    
'    RgCOT.Close

If RgCOT.EOF Then
    RgCOT.AddNew
    RgCOT!FECHA = laLOGIK1
    End If

    RgCOT!dolar = 0
    RgCOT!dolara = 0
    RgCOT!dolarb = 0
    RgCOT!dolarc = 0
    RgCOT!dolard = 0
    RgCOT!dolarE = 0
    
    RgCOT!FECACT = Date
    RgCOT!USUARIO = Mid$(Red_Usuario, 1, 6)

    RgCOT.Update
    RgCOT.Close

dbTABL.CommitTrans

End Sub

Private Sub Timer1_Timer()

If Label5.Visible = True Then
    Label5.Visible = False
Else
    Label5.Visible = True
End If

End Sub
