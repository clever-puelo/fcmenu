VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ConstAFIP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga Constancias de la AFIP"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5520
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
   ScaleHeight     =   3825
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2580
      Left            =   120
      TabIndex        =   6
      Top             =   690
      Width           =   5280
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Left            =   1935
         MaxLength       =   12
         TabIndex        =   0
         Text            =   "112222222233"
         Top             =   810
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   270
         Left            =   1935
         MaxLength       =   12
         TabIndex        =   3
         Top             =   2160
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "ConstAFIP.frx":0000
         Left            =   1935
         List            =   "ConstAFIP.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   1950
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   1935
         TabIndex        =   2
         Top             =   1440
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
         CalendarBackColor=   16777152
         CustomFormat    =   "ddd d MMM yyy"
         Format          =   39649283
         CurrentDate     =   36877
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   330
         Left            =   1935
         TabIndex        =   7
         Top             =   1800
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777152
         CustomFormat    =   "ddd d MMM yyy"
         Format          =   39649283
         CurrentDate     =   36877
      End
      Begin VB.Label Label1 
         Caption         =   "C.U.I.T.         :"
         Height          =   240
         Left            =   540
         TabIndex        =   13
         Top             =   855
         Width           =   1410
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Solic. :"
         Height          =   240
         Left            =   540
         TabIndex        =   12
         Top             =   1530
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Vto.    :"
         Height          =   240
         Left            =   540
         TabIndex        =   11
         Top             =   1845
         Width           =   1365
      End
      Begin VB.Label Label4 
         Caption         =   "Verificad.     :"
         Height          =   240
         Left            =   540
         TabIndex        =   10
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo              :"
         Height          =   240
         Left            =   540
         TabIndex        =   9
         Top             =   1170
         Width           =   1410
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1935
         TabIndex        =   8
         Top             =   270
         Visible         =   0   'False
         Width           =   2940
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   330
      Left            =   2415
      TabIndex        =   4
      Top             =   3345
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   330
      Left            =   3900
      TabIndex        =   5
      Top             =   3345
      Width           =   1410
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Constancias de AFIP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   795
      TabIndex        =   14
      Top             =   150
      Width           =   4245
   End
End
Attribute VB_Name = "ConstAFIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Me.Move 2500, 0, 5600, 4150

LimpPant

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 39 Then SendKeys "{TAB}"
If KeyCode = 38 Or KeyCode = 37 Then SendKeys "+{TAB}"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Command1_Click()
    
'    Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Command2_Click()
    
    Grabar
    
End Sub


Private Sub Combo1_Click()

If DTPicker1.Visible Then
    DTPicker1.SetFocus
End If

End Sub


Private Sub DTPicker1_Click()

DTPicker2.Value = DateAdd("d", 30, DTPicker1.Value)
Text2.SetFocus

End Sub

Private Sub DTPicker1_Change()

DTPicker2.Value = DateAdd("d", 30, DTPicker1.Value)
Text2.SetFocus

End Sub
Private Sub text1_Change()

If Len(Text1.Text) = 11 Then
    If Cuit(Text1.Text) Then
        If BuscaCUIT(Text1.Text) = 1 Then
            Combo1.SetFocus
        Else
            DTPicker1.SetFocus
        End If
    Else
        Text1.SetFocus
    End If
End If

End Sub

Private Sub text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub

If KeyAscii = vbKeyReturn Then
End If

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    Command2.SetFocus
End If

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Function BuscaCUIT(NroCUIT As String) As String

BuscaCUIT = "3"
Label7.Caption = "  Vigente  "
Label7.Visible = True

DTPicker1.Value = Date
Text2.Text = ""
Combo1.ListIndex = 0

miSQL2 = "SELECT * FROM Constancias WHERE CUIT = " & NroCUIT
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
        
    If RgTABL.EOF Then
        BuscaCUIT = "1"
        Label7.Caption = " No Está Cargada "
        RgTABL.Close
        Exit Function
    End If

    If RgTABL!FECVTO < Date Then
        BuscaCUIT = "2"
        Label7.Caption = "  Vencida  "
    End If

DTPicker1.Value = RgTABL!FECSOL
Text2.Text = RgTABL!VERIF
Combo1.ListIndex = RgTABL!TIPO - 1
        
        RgTABL.Close

DTPicker1.SetFocus

End Function

Private Sub Grabar()

miSQL2 = "SELECT * FROM Constancias WHERE CUIT = " & Text1.Text
    RgTABL.Open miSQL2, dbTABL, adOpenDynamic, adLockPessimistic
        
If RgTABL.EOF Then
    RgTABL.AddNew
'    RgTABL!Cuit = Text1.Text
End If
    
    RgTABL!Cuit = Text1.Text
    RgTABL!TIPO = Combo1.ListIndex + 1
    RgTABL!Categ = ""
    RgTABL!FECSOL = DTPicker1.Value
    RgTABL!FECVTO = DTPicker2.Value
    RgTABL!VERIF = Text2.Text
    
RgTABL!FACTUAL = Date
RgTABL!USUARIO = Mid(Red_Usuario, 1, 6)

RgTABL.Update
RgTABL.Close

LimpPant

End Sub

Private Sub LimpPant()

Text1.Text = ""
Text2.Text = ""
DTPicker1.Value = Date
DTPicker2.Value = DateAdd("d", 30, Date)

Combo1.ListIndex = 0

Label7.Visible = False

If Text1.Visible Then
    Text1.SetFocus
End If

End Sub

