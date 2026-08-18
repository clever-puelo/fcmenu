VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form ModPrec 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Modificación de Precios por Porcentaje"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8505
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
   ScaleHeight     =   5445
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   585
      Left            =   6465
      ScaleHeight     =   525
      ScaleWidth      =   1440
      TabIndex        =   23
      Top             =   105
      Width           =   1500
      Begin VB.OptionButton Option1 
         Caption         =   "Todos"
         Height          =   255
         Index           =   1
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   270
         Width           =   1410
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Sección"
         Height          =   255
         Index           =   0
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   15
         Width           =   1410
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Anular"
      Height          =   345
      Left            =   4245
      TabIndex        =   6
      Top             =   4965
      Width           =   1185
   End
   Begin VB.Frame Frame3 
      Height          =   885
      Left            =   5190
      TabIndex        =   18
      Top             =   765
      Width           =   3075
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
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
         Left            =   1740
         MaxLength       =   8
         TabIndex        =   5
         Text            =   "0"
         ToolTipText     =   "Porcentaje en más o menos (colocar signo '-' adelante"
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Porcentaje :"
         Height          =   270
         Left            =   435
         TabIndex        =   19
         Top             =   405
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Height          =   885
      Left            =   330
      TabIndex        =   13
      Top             =   765
      Visible         =   0   'False
      Width           =   4740
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
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
         Left            =   2370
         MaxLength       =   8
         TabIndex        =   3
         Text            =   "99999"
         Top             =   525
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#9.99"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
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
         Left            =   4005
         MaxLength       =   5
         TabIndex        =   4
         Top             =   510
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
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
         Left            =   2370
         MaxLength       =   8
         TabIndex        =   1
         Text            =   "0"
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#9.99"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
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
         Left            =   4005
         MaxLength       =   5
         TabIndex        =   2
         Top             =   135
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "HASTA -->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   90
         TabIndex        =   21
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "DESDE -->"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   105
         TabIndex        =   20
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Pulgadas"
         Height          =   270
         Left            =   1335
         TabIndex        =   17
         Top             =   585
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Metros"
         Height          =   270
         Left            =   3075
         TabIndex        =   16
         Top             =   570
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Pulgadas"
         Height          =   270
         Left            =   1335
         TabIndex        =   15
         Top             =   210
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Metros"
         Height          =   270
         Left            =   3060
         TabIndex        =   14
         Top             =   195
         Visible         =   0   'False
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   315
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   8010
      Begin VSFlex8Ctl.VSFlexGrid FG1 
         Height          =   2970
         Left            =   45
         TabIndex        =   10
         Top             =   150
         Width           =   7890
         _cx             =   13917
         _cy             =   5239
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16710366
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16574407
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"ModPrec.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modificar"
      Height          =   345
      Left            =   5625
      TabIndex        =   7
      Top             =   4965
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   6960
      TabIndex        =   8
      Top             =   4965
      Width           =   1185
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   705
      ScaleHeight     =   465
      ScaleWidth      =   2715
      TabIndex        =   9
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   885
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Seleccione la Sección a Visualizar"
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Sección :"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Label Label7 
      Height          =   270
      Left            =   360
      TabIndex        =   22
      Top             =   5025
      Visible         =   0   'False
      Width           =   3720
   End
End
Attribute VB_Name = "ModPrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Posi(1 To 3) As Double, Valr(1 To 3) As Double, Mensaje As String
Dim LaRow As Integer
Private Const laselec1 = "SELECT * FROM Articulo WHERE COD1 = '"

Dim dbART As New ADODB.Connection, RgART As New ADODB.Recordset

Private Sub Form_Load()
Dim p1 As Integer, p2 As Integer, I1%

LaRow = 0
VerDetArt = False
  
Command1.Enabled = False

p1 = 1600
p2 = 3150

Me.Move p2, p1, 8600, 5800

dbART.ConnectionString = BDatos1
dbART.Open

Option1(0) = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbART.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 39 Then
    SendKeys "{TAB}"
End If

If KeyCode = 38 Or KeyCode = 37 Then
    SendKeys "+{TAB}"
End If

If KeyCode = 27 Then End

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

If KeyAscii = vbKeyEscape Then
    Command3_Click
    Combo2.SetFocus
    Command1.Enabled = False
    End If

End Sub

Private Sub Command1_Click()

AGrabar
Command3_Click

End Sub

Private Sub Command2_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Command3_Click()

If Not Combo2.Visible Then Exit Sub

    Frame2.Visible = False
    LimpiaGrilla
    Text2.Visible = False: Text4.Visible = False
    Text1.Visible = False: Text3.Visible = False
    Label2.Visible = False: Label4.Visible = False
    Label1.Visible = False: Label3.Visible = False
    Combo2.SetFocus
    Command1.Enabled = False

End Sub

Private Sub Combo2_Click()

Dim CAMPO1 As String, CAMPO2 As String, CAMPO3 As String
Dim UMED1 As String, UMED2 As String, UMED3 As String
Dim Codi As String

Command3_Click

Text2.Visible = False: Text3.Visible = False: Label4.Visible = False: Label5.Visible = False
Text4.Visible = False: Text1.Visible = False: Label1.Visible = False: Label5.Visible = False
Command1.Enabled = True

miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'SC   ' AND COD = '" & Combo2.Text & "'"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    
If RgTABL.EOF Then
    RgTABL.Close
    Exit Sub
End If

UMED1 = RgTABL!ALF1
UMED2 = RgTABL!ALF2
UMED3 = RgTABL!ALF3
    
RgTABL.Close

If UMED1 <> "     " Then
    miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED1 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    CAMPO1 = String(RgTABL!NumSD1 - 1, "#") & "0"
    If RgTABL!NumSD2 > 0 Then CAMPO1 = CAMPO1 & "," & String(RgTABL!NumSD2, "0")
    Text2.DataFormat.Format = CAMPO1
    Text2.MaxLength = Len(CAMPO1)
    Text4.DataFormat.Format = CAMPO1
    Text4.MaxLength = Len(CAMPO1)
    Label4.Caption = RgTABL!Descri
    Label4.Visible = True: Text2.Visible = True
    Label2.Caption = RgTABL!Descri
    Label2.Visible = True: Text4.Visible = True
    RgTABL.Close
    End If

If UMED2 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED2 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    CAMPO2 = String(RgTABL!NumSD1 - 1, "#") & "0"
    If RgTABL!NumSD2 > 0 Then CAMPO2 = CAMPO2 & "," & String(RgTABL!NumSD2, "0")
    Text3.DataFormat.Format = CAMPO2
    Text3.MaxLength = Len(CAMPO2)
    Text1.DataFormat.Format = CAMPO2
    Text1.MaxLength = Len(CAMPO2)
    Label5.Caption = RgTABL!Descri
    Label5.Visible = True: Text3.Visible = True
    Label1.Caption = RgTABL!Descri
    Label1.Visible = True: Text1.Visible = True
    If UMED1 = "     " Then Text3.SetFocus
    RgTABL.Close
    End If

DoVer

'If Text2.Visible Then Text2.SetFocus

End Sub

Private Sub Combo2_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 13 Or KeyCode = 9 Then
    KeyCode = 0
    Combo2_Click
End If

End Sub

Private Sub Option1_Click(Index As Integer)

If Index = 0 Then
    CargaSeccion
        Frame1.Visible = True
        Picture1.Visible = True
        
        Text2.Visible = False: Text4.Visible = False
        Label2.Visible = False: Label4.Visible = False
        
        Text1.Visible = False: Text3.Visible = False
        Label1.Visible = False: Label3.Visible = False
        
         DoVer
End If

If Index = 1 Then
        Frame1.Visible = False
        Picture1.Visible = False
        Frame3.Visible = True
        
        Text5.Visible = True
        Label6.Visible = True
        
        Text5.SetFocus
        
         DoVer
End If

End Sub

Private Sub text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub text1_Change()

Command1.Enabled = True
'text1.Text = text1.Text
End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc(",") Then Exit Sub
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_Change()

Command1.Enabled = True
'Text2.Text = Text2.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc(",") Then Exit Sub
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text3_GotFocus()

Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_Change()

Command1.Enabled = True
'Text3.Text = Text3.Text
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc(",") Then Exit Sub
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub text4_GotFocus()

Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub text4_Change()

Command1.Enabled = True
'text4.Text = text4.Text
End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc(",") Then Exit Sub
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'  Porcentaje
Private Sub text5_GotFocus()

Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)

End Sub

Private Sub text5_Change()

Command1.Enabled = True
'text5.Text = text5.Text
End Sub

Private Sub text5_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Or KeyAscii = Asc("-") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text5_LostFocus()
Dim DESDE As String * 5, HASTA As String * 5

DESDE = Combo2.Text
HASTA = Text4.Text

DoVer

End Sub

Private Sub AGrabar()
Dim I1%, Codi1 As String, Codi2 As String

Mensaje = "Desea Cargar los nuevos precios ?"
Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "Actualización de Precios")

If Respuesta = vbNo Then
    GoTo AlFinal
    End If
    
If Respuesta = vbCancel Then
         Exit Sub
         End If

dbART.BeginTrans

For I1 = 1 To FG1.Rows - 1
    
    Codi1 = FG1.TextMatrix(I1, 0)
    Codi2 = FG1.TextMatrix(I1, 1)
    miSQL2 = "SELECT  * FROM Articulo WHERE COD1 = '" & Codi1 & _
                         "' and COD2 = '" & Codi2 & "'"


RgART.Open miSQL2, dbART, adOpenDynamic, adLockPessimistic

    If RgART.EOF = True Then
        GoTo Salta
        End If
    
    RgART!prec = CCur(FG1.TextMatrix(I1, 3))

RgART!FACTUAL = Date
RgART!USUARIO = Mid$(Red_Usuario, 1, 6)

RgART.Update
RgART.Close

Salta:

Next I1

dbART.CommitTrans

AlFinal:

End Sub

Private Sub BuscaSeccion()
Dim Codi As String, ElResto As String
Dim UMED1 As String, UMED2 As String, UMED3 As String

Posi(1) = 0: Posi(2) = 0: Posi(3) = 0
Valr(1) = 0: Valr(2) = 0: Valr(3) = 0

ElResto = RESTO

    miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'SC   ' AND COD = '" & Correa & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    
    If RgTABL.EOF Then
        RgTABL.Close
        Exit Sub
    End If

UMED1 = RgTABL!ALF1
UMED2 = RgTABL!ALF2
UMED3 = RgTABL!ALF3
        
RgTABL.Close

If UMED1 <> "     " Then
    miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED1 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    Valr(1) = Mid(ElResto, 1, RgTABL!NumSD1)
    Posi(1) = RgTABL!NumSD3
    If RgTABL!NumSD2 > 0 Then
        Valr(1) = Valr(1) & "," & Mid(ElResto, RgTABL!NumSD1 + 1, RgTABL!NumSD2)
    End If
    Mid(ElResto, 1, RgTABL!NumSD1 + RgTABL!NumSD2) = Space(RgTABL!NumSD1 + RgTABL!NumSD2)
    ElResto = Trim(ElResto)
    RgTABL.Close
    End If

If UMED2 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED2 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    Valr(2) = Mid(ElResto, 1, RgTABL!NumSD1)
    Posi(2) = RgTABL!NumSD3
    If RgTABL!NumSD2 > 0 Then
        Valr(2) = Valr(2) & "," & Mid(ElResto, RgTABL!NumSD1 + 1, RgTABL!NumSD2)
    End If
    Mid(ElResto, 1, RgTABL!NumSD1 + RgTABL!NumSD2) = Space(RgTABL!NumSD1 + RgTABL!NumSD2)
    ElResto = Trim(ElResto)
    RgTABL.Close
    End If

If UMED3 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED3 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    Valr(3) = Mid(ElResto, 1, RgTABL!NumSD1)
    Posi(3) = RgTABL!NumSD3
    If RgTABL!NumSD2 > 0 Then
        Valr(3) = Valr(3) & "," & Mid(ElResto, RgTABL!NumSD1 + 1, RgTABL!NumSD2)
    End If
    RgTABL.Close
    End If

End Sub

Private Sub CargaSeccion()
Dim I1%

    ' Tabla de Secciones
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'SC   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

I1 = 0

If Not RgTABL.EOF Then
        Do Until RgTABL.EOF
        Combo2.AddItem RgTABL!COD
        Combo2.ItemData(I1) = I1
        RgTABL.MoveNext
        I1 = I1 + 1
        Loop
End If

RgTABL.Close

Combo2.ListIndex = 0
End Sub

Private Sub DoVer()

Frame2.Visible = False
Label7.Visible = False
    
If Option1(0) = True Then
    miSQL2 = laselec1 & Combo2.Text & "' ORDER BY COD2 DESC "
Else
    miSQL2 = "Select * FROM Articulo ORDER BY COD1, COD2 DESC "
End If

Show
DoEvents
 
RgART.Open miSQL2, dbART, adOpenForwardOnly, adLockReadOnly

If RgART.EOF Then
        vbMsgBoxTitle = "Error de Busqueda"
        vbMsgBoxText = " No existen datos para esta Selección "
        vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
        MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle
    RgART.Close
    Exit Sub
    End If

       LimpiaGrilla
       CargaGrilla
       
Frame2.Visible = True

'FG1.SetFocus

If FG1.Rows > 1 Then Command1.Default = False

End Sub

Sub LimpiaGrilla()
Dim I1%
   
LaRow = 0
FG1.Clear flexClearScrollable

For i = (FG1.Rows - 1) To 1 Step -1
FG1.RemoveItem (i)
Next i

End Sub

Sub CargaGrilla()

Dim StSuc01       As Double
Dim PCos          As Currency
Dim PVta          As Currency
Dim Stock        As Double
Dim Stmin        As Double
Dim Cod2   As String

Dim I1, i2
Dim i3 As Long

FG1.Col = 0


'Do Until RgART.EOF Or i3 = 100
Do Until RgART.EOF

PVta = 0: Stock = 0

Correa = RgART!COD1
'RESTO = rgart!cod2
Stock = RgART!Stock
Stmin = RgART!Stmin
PVta = RgART!prec

PCos = RgART!prec * ((CCur(Text5.Text) / 100) + 1)

If Val(RgART!Cod2) < Val(Text2.Text) Or Val(RgART!Cod2) > Val(Text4.Text) Then GoTo Salta

i3 = i3 + 1

'Cod2 = ""
Cod2 = Mid(RgART!Cod2, 1, InStr(1, RgART!Cod2, " ", vbTextCompare) - 1)

'If FG1.FindRow(Correa, , 0) = -1 Then
    FG1.AddItem Correa & vbTab & Cod2 & vbTab & Format(PVta, "$ ###,##0.00") & vbTab & Format(PCos, "$ ###,##0.00") & vbTab & _
            Stock, FG1.Row + 1
'    End If

Salta:

RgART.MoveNext
Loop

FG1.AutoSize 0, 5

FG1.Select 1, 0, FG1.Rows - 1, 1
FG1.Sort = flexSortGenericAscending

If FG1.Rows > 1 Then FG1.Row = 1

' Me.Show
' DoEvents

Label7.Caption = "Artículos seleccionados : " & i3 - 1
Label7.Visible = True

RgART.Close

End Sub


