VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form VerDesp 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Consulta de Despachos"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7365
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
   ScaleHeight     =   3630
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "  Despachos  "
      Height          =   2160
      Left            =   1170
      TabIndex        =   2
      Top             =   450
      Visible         =   0   'False
      Width           =   5325
      Begin VSFlex8Ctl.VSFlexGrid FG2 
         Height          =   1845
         Left            =   60
         TabIndex        =   3
         Top             =   255
         Width           =   5220
         _cx             =   9208
         _cy             =   3254
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
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
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"VerDesp.frx":0000
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
   Begin VB.TextBox Text2 
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
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2220
      MaxLength       =   20
      TabIndex        =   0
      Top             =   90
      Width           =   2745
   End
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   870
      TabIndex        =   4
      Top             =   525
      Width           =   6300
      Begin VSFlex8Ctl.VSFlexGrid FG1 
         Height          =   2385
         Left            =   60
         TabIndex        =   5
         Top             =   150
         Width           =   6180
         _cx             =   10901
         _cy             =   4207
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
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
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"VerDesp.frx":006A
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
      Caption         =   "Mostrar"
      Height          =   345
      Left            =   4530
      TabIndex        =   6
      Top             =   3135
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   5895
      TabIndex        =   7
      Top             =   3135
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Nº Despacho"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   840
      TabIndex        =   1
      Top             =   150
      Width           =   1320
   End
End
Attribute VB_Name = "VerDesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Posi(1 To 3) As Double, Valr(1 To 3) As Double
Dim LaRow As Integer

Dim dbART As New ADODB.Connection, RgART As New ADODB.Recordset
Dim dbSTOK As New ADODB.Connection, RgSTOK As New ADODB.Recordset

Private Sub FG2_Click()

Text2.Text = FG2.TextMatrix(FG2.Row, 0)
Frame2.Visible = False
FG1.SetFocus

End Sub

Private Sub Form_Load()
Dim p1 As Integer, p2 As Integer, HAY As String, I1%

HAY = ""
LaRow = 0
VerDetArt = False

p1 = 1600
p2 = 3150

Me.Move p2, p1, 7500, 4100

dbART.ConnectionString = BDatos1
dbART.Open
dbSTOK.ConnectionString = BDatos1
dbSTOK.Open

CargaDespa

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbART.Close
dbSTOK.Close

End Sub

Private Sub Form_GotFocus()

FG1.SetFocus

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

If KeyAscii = vbKeyEscape Then
    LimpiaGrilla
    Text2.Visible = False
    Command1.Default = True
    End If

End Sub

Private Sub Form_Activate()
 
 FG1.BackColorFixed = &HE0E0E0
 FG1.ForeColorFixed = &HC00000

End Sub

Private Sub Form_Deactivate()
 
 FG1.BackColorFixed = &H8000000F
 FG1.ForeColorFixed = &H80000012

End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)
    
If KeyAscii = vbKeyReturn Then SaleBien

End Sub

Private Sub FG1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo FinSub

If FG1.MouseRow = -1 Then Exit Sub

If FG1.MouseRow > 0 Then
   If FG1.MouseCol = 4 Then
        FG1.ToolTipText = " Canales : " & FG1.TextMatrix(FG1.MouseRow, 8) & " "
   Else
        FG1.ToolTipText = " " & FG1.TextMatrix(FG1.MouseRow, 7) & " "
            End If
    End If

If FG1.MouseRow = FG1.Row Then
   FG1.Cell(7, LaRow, 0, LaRow, 10) = &H80000012
   FG1.BackColorSel = &HFF0000
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 10) = &HFCE7C7
   LaRow = FG1.MouseRow
   Else
   FG1.BackColorSel = &H800000
   End If

If FG1.MouseRow > 0 And FG1.MouseRow < FG1.Rows Then
   FG1.Cell(7, LaRow, 0, LaRow, 10) = &H80000012
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 10) = &HFF0000
   LaRow = FG1.MouseRow
   End If

FinSub:

End Sub

Private Sub Command1_Click()

Text2.Text = ""
LimpiaGrilla
Frame2.Visible = True

End Sub

Private Sub Command2_Click()

'DetFact.Enabled = True
'DetFact.FG1.SetFocus

'Unload Me
If Me.width = 12900 Then
    Command1.Visible = True
    Command2.left = 5130
    Me.width = 6600
Else
    Call SlideWindow(Me, 50)
End If

End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_Change()

'Command1.Enabled = True: Command1.Default = True

DoVer

'Text2.Text = Text2.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    
If KeyAscii = vbKeyReturn And Text2.Text = "" Then
    CargaDespa
    Exit Sub
End If

If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc(",") Then Exit Sub
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub DoVer()

Frame2.Visible = False
        
miSQL1 = "SELECT * FROM Despachos WHERE NRODESP = '" & Text2.Text & "'"
RgSTOK.Open miSQL1, dbSTOK, adOpenForwardOnly, adLockReadOnly

    If RgSTOK.EOF = True Then
        RgSTOK.Close
        Exit Sub
    End If
    
Dim I1%
FG1.Clear flexClearScrollable

For i = (FG1.Rows - 1) To 1 Step -1
FG1.RemoveItem (i)
Next i

Do Until RgSTOK.EOF
        
miSQL2 = "SELECT * FROM Articulo WHERE COD1 = '" & RgSTOK!COD1 & "' and COD2 = '" & RgSTOK!Cod2 & "'"
RgART.Open miSQL2, dbART, adOpenForwardOnly, adLockReadOnly

If Not RgART.EOF Then
    FG1.AddItem RgART!COD1 & vbTab & RgART!Cod2 & vbTab & RgART!Prec & vbTab & RgSTOK!ENTRADA & _
                vbTab & RgSTOK!Stock, FG1.Row + 1
    End If

RgART.Close

RgSTOK.MoveNext

Loop

RgSTOK.Close

'FG1.SetFocus

' If FG1.Rows > 1 Then Command1.Default = False

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

i3 = i3 + 1

PVta = 0: Stock = 0
Correa = RgART!COD1

'RESTO = rgart!cod2
Stock = 0
Stmin = RgART!Stmin
PVta = RgART!Prec
PCos = RgART!PCos

'Cod2 = ""
Cod2 = Mid(RgART!Cod2, 1, InStr(1, RgART!Cod2, " ", vbTextCompare) - 1)

miSQL2 = "SELECT  * FROM Stock WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         Cod2 & "'"
RgSTOK.Open miSQL2, dbSTOK, adOpenForwardOnly, adLockReadOnly

    If Not RgSTOK.EOF Then
        Stock = RgSTOK!STUnid
    End If
    RgSTOK.Close

'If FG1.FindRow(Correa, , 0) = -1 Then
    FG1.AddItem Correa & vbTab & Cod2 & vbTab & PVta & vbTab & PCos & vbTab & _
            Stock & vbTab & Stmin & vbTab & RgART!Descri, FG1.Row + 1
'    End If

Salta:

RgART.MoveNext
Loop

FG1.AutoSize 0, 5

'FG1.Select 1, 2, 1, 2
FG1.Sort = flexSortGenericAscending

If FG1.Rows > 1 Then FG1.Row = 1

' Me.Show
' DoEvents

RgART.Close

End Sub

Private Sub SaleBien()
'DetFact.FG1.EditCell

    Correa = FG1.TextMatrix(FG1.Row, 0)
    RESTO = FG1.TextMatrix(FG1.Row, 1)
    
    BuscaSeccion

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

Private Sub CargaDespa()
Dim I1%
    
Frame2.Visible = True

    ' Busca un resumen de Despachos
miSQL2 = "SELECT NRODESP, FECENT, COUNT(NRODESP) as TotReg FROM Despachos" & _
         " group by NRODESP, FECENT"

RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    
FG2.Clear flexClearScrollable

For i = (FG2.Rows - 1) To 1 Step -1
FG2.RemoveItem (i)
Next i

Do Until RgTABL.EOF

FG2.AddItem RgTABL!NRODESP & vbTab & RgTABL!fecent & vbTab & RgTABL!TOTREG, FG2.Row + 1

RgTABL.MoveNext

Loop

RgTABL.Close
       
FG2.AutoSize 0, 2
FG2.Select 1, 1, 1, 1
FG2.Sort = flexSortGenericDescending

End Sub




