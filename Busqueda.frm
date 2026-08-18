VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Busqueda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Busqueda de Artículos   "
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8010
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Alta"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   3330
      Width           =   1185
   End
   Begin VB.Frame Espera 
      BackColor       =   &H00FCE7C7&
      Caption         =   "   Carga de Datos   "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2325
      Left            =   1320
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   5235
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   8
         Top             =   930
         Width           =   4755
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   6705
      TabIndex        =   3
      Top             =   3330
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
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
      Left            =   5325
      TabIndex        =   2
      Top             =   3315
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   705
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7995
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   1125
         ScaleHeight     =   435
         ScaleWidth      =   5985
         TabIndex        =   6
         Top             =   150
         Width           =   6045
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   900
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Seleccione la Sección a Visualizar"
            Top             =   60
            Width           =   1335
         End
         Begin VB.TextBox Text3 
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
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5370
            MaxLength       =   5
            TabIndex        =   1
            Top             =   60
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
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
            Left            =   3330
            MaxLength       =   8
            TabIndex        =   0
            Top             =   60
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Metros"
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
            Left            =   4110
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Pulgadas"
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
            Left            =   2010
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label3 
            Caption         =   "Sección :"
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
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   840
         End
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2580
      Left            =   0
      TabIndex        =   4
      Top             =   690
      Visible         =   0   'False
      Width           =   8010
      Begin VSFlex8Ctl.VSFlexGrid FG1 
         Height          =   2385
         Left            =   60
         TabIndex        =   15
         Top             =   150
         Width           =   7890
         _cx             =   13917
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
         FormatString    =   $"Busqueda.frx":0000
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   "DSN=STOCK"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STOCK"
      OtherAttributes =   ""
      UserName        =   "clever"
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Stock"
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
   Begin VB.Label Label2 
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
      Left            =   90
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   2985
   End
End
Attribute VB_Name = "Busqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Posi(1 To 3) As Double, Valr(1 To 3) As Double
Dim LaRow As Integer
Private Const laselec1 = "SELECT * FROM Articulo WHERE COD1 = '"

Dim dbART As New ADODB.Connection, RgART As New ADODB.Recordset

Private Sub Command3_Click()

TipoMov2 = 1
VerDetArt = False
ABMArt.Show

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
    Text2.Visible = False: Text3.Visible = False
    Combo2.SetFocus
    Command1.Default = True
    End If

End Sub

Private Sub Form_Load()
Dim p1 As Integer, p2 As Integer, HAY As String, I1%

HAY = ""
LaRow = 0
VerDetArt = False

p1 = 1600
p2 = 3150

Me.Move p2, p1, 8100, 4100

dbART.ConnectionString = BDatos1
dbART.Open

CargaSeccion

    Text2.Visible = False: Text3.Visible = False

miSQL2 = laselec1
    
If HAY = " AND " Then DoVer

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbART.Close

End Sub

Private Sub Form_Activate()
 
 FG1.BackColorFixed = &HE0E0E0
 FG1.ForeColorFixed = &HC00000

End Sub

Private Sub Form_Deactivate()
 
 FG1.BackColorFixed = &H8000000F
 FG1.ForeColorFixed = &H80000012

End Sub

Private Sub FG1_Click()

SaleBien

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

    miSQL2 = laselec1 & Combo2.Text & "'"

If Len(Text2.Text) > 0 Then
    miSQL2 = miSQL2 & " AND COD2 = '" & Text2.Text
End If

If Len(Text3.Text) > 0 Then
    miSQL2 = miSQL2 & Text3.Text & "'"
End If

If Len(Text2.Text) > 0 And Len(Text3.Text) = 0 Then
    miSQL2 = miSQL2 & "'"
End If

DoVer

End Sub

Private Sub Command2_Click()

'DetFact.Enabled = True
'DetFact.FG1.SetFocus

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Combo2_Click()
Dim CAMPO1 As String, CAMPO2 As String, CAMPO3 As String
Dim UMED1 As String, UMED2 As String, UMED3 As String
Dim Codi As String

Text2.Visible = False: Text3.Visible = False: Label4.Visible = False: Label5.Visible = False
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
    Label4.Caption = RgTABL!Descri
    Label4.Visible = True: Text2.Visible = True
    RgTABL.Close
    End If

If UMED2 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED2 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    CAMPO2 = String(RgTABL!NumSD1 - 1, "#") & "0"
    If RgTABL!NumSD2 > 0 Then CAMPO2 = CAMPO2 & "," & String(RgTABL!NumSD2, "0")
    Text3.DataFormat.Format = CAMPO2
    Text3.MaxLength = Len(CAMPO2)
    Label5.Caption = RgTABL!Descri
    Label5.Visible = True: Text3.Visible = True
    If UMED1 = "     " Then Text3.SetFocus
    RgTABL.Close
    End If

'DoVer

End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_Change()

Command1.Enabled = True: Command1.Default = True
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

Command1.Enabled = True: Command1.Default = True
'Text3.Text = Text3.Text
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub
If KeyAscii = Asc(",") Then Exit Sub
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub DoVer()

Frame2.Visible = False
Label2.Visible = False
    
    Espera.BackColor = &HFCE7C7
    Espera.ForeColor = &HC00000
    Label1.ForeColor = &HC00000
    Espera.Caption = "   Carga de Datos  "
    Label1 = "Buscando Datos, Por Favor Espere ..."
Espera.Visible = True

miSQL2 = miSQL2 & " ORDER BY COD2 DESC "

'Show
'DoEvents
 
RgART.Open miSQL2, dbART, adOpenForwardOnly, adLockReadOnly

If RgART.EOF Then
    Espera.Visible = True
    Espera.BackColor = &HC0&
    Espera.ForeColor = &HFFFF&
    Label1.ForeColor = &HFFFF&
    Espera.Caption = "   Error de Busqueda  "
    Label1.Caption = " No existen datos para esta Selección "
    RgART.Close
    Exit Sub
    End If

       LimpiaGrilla
       CargaGrilla
       
Espera.Visible = False
Frame2.Visible = True

If FG1.Visible Then FG1.SetFocus

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

i3 = i3 + 1

PVta = 0: Stock = 0

Correa = RgART!COD1
'RESTO = rgart!cod2
Stock = RgART!Stock
Stmin = RgART!Stmin
PVta = RgART!prec
PCos = RgART!PCos

'Cod2 = ""
Cod2 = Mid(RgART!Cod2, 1, InStr(1, RgART!Cod2, " ", vbTextCompare) - 1)

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

Label2.Caption = "Renglones Encontrados : " & i2 - 1
Label2.Visible = True

RgART.Close

End Sub

Private Sub SaleBien()
'DetFact.FG1.EditCell

    Correa = FG1.TextMatrix(FG1.Row, 0)
    RESTO = FG1.TextMatrix(FG1.Row, 1)
    
    BuscaSeccion

'If EstaFactu Then
Select Case VolverArt
Case 1
    Precio = FG1.TextMatrix(FG1.Row, 2)
    DetFact.FG1.TextMatrix(DetFact.FG1.Row, 0) = Trim(Correa)
    DetFact.FG1.TextMatrix(DetFact.FG1.Row, 5) = Trim(FG1.TextMatrix(FG1.Row, 6))
        DesArt = Trim(FG1.TextMatrix(FG1.Row, 6))
    DetFact.FG1.TextMatrix(DetFact.FG1.Row, 7) = FG1.TextMatrix(FG1.Row, 2)
    DetFact.FG1.TextMatrix(DetFact.FG1.Row, Posi(1)) = Valr(1)
        MatrItems(DetFact.FG1.Row, 1, 2) = Valr(1)
    If Posi(2) > 0 Then
        DetFact.FG1.TextMatrix(DetFact.FG1.Row, Posi(2)) = Valr(2)
        MatrItems(DetFact.FG1.Row, 2, 2) = Valr(2)
    End If
    
    If Posi(3) > 0 Then
        DetFact.FG1.TextMatrix(DetFact.FG1.Row, Posi(3)) = Valr(3)
        MatrItems(DetFact.FG1.Row, 3, 2) = Valr(3)
    End If
'   DetFact.FG1.TextMatrix(DetFact.FG1.Row, 1) = RESTO
'    DetFact.FG1.Col = 4
    DetFact.QuePosic
    
    Unload Me
    
Case 2
    Stock.FG1.TextMatrix(Stock.FG1.Row, 0) = Trim(Correa)
    Stock.FG1.TextMatrix(Stock.FG1.Row, 2) = Trim(FG1.TextMatrix(FG1.Row, 4))
        DesArt = ""
    Stock.FG1.TextMatrix(Stock.FG1.Row, 1) = Valr(1)
    
    Unload Me

Case Else
    TipoMov2 = 3
    VerDetArt = True
    ABMArt.Show
End Select

End Sub

Sub BuscaDesdeAfuera()
Dim I1

If Combo2.ListCount = 0 Then
    CargaSeccion
End If


For I1 = 0 To Combo2.ListCount - 1

Combo2.ListIndex = I1

If Trim(Combo2.Text) = Trim(Correa) Then
   Exit For
End If

Next


miSQL2 = laselec1 & Combo2.Text & "'"

If Len(Text2.Text) > 0 Then
    miSQL2 = miSQL2 & " AND COD2 = '" & Text2.Text & "'"
End If

DoVer

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
