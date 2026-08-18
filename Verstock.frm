VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form VerStock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ver Stock"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12795
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
   ScaleHeight     =   4245
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Movimientos de Stock"
      Height          =   3315
      Left            =   6585
      TabIndex        =   16
      Top             =   480
      Width           =   6030
      Begin VSFlex8Ctl.VSFlexGrid FG2 
         Height          =   3060
         Left            =   30
         TabIndex        =   17
         Top             =   210
         Width           =   5955
         _cx             =   10504
         _cy             =   5397
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
         BackColor       =   14608342
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   32768
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648384
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Verstock.frx":0000
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
   Begin VB.Frame Frame2 
      Height          =   2580
      Left            =   105
      TabIndex        =   12
      Top             =   1215
      Visible         =   0   'False
      Width           =   6300
      Begin VSFlex8Ctl.VSFlexGrid FG1 
         Height          =   2385
         Left            =   60
         TabIndex        =   13
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
         FormatString    =   $"Verstock.frx":00A6
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
      Left            =   105
      TabIndex        =   4
      Top             =   480
      Width           =   6285
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   5985
         TabIndex        =   5
         Top             =   150
         Width           =   6045
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
            TabIndex        =   8
            Top             =   60
            Visible         =   0   'False
            Width           =   735
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
            TabIndex        =   7
            Top             =   60
            Visible         =   0   'False
            Width           =   585
         End
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
            TabIndex        =   6
            ToolTipText     =   "Seleccione la Sección a Visualizar"
            Top             =   60
            Width           =   1335
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
            TabIndex        =   11
            Top             =   120
            Width           =   840
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
            TabIndex        =   10
            Top             =   120
            Visible         =   0   'False
            Width           =   1320
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
            TabIndex        =   9
            Top             =   120
            Visible         =   0   'False
            Width           =   1245
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mostrar"
      Height          =   345
      Left            =   3765
      TabIndex        =   3
      Top             =   3825
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   5130
      TabIndex        =   2
      Top             =   3825
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
      Left            =   600
      TabIndex        =   0
      Top             =   1320
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
         TabIndex        =   1
         Top             =   930
         Width           =   4755
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   105
      Top             =   480
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
   Begin VB.Label Label8 
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
      Left            =   6645
      TabIndex        =   19
      Top             =   3870
      Width           =   2985
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8100
      TabIndex        =   18
      Top             =   75
      Width           =   3225
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Consultas al Stock"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   1260
      TabIndex        =   15
      Top             =   60
      Width           =   4245
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
      Left            =   195
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   2985
   End
End
Attribute VB_Name = "VerStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Posi(1 To 3) As Double, Valr(1 To 3) As Double
Dim LaRow As Integer
Private Const laselec1 = "SELECT * FROM Articulo WHERE COD1 = '"

Dim dbART As New ADODB.Connection, RgART As New ADODB.Recordset
Dim dbSTOK As New ADODB.Connection, RgSTOK As New ADODB.Recordset

Private Sub Form_Load()
Dim p1 As Integer, p2 As Integer, HAY As String, I1%

HAY = ""
LaRow = 0
VerDetArt = False

p1 = 1600
p2 = 3150

Me.Move p2, p1, 6600, 4600

dbART.ConnectionString = BDatos1
dbART.Open
dbSTOK.ConnectionString = BDatos1
dbSTOK.Open

CargaSeccion

    Text2.Visible = False: Text3.Visible = False

miSQL2 = laselec1
    
If HAY = " AND " Then DoVer

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
    Text2.Visible = False: Text3.Visible = False
    Combo2.SetFocus
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

Private Sub FG1_Click()

Label6.Caption = FG1.TextMatrix(FG1.Row, 0) & FG1.TextMatrix(FG1.Row, 1)
Me.width = 12900
Command1.Visible = False
Command2.left = 11300
CargaMOVS

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
If Me.width = 12900 Then
    Command1.Visible = True
    Command2.left = 5130
    Me.width = 6600
Else
    Call SlideWindow(Me, 50)
End If

End Sub

Private Sub Combo2_GotFocus()

If Me.width = 12900 Then
    Command1.Visible = True
    Command2.left = 5130
    Me.width = 6600
End If

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

Show
DoEvents
 
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

FG1.SetFocus

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

Label2.Caption = "Renglones Encontrados : " & i3
Label2.Visible = True

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

Private Sub CargaMOVS()

Dim TipoMS       As String
Dim CpbteMS      As Double
Dim FechaMS      As Date
Dim CantMS       As Double
Dim Cod2         As String

Dim I1, i2
Dim i3 As Long

FG2.Col = 0

LimpiaMOVS
 
Correa = FG1.TextMatrix(FG1.Row, 0)
Cod2 = FG1.TextMatrix(FG1.Row, 1)

miSQL2 = "SELECT  * FROM Movstock WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         Cod2 & "'"
RgSTOK.Open miSQL2, dbSTOK, adOpenForwardOnly, adLockReadOnly

If RgSTOK.EOF Then
    Espera.Visible = True
    Espera.BackColor = &HC0&
    Espera.ForeColor = &HFFFF&
    Label1.ForeColor = &HFFFF&
    Espera.Caption = "   Error de Busqueda  "
    Label1.Caption = " No existen datos para este Artículo "
    RgSTOK.Close
    Exit Sub
    End If

Do Until RgSTOK.EOF

i3 = i3 + 1

CantMS = RgSTOK!Cant
CpbteMS = RgSTOK!Cpbte
FechaMS = RgSTOK!FECHA

TipoMS = ""
Select Case RgSTOK!TIPO
    Case "11": TipoMS = "N/C A "
    Case "12": TipoMS = "N/C B "
    Case "13": TipoMS = "Produc"
    Case "14": TipoMS = "Compra"
    Case "15": TipoMS = "Import"
    Case "18": TipoMS = "Aj.I.+"
    Case "21": TipoMS = "Fact A"
    Case "22": TipoMS = "Fact B"
    Case "23": TipoMS = "Remito"
    Case "24": TipoMS = "Promoc"
    Case "25": TipoMS = "Aj. - "
    Case "28": TipoMS = "Aj.I.-"
End Select

    FG2.AddItem TipoMS & vbTab & CpbteMS & vbTab & FechaMS & vbTab & CantMS & vbTab & FechaMS, FG2.Row + 1

Salta:

RgSTOK.MoveNext
Loop

FG2.AutoSize 0, 3

FG2.Col = 2
FG2.Sort = flexSortGenericAscending

If FG2.Rows > 1 Then FG2.Row = 1

' Me.Show
' DoEvents

Label8.Caption = "Renglones Encontrados : " & i3
Label8.Visible = True

RgSTOK.Close

End Sub

Sub LimpiaMOVS()
Dim I1%
   
LaRow = 0
FG2.Clear flexClearScrollable

For i = (FG2.Rows - 1) To 1 Step -1
FG2.RemoveItem (i)
Next i

End Sub

