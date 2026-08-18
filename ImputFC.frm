VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form ImputFC 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Selección de la Factura a Imputar"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Height          =   330
      Left            =   4770
      TabIndex        =   1
      Top             =   2160
      Width           =   1185
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   2100
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _cx             =   10610
      _cy             =   3704
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14352126
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16576
      ForeColorSel    =   12640511
      BackColorBkg    =   8438015
      BackColorAlternate=   12640511
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ImputFC.frx":0000
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
Attribute VB_Name = "ImputFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset
Dim RgCCTE As New ADODB.Recordset

Private Sub Command1_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub FG1_Click()
Dim Letra As String, ElCpbte As Double, ElTipo As Integer, LaFECHA As Date

If FG1.Row < 1 Then Exit Sub

'MsgBox "apfecha: " & FG1.TextMatrix(FG1.Row, 3)
ApFecha = FG1.TextMatrix(FG1.Row, 3)
'MsgBox "ApLetra: " & FG1.TextMatrix(FG1.Row, 4)
ApLetra = FG1.TextMatrix(FG1.Row, 4)
'MsgBox "ApTipo: " & FG1.TextMatrix(FG1.Row, 5)
ApTipo = FG1.TextMatrix(FG1.Row, 5)
'MsgBox "ApPtoVta: " & FG1.TextMatrix(FG1.Row, 7)
ApPtoVta = FG1.TextMatrix(FG1.Row, 7)
'MsgBox "ApCpbte: " & FG1.TextMatrix(FG1.Row, 6)
ApCpbte = FG1.TextMatrix(FG1.Row, 6)
'MsgBox "ApImpte: " & FG1.TextMatrix(FG1.Row, 2)
ApImpte = FG1.TextMatrix(FG1.Row, 2)

EmiFact.Show

Unload Me

End Sub

Private Sub Form_Load()

Me.Move 2400, 3100, 6100, 2900

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

CargaDeuda


End Sub

Private Sub Form_Unload(Cancel As Integer)
 
 dbCLTE.Close

End Sub

Sub CargaDeuda()

Dim Impte         As String
Dim Cpbte         As String
Dim ElCpbte       As String * 14
Dim FECHA         As Date
Dim ElImpte       As String * 12
Dim I1, i2, Tc
Dim i3 As Long
   
FG1.Clear flexClearScrollable

For I1 = (FG1.Rows - 1) To 1 Step -1
FG1.RemoveItem (I1)
Next I1
     
miSQL = "SELECT * FROM CtasCtes WHERE CLTE = " & CodCLTE & _
        " and ( TIPO = 1 OR TIPO = 3 ) AND DEBE > 0 ORDER BY FECHA"
RgCCTE.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgCCTE.EOF Then
    RgCCTE.Close
    Exit Sub
    End If


Do Until RgCCTE.EOF

i3 = i3 + 1

FECHA = RgCCTE!FECHA

Select Case RgCCTE!TIPO
    Case 1: Cpbte = "Fact."
    Case 3: Cpbte = "N/D. "
End Select

Cpbte = Cpbte & RgCCTE!Letra
Cpbte = Cpbte & " " & RgCCTE!Cpbte
LSet ElCpbte = Cpbte
RSet ElImpte = Format(RgCCTE!DEBE, "$ ###,##0.00")

FG1.AddItem FECHA & vbTab & ElCpbte & vbTab & ElImpte & vbTab & FECHA & vbTab & _
            RgCCTE!Letra & vbTab & RgCCTE!TIPO & vbTab & RgCCTE!Cpbte & _
            vbTab & RgCCTE!prefijo, FG1.Rows

Me.Show
DoEvents

Salta:

RgCCTE.MoveNext

Loop

    RgCCTE.Close

FG1.AutoSize 0, 6

If FG1.Rows > 1 Then FG1.Row = 1


' FG1.SetFocus

End Sub

