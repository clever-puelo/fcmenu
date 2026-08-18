VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form DetRec 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Pendiente"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   3225
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2835
      Left            =   0
      TabIndex        =   1
      Top             =   -60
      Width           =   7530
      Begin VSFlex8Ctl.VSFlexGrid FG1 
         Height          =   2640
         Left            =   45
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   135
         Width           =   7425
         _cx             =   13097
         _cy             =   4657
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
         BackColor       =   16710366
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12582912
         ForeColorSel    =   16777215
         BackColorBkg    =   8421376
         BackColorAlternate=   16574407
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
         SelectionMode   =   3
         GridLines       =   3
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"DetRec.frx":0000
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
      Height          =   570
      Left            =   15
      TabIndex        =   2
      Top             =   2730
      Width           =   7515
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Left            =   1470
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6090
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pagos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4815
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Anticipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   195
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
   End
End
Attribute VB_Name = "DetRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LaRow As Integer, FinCod As Boolean, FinCant As Boolean, EsEnter As Boolean, VeNotaC As Boolean
Dim ImpAplic As Currency, ImpDebe As Currency, ImpDtos As Currency
Dim dbCCTE As New ADODB.Connection, RgCCTE As New ADODB.Recordset

'--------------------------------------------------------------
'           G R I L L A
'------------------------------------------------------------
'           0-Cpbte           1-Tipo (Alf.)  2-Fecha
'           3-Aplicado        4-Descto.      5-Debe
'           6-Fec.Vto.        7-Tipo (Num.)
'------------------------------------------------------------

Private Sub Form_Activate()

If NroRec = 0 Then
    CabRec.Show
End If

If FG1.Rows > 1 Then
    FG1.Row = 1
End If

End Sub

Private Sub Form_Load()

Me.Move 0, 1650, 7700, 3650

dbCCTE.ConnectionString = BDatos1
dbCCTE.Open

LaRow = 1

CabRec.Command2.Enabled = False

PieRec.Label304.Caption = 0: PieRec.Label303.Caption = 0: PieRec.Label302.Caption = 0
PieRec.Label301.Caption = 0: PieRec.Label305.Caption = 0: PieRec.Label306.Caption = 0
Text1.Text = 0

DoVer3

End Sub

Private Sub Form_Unload(Cancel As Integer)

'FCMENU.Toolbar2.Buttons(5).Value = tbrUnpressed
dbCCTE.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF9 Then
                    If FG1.Rows > 2 Then
                                FG1.TextMatrix(FG1.Row, 0) = ""
                               EmiRec.Show
                    End If
        End If

If KeyCode = vbKeyF8 Then
    Command1_Click
    Exit Sub
        End If

If KeyCode = vbKeyTab Then
    If Command2.Enabled Then
            DetPago.Show
            DetPago.SetFocus
            End If
    Exit Sub
        End If

If KeyCode = vbKeyF5 Then
    PagoParcial
    Exit Sub
        End If

If KeyCode = vbKeyF6 Then
    CargaDescto
    Exit Sub
        End If

End Sub

Private Sub Command1_Click()

'If Text1.Text <> "" Then
'    Text1.Text = ""
'    Text1.Visible = False
'Else
    Text1.Text = 0
    Text1.Visible = True
    Text1.SetFocus
'        End If

End Sub

Private Sub Command2_Click()

    
    DetPago.Show
    DetPago.FG1.SetFocus
    
End Sub

Private Sub Command3_Click()

Unload Me
Unload CabRec
Unload DetPago
Unload PieRec

End Sub

'--------------------------------------------------------------
'           G R I L L A
'--------------------------------------------------------------

'------------------------------------------------------------
'           0-Cpbte           1-Tipo (Alf.)  2-Fecha
'           3-Aplicado        4-Descto.      5-Debe
'           6-Fec.Vto.        7-Tipo (Num.)
'------------------------------------------------------------

Private Sub FG1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
        If FG1.Row > 0 Then
            If FG1.TextMatrix(FG1.Row, 3) > 0 Then
                FG1.TextMatrix(FG1.Row, 3) = 0
            Else
                FG1.TextMatrix(FG1.Row, 3) = FG1.TextMatrix(FG1.Row, 5)
            End If
        End If
End If

CalcAplic

End Sub

Private Sub FG1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

If KeyAscii = 13 Then
    EsEnter = True
End If

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If
' If Col = 0 Or Col = 2 Then LaTecla = KeyAscii

End Sub


'******************************************************

Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)

If FG1.TextMatrix(Row, Col) = "" Then Exit Sub

If Not EsEnter Then Exit Sub

If Not IsNumeric(FG1.TextMatrix(Row, 3)) Then
    FG1.TextMatrix(FG1.Row, 3) = FG1.TextMatrix(FG1.Row, 5)
    GoTo Abajo
End If

Select Case Col
    Case 3
        FG1.Col = 4
        GoTo Abajo
    Case 4
        If CCur(FG1.TextMatrix(FG1.Row, 4)) > CCur(FG1.TextMatrix(FG1.Row, 3)) Then
           FG1.TextMatrix(FG1.Row, 4) = 0
           Exit Sub
        End If
        FG1.Col = 3
        If FG1.Row < (FG1.Rows - 1) Then
            FG1.Row = FG1.Row + 1
        Else
            FG1.Row = 1
        End If
    Case Else
        Exit Sub
End Select

Abajo:

If EsEnter Then
    EsEnter = False
    
       FG1.EditCell
       If FG1.EditWindow = 0 Then FG1.EditCell

End If

CalcAplic

End Sub
'******************************************************

Private Sub FG1_LostFocus()

CalcAplic

End Sub


Private Sub FG1_Click()

If FG1.Row > 0 Then
    If FG1.TextMatrix(FG1.Row, 3) > 0 Then
        FG1.TextMatrix(FG1.Row, 3) = 0
        FG1.TextMatrix(FG1.Row, 4) = 0
    Else
        FG1.TextMatrix(FG1.Row, 3) = FG1.TextMatrix(FG1.Row, 5)
    End If
End If

End Sub

Private Sub FG1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo SALIR

If FG1.MouseRow = -1 Then Exit Sub
If LaRow > (FG1.Rows - 1) Then LaRow = (FG1.Rows - 1)

If FG1.MouseCol = 3 Then
        FG1.ToolTipText = " [F5] Importe Parcial // [Enter] ó Click Importe Total"
Else
        FG1.ToolTipText = " " & FG1.TextMatrix(FG1.MouseRow, 1) & " "
    End If

If FG1.MouseRow = FG1.Row Then
   FG1.Cell(7, LaRow, 0, LaRow, 5) = &H80000012
   FG1.BackColorSel = vbBlue
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 5) = &HFF0000
   LaRow = FG1.MouseRow
   Else
   FG1.BackColorSel = vbBlue
   End If

If FG1.MouseRow > 0 And FG1.MouseRow < FG1.Rows Then
   FG1.Cell(7, LaRow, 0, LaRow, 5) = &H80000012
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 5) = &HFF0000
   LaRow = FG1.MouseRow
   End If

SALIR:

End Sub

'----------------------------------------------------------------------------

Private Sub PagoParcial()

If FG1.Row = 0 Then Exit Sub

FG1.Col = 3
FG1.EditCell

End Sub

Private Sub CargaDescto()

If FG1.Row = 0 Or FG1.TextMatrix(FG1.Row, 3) = 0 Then Exit Sub

FG1.Col = 4
FG1.EditCell

End Sub

Sub DoVer3()

Dim FECCPB As Date
Dim FECVTO As Date
Dim Cpbte As Double
Dim TIPO As String * 10
Dim RESTO As String
Dim DEBE As Currency
Dim HABER As String
Dim Saldo As Double
Dim PLIS As Double
Dim CANFAC As Single

  If ORDENADOx <> "" Then
         Clasifica = elORDEN & ORDENADOx & ORDENADOad
  Else
         Clasifica = ""
End If
'///           ----------------------------------------
 
cpbSAL = 0: totSAL = 0: cpbENT = 0: totENT = 0: Promed = 0: Prom2mes = 0
         
miSQL = "SELECT FECHA, TIPO, CPBTE, IMPUT1, IMPUT2, IMPUT3, IMPUT4, IMPUT5, DEBE, IMPTE, fecvto " _
        & " FROM CtasCtes WHERE TIPO9 <> '1' AND TIPO <> 0 AND TIPO <> 4 AND " & _
        " TIPO <> 6 AND TIPO <> 2 AND TIPO <> 8 AND TIPO <> 9 AND CLTE = " & CodCLTE & " AND DEBE > 0 ORDER BY FECHA ASC"

RgCCTE.Open miSQL, dbCCTE, adOpenForwardOnly, adLockReadOnly

    If RgCCTE.EOF Then
        RgCCTE.Close
        Exit Sub
        End If

Do While Not RgCCTE.EOF

FECCPB = RgCCTE!FECHA
FECVTO = RgCCTE!FECVTO
Cpbte = RgCCTE!Cpbte

Select Case RgCCTE!TIPO
    Case 1
        TIPO = "Fact."
    Case 2
        TIPO = "N/C."
    Case 3
        TIPO = "N/D."
    Case 5
        TIPO = "P/C."
    Case 7
        TIPO = "Ctdo."
    Case 8
        TIPO = "NCC"
End Select


RESTO = RgCCTE!DEBE

If RgCCTE!TIPO = 2 Or RgCCTE!TIPO = 5 Or RgCCTE!TIPO = 8 Then
                Saldo = Saldo - RgCCTE!DEBE
Else
                Saldo = Saldo + RgCCTE!DEBE
                End If
                
FG1.AddItem Cpbte & vbTab & TIPO & vbTab & FECCPB & vbTab & 0 & vbTab & 0 & vbTab & RESTO & vbTab _
                    & FECVTO & vbTab & RgCCTE!TIPO, FG1.Rows


RgCCTE.MoveNext

Loop

If FG1.Rows > 1 Then

    For i = 1 To FG1.Rows - 1
        
        DEBE = DEBE + FG1.TextMatrix(i, 5)
        CANFAC = CANFAC + 1

    Next i


End If

PieRec.Label301.Caption = Format(DEBE, "$###,##0.00")

End Sub


Private Sub CalcAplic()
Dim A1 As Integer

ImpAplic = 0: ImpDtos = 0

For A1 = 1 To FG1.Rows - 1

    ImpAplic = ImpAplic + FG1.TextMatrix(A1, 3)
    ImpDtos = ImpDtos + FG1.TextMatrix(A1, 4)
    
    If FG1.TextMatrix(A1, 3) > 0 Then FG1.IsSelected(A1) = True
    
Next A1

ImpAplic = ImpAplic + CCur(Text1.Text)
PieRec.Label306.Caption = Format(ImpAplic, "$###,##0.00")
PieRec.Label305.Caption = Format(ImpDtos, "$###,##0.00")
PieRec.Label302.Caption = Format((ImpAplic - ImpDtos), "$###,##0.00")

End Sub

Private Sub Text1_Change()

If Not IsNumeric(Text1.Text) Then Text1.Text = 0

PieRec.Label302.Caption = Format((ImpAplic - ImpDtos) + CCur(Text1.Text), "$ ###,##0.00")

End Sub

Private Sub text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyBack Then Exit Sub

If KeyAscii = vbKeyReturn Then FG1.SetFocus

If KeyAscii = 46 Then KeyAscii = 44

If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text1_LostFocus()

If Not IsNumeric(Text1.Text) Then
    Text1.Text = 0
    End If
    
Text1.Text = Format(Val(Replace(Text1.Text, ",", ".")), "###,##0.00")

End Sub

