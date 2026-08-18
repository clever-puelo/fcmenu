VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form DetFact 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Detalle"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   3285
   ClientWidth     =   8385
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4395
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   8295
      Begin VSFlex8Ctl.VSFlexGrid FG1 
         Height          =   4200
         Left            =   45
         TabIndex        =   0
         Top             =   135
         Width           =   8205
         _cx             =   14473
         _cy             =   7408
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
         ForeColorSel    =   16711136
         BackColorBkg    =   12632256
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
         SelectionMode   =   1
         GridLines       =   3
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"DetFact.frx":0000
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
         Begin VB.Frame Frame2 
            BackColor       =   &H0080FF80&
            Caption         =   " Despachos "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   1695
            Left            =   2160
            TabIndex        =   2
            Top             =   1230
            Visible         =   0   'False
            Width           =   5910
            Begin VSFlex8Ctl.VSFlexGrid FG2 
               Height          =   1410
               Left            =   15
               TabIndex        =   3
               Top             =   195
               Width           =   5850
               _cx             =   10319
               _cy             =   2487
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
               ForeColorSel    =   16711136
               BackColorBkg    =   12648447
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
               SelectionMode   =   1
               GridLines       =   3
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"DetFact.frx":01BA
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
      End
   End
End
Attribute VB_Name = "DetFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LaRow As Integer, FinCod As Boolean, FinCant As Boolean, EsEnter As Boolean
Dim LaTecla As Byte, SeccAnt As String, SeMueve As Boolean, SaleMal As Boolean
Dim CposVan1 As Single, CposVan2 As Single, CanCpos1 As Single, CanCpos2 As Single
Dim SumDtos As Currency
Dim EsAlgo(1 To 3) As Double
Dim Edita(1 To 7) As String
Dim Lgo(1 To 7) As Single
Dim Deci(1 To 7) As Single
Dim Posi(1 To 7) As Single

Dim dbART As New ADODB.Connection, RgART As New ADODB.Recordset

Private Sub Form_Load()
  
Me.Move 0, 1610, 8450, 4900

dbART.ConnectionString = BDatos1
dbART.Open
    
Correa = "": RESTO = ""
'EstaFactu = True
VolverArt = 1

LaRow = 1

 '   FG1.Col = 0
 '   FG1.EditCell

End Sub

Private Sub Form_Unload(Cancel As Integer)

'EstaFactu = False
VolverArt = 0
dbART.Close

End Sub

Private Sub Form_Activate()
 
 FG1.BackColorFixed = &HE0E0E0
 FG1.ForeColorFixed = &HC00000

If VeNotaC Or CodCLTE = 0 Then GoTo AOtraCosa

VeNotaC = True

miSQL2 = "SELECT * FROM Notaclte WHERE CLTE = " & CodCLTE
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
   
If Not RgTABL.EOF Then
    RgTABL.Close
    NOTACLTE.Show
    Exit Sub
End If

RgTABL.Close
         
AOtraCosa:
         
If FG1.Rows <= 2 Then Exit Sub
If FG1.Row < 1 Then FG1.Row = 1

End Sub

Private Sub Form_Deactivate()
 
 FG1.BackColorFixed = &H8000000F
 FG1.ForeColorFixed = &H80000012

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

' si oprime alguna flecha marca para el "AfterRowColChange"
If KeyCode > 36 And KeyCode < 41 Then
    SeMueve = True
End If

If KeyCode = vbKeyF2 Then
                    If Len(FG1.EditText) < 8 Then
                        Correa = Space(8 - Len(FG1.EditText)) & FG1.EditText
                        Else
                        Correa = FG1.EditText
                    End If
    Busqueda.Show
        End If

If KeyCode = vbKeyF9 Then
                    If FG1.Rows > 2 Then
                                FG1.TextMatrix(FG1.Row, 0) = ""
                               EmiFact.Show
                    End If
        End If

End Sub

'--------------------------------------------------------------
'           G R I L L A
'--------------------------------------------------------------

'------------------------------------------------------------
'           0-Rubro           1-Subrubro         2-Nro.
'           3-Descripción     4-Cantidad         5-Precio
'           6-Importe         7-Proveedor
'------------------------------------------------------------

Private Sub FG1_GotFocus()
            
If ClteNomb = "" Then Exit Sub
'FG1.Row = LaRow
' FG1.Col = 1
    FG1.EditCell
    If FG1.EditWindow = 0 Then FG1.EditCell

End Sub

Private Sub FG1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
            
If NewCol = 0 Then FinCod = False: FinCant = False

If SeMueve = False Then Exit Sub
            
    FG1.Col = 0

    FG1.Select FG1.Row, FG1.Col
    FG1.EditCell
    If FG1.EditWindow = 0 Then FG1.EditCell

SeMueve = False

End Sub

Private Sub FG1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim ShiftTest As Integer

If Col = 8 Then Exit Sub

LaTecla = KeyCode

'   ALT-DEL Borra el Renglón (si no es el único)
ShiftTest = Shift And 7
If (KeyCode = 46 And (ShiftTest = 4 Or ShiftTest = 6)) And FG1.Rows > 2 Then
    BorraRenglon (Row)
'    VerSiAgrega
    FG1.Select (FG1.Rows - 1), 0
    FG1.EditCell
    If FG1.EditWindow = 0 Then FG1.EditCell
    Exit Sub
End If

End Sub

Private Sub FG1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

If Col = 0 And KeyAscii = vbKeyF1 Then
    Busqueda.BuscaDesdeAfuera
End If

If KeyAscii = 13 Then
    EsEnter = True
End If

If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 9 Then Exit Sub

If Col = 1 Or Col = 2 Or Col = 3 Or Col = 7 Or (Col = 6 And FG1.TextMatrix(Row, 0) = "**") Then
            If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
            If KeyAscii = Asc(",") Then Exit Sub
End If
                
If Col <> 0 And Col <> 5 Then
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If
End If

End Sub

Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i2%, Ii1%, Ii2%
' FinCod : Indica si ya se terminó de cargar el código de artículo (por ej. "PL")
' CanCpos1 : Cantidad de Entradas del código además de la Sección
' CposVan1 : Indica cuantos campos del cód. ya se cargaron
'
' FinCant  : Indica si ya se terminó de cargar las cantidades
' CanCpos2 : Cantidad de Entradas de la unidad de medida de venta (ej. Unidades)
' CposVan2 : Indica cuantos campos de la unidad de medida ya se cargaron
'

'------------------------------------------------------------------

If Not EsEnter Then Exit Sub

If FG1.TextMatrix(Row, Col) = "" Then GoTo Abajo

If Col = 0 Then
        Correa = FG1.TextMatrix(FG1.Row, 0)
        FinCod = False: FinCant = False
        BuscaSeccion
        SeccAnt = Correa
            If Posi(1) = 0 Then
                CposVan2 = 1
                FG1.Col = Posi(4)
                FinCod = True
            Else
                CposVan1 = 1
                FG1.Col = Posi(1)
                RESTO = ""
            End If
        GoTo Abajo
End If

'If FG1.TextMatrix(Row, Col) = 0 Then GoTo Abajo
                    
If FinCant Then GoTo AlPrecio
If FinCod Then GoTo AlaCant

'------------------------------------------------------------------

' Si NO ES "**"
If FG1.TextMatrix(FG1.Row, 0) <> "**" Then
    If FG1.TextMatrix(FG1.Row, Posi(CposVan1)) = "" Then
        EsAlgo(CposVan1) = 0
    Else
        EsAlgo(CposVan1) = CCur(FG1.TextMatrix(FG1.Row, Posi(CposVan1)))
    End If
End If
    

MatrItems(FG1.Row, CposVan1, 2) = EsAlgo(CposVan1)
                    
'  Acá se arma el codigo. Si  es compuesto lo va armando en RESTO
If CposVan1 <= CanCpos1 Then
    If Deci(CposVan1) > 0 Then
        ' Para el caso que tenga decimales: debe juntar los campos y sin comas
        RESTO = RESTO & SacaComas(EsAlgo(CposVan1), Edita(CposVan1), Deci(CposVan1))
    Else
        ' Si el código es simple, lo mueve directamente
        RESTO = RESTO & EsAlgo(CposVan1)
    End If
End If
    
If CposVan1 >= CanCpos1 Then
    If FG1.TextMatrix(FG1.Row, 0) <> "**" Then
        VerSiBusca
        If SaleMal Then
            CposVan1 = 1
            SaleMal = False
            GoTo Abajo
        End If
    End If
    
    FinCod = True
    If Posi(4) > 0 Then
        FG1.Col = Posi(4)
        CposVan2 = 1
    Else
        FG1.Col = Posi(7)
        CposVan2 = 4
    End If
    GoTo Abajo
End If
                    
If Posi(1) > 0 Then
    If EsAlgo(1) > 0 Then
        MatrItems(FG1.Row, 1, 2) = EsAlgo(1)
        FG1.Col = Posi(2)
        CposVan1 = CposVan1 + 1
        GoTo Abajo
    End If
End If
                    
If Posi(2) > 0 Then
    If EsAlgo(2) > 0 Then
        MatrItems(FG1.Row, 2, 2) = EsAlgo(2)
        FG1.Col = Posi(3)
        CposVan1 = CposVan1 + 1
        GoTo Abajo
    End If
End If
                    
If Posi(3) > 0 Then
    If CCur(FG1.TextMatrix(FG1.Row, Posi(3))) > 0 Then
        MatrItems(FG1.Row, 3, 2) = CCur(FG1.TextMatrix(FG1.Row, Posi(3)))
        FG1.Col = Posi(3)
        CposVan1 = CposVan1 + 1
        GoTo Abajo
    End If
End If

'---------------------------------------------------
AlaCant:
        
' If Posi(CposVan2 + 3) = 0 Then Exit Sub

'MatrItems(FG1.Row, CposVan2 + 3, 2) = Val(FG1.TextMatrix(FG1.Row, Posi(CposVan2 + 3)))
MatrItems(FG1.Row, CposVan2 + 3, 2) = CCur(FG1.TextMatrix(FG1.Row, Posi(CposVan2 + 3)))

If Posi(4) > 0 Then
    If Val(FG1.TextMatrix(FG1.Row, Posi(4))) = 0 Then
        FG1.Col = Posi(4)
        CposVan2 = CposVan2 + 1
        GoTo Abajo
    End If
End If
                    
If Posi(5) > 0 Then
    If Val(FG1.TextMatrix(FG1.Row, Posi(5))) = 0 Then
        FG1.Col = Posi(5)
        CposVan2 = CposVan2 + 1
        GoTo Abajo
    End If
End If
                    
If Posi(6) > 0 Then
    If Val(FG1.TextMatrix(FG1.Row, Posi(6))) = 0 Then
        FG1.Col = Posi(6)
        CposVan2 = CposVan2 + 1
        GoTo Abajo
    End If
End If
                    
If Posi(7) > 0 Then
    If Val(FG1.TextMatrix(FG1.Row, Posi(7))) = 0 Then
        FG1.Col = Posi(7)
        CposVan2 = 4
        GoTo Abajo
    End If
End If
    
If CposVan2 >= CanCpos2 Then
    FinCant = True
    FG1.Col = 7
    CposVan2 = 0
    GoTo Abajo
End If

'---------------------------------------------------
AlPrecio:
    
FinCant = False
FG1.TextMatrix(FG1.Row, 8) = CCur(FG1.TextMatrix(FG1.Row, 7))

If Posi(4) > 0 Then
    FG1.TextMatrix(FG1.Row, 8) = CCur(FG1.TextMatrix(FG1.Row, Posi(4))) * CCur(FG1.TextMatrix(FG1.Row, 7))
End If

If Posi(5) > 0 Then
    FG1.TextMatrix(FG1.Row, 8) = FG1.TextMatrix(FG1.Row, 8) * CCur(FG1.TextMatrix(FG1.Row, Posi(5)))
End If
    
If Posi(6) > 0 Then
    FG1.TextMatrix(FG1.Row, 8) = FG1.TextMatrix(FG1.Row, 8) * CCur(FG1.TextMatrix(FG1.Row, Posi(6)))
End If
    
If Posi(7) > 0 Then
    FG1.TextMatrix(FG1.Row, 8) = FG1.TextMatrix(FG1.Row, 8) * CCur(FG1.TextMatrix(FG1.Row, Posi(7)))
End If

' Obtiene el Precio Unitario según la UNIDAD DE FACTURACIÓN
FG1.TextMatrix(FG1.Row, 7) = FG1.TextMatrix(FG1.Row, 8) / FG1.TextMatrix(FG1.Row, Posi(7))

DxClte CodCLTE, FG1.TextMatrix(FG1.Row, 0)

FG1.TextMatrix(FG1.Row, 9) = "": SumDtos = 0
If Dtos(1) > 0 Then FG1.TextMatrix(FG1.Row, 9) = Dtos(1)
If Dtos(2) > 0 Then FG1.TextMatrix(FG1.Row, 9) = FG1.TextMatrix(FG1.Row, 9) & "+" & Dtos(2)
If Dtos(3) > 0 Then FG1.TextMatrix(FG1.Row, 9) = FG1.TextMatrix(FG1.Row, 9) & "+" & Dtos(3)
If Dtos(4) > 0 Then FG1.TextMatrix(FG1.Row, 9) = FG1.TextMatrix(FG1.Row, 9) & "+" & Dtos(4)
If Dtos(5) > 0 Then FG1.TextMatrix(FG1.Row, 9) = FG1.TextMatrix(FG1.Row, 9) & "+" & Dtos(5)
SumDtos = 0
For i2 = 1 To 5
    If Dtos(i2) > 0 Then
        SumDtos = SumDtos + ((FG1.TextMatrix(FG1.Row, 8) - SumDtos) * (Dtos(i2) / 100))
    End If
Next i2

FG1.TextMatrix(FG1.Row, 10) = Format(SumDtos, "##,##0.00")
    
    VerSiAgrega

Abajo:

If EsEnter Then
    EsEnter = False
    
    If Col < 8 Then
       FG1.EditCell
       If FG1.EditWindow = 0 Then FG1.EditCell
    Else
        VerSiAgrega
    End If

End If

End Sub


Private Function SacaComas(Dato As Double, Edita As String, Deci As Single) As String
Dim i2%, Ii1%, Ii2%

        laLOGIK2 = Dato
        Ii1 = InStr(1, Edita, ".")               '   Busca posición de punto decimal
        Ii2 = InStr(1, laLOGIK2, ",")             '   Idem
        If Ii1 > 0 And Ii2 = 0 Then
            laLOGIK2 = laLOGIK2 & "," & String(Deci, "0")  ' Agrega un 0 a la derecha si corresponde
        End If

        SacaComas = Replace(laLOGIK2, ",", "")                      '  Quita la coma
        SacaComas = String(Len(Edita) - (Len(laLOGIK2)), "0") & SacaComas  '  Agrega un cero a la izquierda p/completar

End Function

Private Sub FG1_Click()
            
If FG1.Col = 8 Then FG1.Col = 0

    FG1.Select FG1.Row, FG1.Col
    FG1.EditCell

End Sub

Private Sub FG1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Sale

If FG1.MouseRow = -1 Then Exit Sub
If LaRow > (FG1.Rows - 1) Then LaRow = (FG1.Rows - 1)

If FG1.MouseRow > 0 Then
        FG1.ToolTipText = " " & FG1.TextMatrix(FG1.MouseRow, 5) & " "
    End If

If FG1.MouseRow = FG1.Row Then
   FG1.Cell(7, LaRow, 0, LaRow, 5) = &H80000012
   FG1.BackColorSel = &HFF0000
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 5) = &HFCE7C7
   LaRow = FG1.MouseRow
   Else
   FG1.BackColorSel = &H800000
   End If

If FG1.MouseRow > 0 And FG1.MouseRow < FG1.Rows Then
   FG1.Cell(7, LaRow, 0, LaRow, 5) = &H80000012
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 5) = &HFF0000
   LaRow = FG1.MouseRow
   End If

Sale:

End Sub
'------------------------------------------------------
'                Fin Grilla
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub FG2_Click()

FG1.TextMatrix(FG1.Row, 13) = FG2.TextMatrix(FG2.Row, 0)
Frame2.Visible = False

End Sub

Sub VerSiAgrega()
Dim c1%

FCMENU.StatusBar1.Panels.item(2).Text = ""

Importe = 0: SumDtos = 0
For c1 = 1 To FG1.Rows - 1
      If FG1.TextMatrix(c1, 8) <> "" Then
          Importe = Importe + FG1.TextMatrix(c1, 8)
          SumDtos = SumDtos + FG1.TextMatrix(c1, 10)
          End If
Next

PieFact.Label306.Caption = Format(SumDtos, "$ ###,##0.00")
PieFact.Label301.Caption = Format(Importe, "$ ###,##0.00")

'  Vuelve a Esconder la Descripción
If FG1.TextMatrix(FG1.Row, 0) = "**" Then

    MatrItems(FG1.Row, 1, 1) = "": MatrItems(FG1.Row, 2, 1) = ""
    MatrItems(FG1.Row, 3, 1) = "": MatrItems(FG1.Row, 4, 1) = ""
    MatrItems(FG1.Row, 5, 1) = "": MatrItems(FG1.Row, 6, 1) = ""
    MatrItems(FG1.Row, 7, 1) = ""
    
    FG1.ColHidden(0) = False
    FG1.ColHidden(1) = False
    FG1.ColHidden(2) = False
    FG1.ColHidden(3) = False
    FG1.ColHidden(4) = False
    FG1.ColHidden(5) = True
    FG1.TextMatrix(FG1.Row, 11) = False
    FG1.TextMatrix(FG1.Row, 12) = False
    FG1.TextMatrix(FG1.Row, 13) = ""
Else
    If FG1.TextMatrix(FG1.Row, 11) = False Then
       FG1.TextMatrix(FG1.Row, 5) = DesArt
    Else
       FG1.TextMatrix(FG1.Row, 5) = DesSecc
    End If
End If
    
FG1.TextMatrix(FG1.Row, 14) = RESTO


If FG1.TextMatrix((FG1.Rows - 1), 8) <> "" Then
        FG1.Rows = FG1.Rows + 1
        If FG1.Rows < 26 Then
                    FG1.Row = FG1.Rows - 1
        Else
            MsgBox "Ha ingresado 25 reglones. Llegó al límite", vbExclamation + vbOKOnly, "Fin de la Carga"
            EmiFact.Show
        End If
Else
    FG1.Row = FG1.Rows - 1
                    End If
                    
paraSTATUS = "Items : " & (FG1.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS
                    
    FG1.TextMatrix(FG1.Row, 0) = SeccAnt
    CposVan1 = 1
    FG1.Col = Posi(1)
    FinCod = False
    FG1.TextMatrix(FG1.Row, 5) = FG1.TextMatrix(FG1.Row - 1, 5)
    FG1.TextMatrix(FG1.Row, 11) = FG1.TextMatrix(FG1.Row - 1, 11)
    FG1.TextMatrix(FG1.Row, 12) = FG1.TextMatrix(FG1.Row - 1, 12)
    FG1.TextMatrix(FG1.Row, 13) = ""
     MatrItems(FG1.Row, 1, 1) = MatrItems(FG1.Row - 1, 1, 1)
     MatrItems(FG1.Row, 2, 1) = MatrItems(FG1.Row - 1, 2, 1)
     MatrItems(FG1.Row, 3, 1) = MatrItems(FG1.Row - 1, 3, 1)
     MatrItems(FG1.Row, 4, 1) = MatrItems(FG1.Row - 1, 4, 1)
     MatrItems(FG1.Row, 5, 1) = MatrItems(FG1.Row - 1, 5, 1)
     MatrItems(FG1.Row, 6, 1) = MatrItems(FG1.Row - 1, 6, 1)
     
     MatrItems(FG1.Row, 1, 2) = MatrItems(FG1.Row - 1, 1, 2)
     MatrItems(FG1.Row, 2, 2) = MatrItems(FG1.Row - 1, 2, 2)
     MatrItems(FG1.Row, 3, 2) = MatrItems(FG1.Row - 1, 3, 2)
     MatrItems(FG1.Row, 4, 2) = MatrItems(FG1.Row - 1, 4, 2)
     MatrItems(FG1.Row, 5, 2) = MatrItems(FG1.Row - 1, 5, 2)
     MatrItems(FG1.Row, 6, 2) = MatrItems(FG1.Row - 1, 6, 2)

   
    FG1.EditCell

RESTO = ""

End Sub

Sub VerSiBusca()
Dim PCos          As Currency
Dim PVta          As Currency
Dim Stock         As Single
Dim Stmin         As Single
Dim Seccio     As String

Seccio = FG1.TextMatrix(FG1.Row, 1)
        
If FG1.FindRow(Seccio, , 1) <> -1 And FG1.FindRow(Seccio, , 1) <> FG1.Row Then
    If FG1.TextMatrix(FG1.FindRow(Seccio, , 1), 0) = FG1.TextMatrix(FG1.Row, 0) Then
        MsgBox "Artículo ya Ingresado", vbExclamation + vbOKOnly, "Duplicación"
        SaleMal = True
        Exit Sub
    End If
End If

        
miSQL2 = "SELECT  * FROM Articulo WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         RESTO & "'"
RgART.Open miSQL2, dbART, adOpenForwardOnly, adLockReadOnly

    If RgART.EOF = True Then
        RgART.Close
        Busqueda.BuscaDesdeAfuera
        Exit Sub
    End If
    
'FG1.TextMatrix(FG1.Row, 5) = RgART!Descri
 
' Calcula el precio con o sin dolares, segun mareca en tabla de secciones
If EnDolares Then                                   '   si es factura en Dolares
    If FG1.TextMatrix(FG1.Row, 12) Then             '   si tiene mca de dolares la sección
            FG1.TextMatrix(FG1.Row, 7) = RgART!Prec
     Else
            FG1.TextMatrix(FG1.Row, 7) = RgART!Prec / LaCotiz
    End If
Else                                                   ' si es factura normal
    If FG1.TextMatrix(FG1.Row, 12) Then             '   si tiene mca de dolares la sección
            FG1.TextMatrix(FG1.Row, 7) = RgART!Prec * LaCotiz
     Else
            FG1.TextMatrix(FG1.Row, 7) = RgART!Prec
    End If
End If

DesArt = Trim(RgART!Descri)

Correa = RgART!COD1
RESTO = RgART!Cod2

QuePosic

'        PVta = RgART!PREC
'        PCos = RgART!PCos
'        STOCK = RgART!STOCK
'        Stmin = RgART!Stmin

' FG1.TextMatrix(FG1.Row, 6) = PVta
        
RgART.Close

'FCMENU.StatusBar1.Panels.item(2).Text = "Val.Pulg.: " & Format(PVta, "$ ###,##0.00")
        
'---------------------------------------
'  Despachos
'--------------------------------------------
        Frame2.Visible = False
        
miSQL2 = "SELECT  * FROM Despachos WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         RESTO & "' and STOCK > 0"
RgART.Open miSQL2, dbART, adOpenForwardOnly, adLockReadOnly

    If RgART.EOF = True Then
        RgART.Close
        Exit Sub
    End If
    
Dim I1%
FG2.Clear flexClearScrollable

For i = (FG2.Rows - 1) To 1 Step -1
FG2.RemoveItem (i)
Next i

Do Until RgART.EOF

FG2.AddItem RgART!NRODESP & vbTab & RgART!Cpbte & vbTab & RgART!fecent & vbTab & RgART!Stock, FG2.Row + 1

RgART.MoveNext

Loop

RgART.Close
       
FG2.AutoSize 0, 2
    
        Frame2.top = FG1.RowPos(FG1.Row)
        Frame2.Visible = True
        FG2.SetFocus
        
End Sub

Sub QuePosic()
        
If Posi(4) > 0 Then
    FG1.TextMatrix(FG1.Row, Posi(4)) = 0
    End If
If Posi(5) > 0 Then
    FG1.TextMatrix(FG1.Row, Posi(5)) = 0
    End If
If Posi(6) > 0 Then
    FG1.TextMatrix(FG1.Row, Posi(6)) = 0
    End If
If Posi(7) > 0 Then
    FG1.TextMatrix(FG1.Row, Posi(7)) = 0
    End If

End Sub

Private Sub BuscaSeccion()
Dim UMED1 As String, UMED2 As String, UMED3 As String, UMED4 As String
Dim UMED5 As String, UMED6 As String, UMED7 As String
Dim Codi As String

CposVan1 = 0: CanCpos1 = 0
Posi(1) = 0: Posi(2) = 0: Posi(3) = 0: Posi(4) = 0: Posi(5) = 0: CposVan2 = 0: CanCpos2 = 0
Posi(6) = 0:  Posi(7) = 0

'  Para cargar una descripción / Artículo que no está en DB
If Correa = "**" Then
    FG1.ColHidden(0) = True
    FG1.ColHidden(1) = True
    FG1.ColHidden(2) = True
    FG1.ColHidden(3) = True
    FG1.ColHidden(4) = True
    FG1.ColHidden(5) = False
    
    CanCpos1 = 1
    CanCpos2 = 1
    Edita(1) = "####0"
    Posi(1) = 5
    Posi(7) = 6
'    FinCod = True
    
    Exit Sub

End If


    miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'SC   ' AND COD = '" & Correa & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    
    If RgTABL.EOF Then
        RgTABL.Close
        MsgBox "La Sección " & FG1.TextMatrix(FG1.Row, 0) & " No Existe en la Tabla", vbCritical + vbOKOnly, "Error de Ingreso"
            FG1.Col = 0
            FG1.EditCell
            If FG1.EditWindow = 0 Then FG1.EditCell
        Exit Sub
    End If
    
    If RgTABL!ALF7 = "     " Or RgTABL!ALF7 = "" Then
        RgTABL.Close
        MsgBox "               La Sección  '" & FG1.TextMatrix(FG1.Row, 0) & "'" & _
               vbCrLf & " No Tiene cargada la UNIDAD DE FACTURACIÓN", _
                vbCritical + vbOKOnly, "Error en Tabla de Secciones"
                FG1.Col = 0
            FG1.EditCell
            If FG1.EditWindow = 0 Then FG1.EditCell
        Exit Sub
    End If

DesSecc = Trim(RgTABL!Descri)
FG1.TextMatrix(FG1.Row, 1) = ""
FG1.TextMatrix(FG1.Row, 2) = ""
FG1.TextMatrix(FG1.Row, 3) = ""
FG1.TextMatrix(FG1.Row, 4) = ""
FG1.TextMatrix(FG1.Row, 6) = ""
FG1.TextMatrix(FG1.Row, 7) = ""
FG1.TextMatrix(FG1.Row, 8) = ""
FG1.TextMatrix(FG1.Row, 9) = ""
FG1.TextMatrix(FG1.Row, 10) = ""
FG1.TextMatrix(FG1.Row, 11) = RgTABL!mca1
FG1.TextMatrix(FG1.Row, 12) = RgTABL!mca2
FG1.TextMatrix(FG1.Row, 13) = ""

UMED1 = RgTABL!ALF1
UMED2 = RgTABL!ALF2
UMED3 = RgTABL!ALF3
UMED4 = RgTABL!ALF4
UMED5 = RgTABL!ALF5
UMED6 = RgTABL!ALF6
UMED7 = RgTABL!ALF7
        
RgTABL.Close

MatrItems(FG1.Row, 1, 1) = "": MatrItems(FG1.Row, 2, 1) = ""
MatrItems(FG1.Row, 3, 1) = "": MatrItems(FG1.Row, 4, 1) = ""
MatrItems(FG1.Row, 5, 1) = "": MatrItems(FG1.Row, 6, 1) = ""
MatrItems(FG1.Row, 7, 1) = ""


MatrItems(FG1.Row, 1, 2) = "": MatrItems(FG1.Row, 2, 2) = ""
MatrItems(FG1.Row, 3, 2) = "": MatrItems(FG1.Row, 4, 2) = ""
MatrItems(FG1.Row, 5, 2) = "": MatrItems(FG1.Row, 6, 2) = ""
MatrItems(FG1.Row, 7, 2) = ""

If UMED1 <> "     " Then
    miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED1 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    MatrItems(FG1.Row, 1, 1) = RgTABL!Descri
    Lgo(1) = RgTABL!NumSD1 + RgTABL!NumSD2
    Deci(1) = RgTABL!NumSD2
    Edita(1) = String(RgTABL!NumSD1 - 1, "0") & "0"
    If RgTABL!NumSD2 > 0 Then Edita(1) = Edita(1) & "." & String(RgTABL!NumSD2, "0")
    'Text2.DataFormat.Format = CAMPO1
    'Text2.MaxLength = Len(CAMPO1)
    Posi(1) = RgTABL!NumSD3
    CanCpos1 = 1
    RgTABL.Close
    End If

If UMED2 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED2 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    MatrItems(FG1.Row, 2, 1) = RgTABL!Descri
    Lgo(2) = RgTABL!NumSD1 + RgTABL!NumSD2
    Deci(2) = RgTABL!NumSD2
    Edita(2) = String(RgTABL!NumSD1 - 1, "0") & "0"
    If RgTABL!NumSD2 > 0 Then Edita(2) = Edita(2) & "." & String(RgTABL!NumSD2, "0")
    'Text3.DataFormat.Format = CAMPO2
    'Text3.MaxLength = Len(CAMPO2)
    Posi(2) = RgTABL!NumSD3
    CanCpos1 = 2
    RgTABL.Close
    End If

If UMED3 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED3 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    MatrItems(FG1.Row, 3, 1) = RgTABL!Descri
    Lgo(3) = RgTABL!NumSD1 + RgTABL!NumSD2
    Deci(3) = RgTABL!NumSD2
    Edita(3) = String(RgTABL!NumSD1 - 1, "0") & "0"
    If RgTABL!NumSD2 > 0 Then Edita(3) = Edita(3) & "." & String(RgTABL!NumSD2, "0")
    'Text3.DataFormat.Format = CAMPO2
    'Text3.MaxLength = Len(CAMPO2)
    Posi(3) = RgTABL!NumSD3
    CanCpos1 = 3
    RgTABL.Close
    End If

If UMED4 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED4 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    MatrItems(FG1.Row, 4, 1) = RgTABL!Descri
    Lgo(4) = RgTABL!NumSD1 + RgTABL!NumSD2
    Edita(4) = String(RgTABL!NumSD1 - 1, "0") & "0"
    If RgTABL!NumSD2 > 0 Then Edita(4) = Edita(4) & "," & String(RgTABL!NumSD2, "0")
    'Text3.DataFormat.Format = CAMPO2
    'Text3.MaxLength = Len(CAMPO2)
    Posi(4) = RgTABL!NumSD3
    CanCpos2 = 1
    RgTABL.Close
    End If

If UMED5 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED5 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    MatrItems(FG1.Row, 5, 1) = RgTABL!Descri
    Lgo(5) = RgTABL!NumSD1 + RgTABL!NumSD2
    Edita(5) = String(RgTABL!NumSD1 - 1, "0") & "0"
    If RgTABL!NumSD2 > 0 Then Edita(5) = Edita(5) & "," & String(RgTABL!NumSD2, "0")
    'Text3.DataFormat.Format = CAMPO2
    'Text3.MaxLength = Len(CAMPO2)
    Posi(5) = RgTABL!NumSD3
    CanCpos2 = CanCpos2 + 1
    RgTABL.Close
    End If

If UMED6 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED6 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    MatrItems(FG1.Row, 6, 1) = RgTABL!Descri
    Lgo(6) = RgTABL!NumSD1 + RgTABL!NumSD2
    Edita(6) = String(RgTABL!NumSD1 - 1, "0") & "0"
    If RgTABL!NumSD2 > 0 Then Edita(6) = Edita(6) & "," & String(RgTABL!NumSD2, "0")
    'Text3.DataFormat.Format = CAMPO2
    'Text3.MaxLength = Len(CAMPO2)
    Posi(6) = RgTABL!NumSD3
    CanCpos2 = CanCpos2 + 1
    RgTABL.Close
    End If
    
If UMED7 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED7 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    MatrItems(FG1.Row, 7, 1) = RgTABL!Descri
    Lgo(7) = RgTABL!NumSD1 + RgTABL!NumSD2
    Edita(7) = String(RgTABL!NumSD1 - 1, "0") & "0"
    If RgTABL!NumSD2 > 0 Then Edita(7) = Edita(7) & "," & String(RgTABL!NumSD2, "0")
    'Text3.DataFormat.Format = CAMPO2
    'Text3.MaxLength = Len(CAMPO2)
    Posi(7) = RgTABL!NumSD3
    CanCpos2 = CanCpos2 + 1
    RgTABL.Close
    End If
    

End Sub

Private Function BorraRenglon(Row As Integer)
Dim ColCell As Long

ColCell = FG1.Cell(6, Row, 0, Row, 5)

FG1.Cell(6, FG1.Row, 0, FG1.Row, 5) = vbGreen

vbMsgBoxText = "Desea Eliminar el renglón ?"
vbMsgBoxResp = MsgBox(vbMsgBoxText, vbQuestion + vbYesNo, "Carga de Factura")

If vbMsgBoxResp = vbNo Then
    FG1.Cell(6, Row, 0, Row, 5) = ColCell
    Exit Function
    End If

FG1.RemoveItem (Row)

End Function
