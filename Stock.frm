VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form Stock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso de Movimientos al Stock"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7110
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
   ScaleHeight     =   5940
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
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
      Left            =   2400
      TabIndex        =   22
      Top             =   1680
      Visible         =   0   'False
      Width           =   4365
      Begin VSFlex8Ctl.VSFlexGrid FG2 
         Height          =   1410
         Left            =   15
         TabIndex        =   23
         Top             =   195
         Width           =   4335
         _cx             =   7646
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Stock.frx":0000
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
   Begin VB.CommandButton Command3 
      Caption         =   "Limpiar"
      Height          =   315
      Left            =   420
      TabIndex        =   15
      Top             =   5460
      Width           =   1300
   End
   Begin VB.Frame Frame4 
      Height          =   465
      Left            =   360
      TabIndex        =   19
      Top             =   1170
      Width           =   6480
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   4005
         MaxLength       =   20
         TabIndex        =   11
         Top             =   165
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   10
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Despacho:"
         Height          =   240
         Left            =   2640
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Cpbante.:"
         Height          =   240
         Left            =   90
         TabIndex        =   20
         Top             =   150
         Width           =   1320
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   315
      Left            =   4050
      TabIndex        =   13
      Top             =   5490
      Width           =   1300
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   315
      Left            =   5460
      TabIndex        =   14
      Top             =   5490
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Caption         =   " Tipo de Movimiento  "
      Height          =   990
      Left            =   360
      TabIndex        =   16
      Top             =   150
      Width           =   2400
      Begin VB.OptionButton Option1 
         Caption         =   "Salidas"
         Height          =   330
         Index           =   1
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1545
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Entradas"
         Height          =   330
         Index           =   0
         Left            =   390
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   255
         Width           =   1545
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   3690
      Left            =   390
      TabIndex        =   12
      Top             =   1680
      Width           =   6405
      _cx             =   11298
      _cy             =   6509
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Stock.frx":006A
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
   Begin VB.Frame Frame2 
      Caption         =   "  Forma de Entrada"
      Height          =   990
      Left            =   2880
      TabIndex        =   17
      Top             =   165
      Visible         =   0   'False
      Width           =   3945
      Begin VB.OptionButton Option2 
         Caption         =   "Ajuste ""+"" Inv."
         Height          =   285
         Index           =   3
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   585
         Width           =   1770
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Compra Local"
         Height          =   285
         Index           =   2
         Left            =   1965
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1770
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Importación"
         Height          =   285
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   585
         Width           =   1770
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Fábrica"
         Height          =   285
         Index           =   0
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   255
         Width           =   1770
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Forma de Salida  "
      Height          =   990
      Left            =   2880
      TabIndex        =   18
      Top             =   150
      Visible         =   0   'False
      Width           =   3945
      Begin VB.OptionButton Option3 
         Caption         =   "Remito"
         Height          =   330
         Index           =   0
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1770
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ajuste -"
         Height          =   330
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   585
         Width           =   1770
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Promoción"
         Height          =   330
         Index           =   2
         Left            =   1965
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1770
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Ajuste ""-"" Inv."
         Height          =   330
         Index           =   3
         Left            =   1980
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   585
         Width           =   1770
      End
   End
End
Attribute VB_Name = "Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LaRow As Integer, FinCod As Boolean, FinCant As Boolean, EsEnter As Boolean, VeNotaC As Boolean
Dim LaTecla As Byte, SeccAnt As String, SeMueve As Boolean, SaleMal As Boolean
Dim CodEntr As Integer, CodSali As Integer, EntrSali As Integer
Dim EsNuevo As Boolean

Private Const laselec1 = "SELECT * FROM Articulo WHERE COD1 = '"

Dim dbART As New ADODB.Connection, RgART As New ADODB.Recordset
Dim dbSTOK As New ADODB.Connection, RgSTOK As New ADODB.Recordset

Private Sub Form_Load()

LaRow = 1
CodEntr = 1: CodSali = 1
VolverArt = 2

Me.Move 3500, 1600, 7200, 6300

dbART.ConnectionString = BDatos1
dbART.Open
dbSTOK.ConnectionString = BDatos1
dbSTOK.Open

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbART.Close
dbSTOK.Close
VolverArt = 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then
    Option1(0).SetFocus
    End If

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

'  Si es ENTRADA y es IMPORTACIÓN pide nro.despacho
If CodEntr = 2 And Frame2.Visible And Text2.Text = "" Then
    Text2.SetFocus
    MsgBox "Debe colocar el Nº de DESPACHO si es importación", vbCritical + vbOKOnly, "Error de Ingreso"
    Exit Sub
End If

Graba

End Sub

Private Sub Command3_Click()

Text1.Text = ""
Text2.Text = ""

    Label2.Visible = False    '  Nº Despacho Invisible
    Text2.Visible = False

LimpiaGrilla

Text1.SetFocus
FG1.ColHidden(4) = True
Frame5.Visible = False
'Option2(0).Value = True

End Sub

Private Sub Option1_Click(Index As Integer)

Option2(CodEntr - 1) = False: Option3(CodSali - 1) = False

EntrSali = Index + 1

Command3_Click

Select Case Index
    Case 0
            Frame2.Visible = True
            Frame3.Visible = False
            Option2(0).SetFocus
    Case 1
            Frame3.Visible = True
            Frame2.Visible = False
            Option3(0).SetFocus
End Select

End Sub

Private Sub Option2_Click(Index As Integer)
    
    Label2.Visible = False
    Text2.Visible = False

Command3_Click

CodEntr = Index + 1
Text1.SetFocus

'  Si es ENTRADA y es IMPORTACIÓN pide nro.despacho
If CodEntr = 2 And Frame2.Visible Then
    Label2.Visible = True
    Text2.Visible = True
End If

End Sub

Private Sub Option3_Click(Index As Integer)

CodSali = Index + 1
Text1.SetFocus

End Sub

'--------------------------------------------------------------
'           G R I L L A
'--------------------------------------------------------------

'------------------------------------------------------------
'           0-Sección         1-Número
'           2-Stock           3-Cantidad
'------------------------------------------------------------

Private Sub FG1_GotFocus()
            
If Text1.Text = "" Then Exit Sub
'FG1.Row = LaRow
' FG1.Col = 1
    FG1.EditCell
    If FG1.EditWindow = 0 Then FG1.EditCell

End Sub

Private Sub FG1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
            
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

If Col = 1 Or Col = 2 Or Col = 3 Then
            If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
            If KeyAscii = Asc(",") Then Exit Sub
End If
                
If Col <> 0 And Col <> 3 Then
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If
End If

End Sub

Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim i2%, Ii1%, Ii2%
' FinCod : Indica si ya se terminó de cargar el código de artículo (por ej. "PL")
'
' FinCant  : Indica si ya se terminó de cargar las cantidades
' CanCpos2 : Cantidad de Entradas de la unidad de medida de venta (ej. Unidades)
'

'------------------------------------------------------------------

If Not EsEnter Then Exit Sub

'If FG1.TextMatrix(Row, Col) = "" Then GoTo Abajo

If Col = 0 Then
        SeccAnt = Correa
        FG1.Col = 1
        Correa = FG1.TextMatrix(FG1.Row, 0)
        BuscaSeccion
        GoTo Abajo
End If

If Col = 1 Then
        VerSiBusca
        If SaleMal Then
            SaleMal = False
            GoTo Abajo
        End If
        FG1.Col = 3
        GoTo Abajo
End If

'------------------------------------------------------------------

AlaCant:

    If Val(FG1.TextMatrix(FG1.Row, 3)) = 0 Then
        FG1.Col = 3
        Col = 1
    End If

Abajo:

If EsEnter Then
    EsEnter = False
    
    If Col < 3 Then
       FG1.EditCell
       If FG1.EditWindow = 0 Then FG1.EditCell
    Else
        VerSiAgrega
    End If

End If

End Sub

Private Sub FG1_Click()
            
If FG1.Col = 3 Then FG1.Col = 0

    FG1.Select FG1.Row, FG1.Col
    FG1.EditCell

End Sub

Private Sub FG1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Salir

If FG1.MouseRow = -1 Then Exit Sub
If LaRow > (FG1.Rows - 1) Then LaRow = (FG1.Rows - 1)

If FG1.MouseRow > 0 Then
        FG1.ToolTipText = " " & "" & " "
    End If

If FG1.MouseRow = FG1.Row Then
   FG1.Cell(7, LaRow, 0, LaRow, 3) = &H80000012
   FG1.BackColorSel = &HFF0000
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 3) = &HFCE7C7
   LaRow = FG1.MouseRow
   Else
   FG1.BackColorSel = &H800000
   End If

If FG1.MouseRow > 0 And FG1.MouseRow < FG1.Rows Then
   FG1.Cell(7, LaRow, 0, LaRow, 3) = &H80000012
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 3) = &HFF0000
   LaRow = FG1.MouseRow
   End If

Salir:

End Sub

Private Sub FG2_Click()

FG1.TextMatrix(FG1.Row, 4) = FG2.TextMatrix(FG2.Row, 0)
Frame5.Visible = False
FG1.ColHidden(4) = False
FG1.SetFocus


End Sub

'------------------------------------------------------
'                Fin Grilla
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

'---------------------------------------------------------------------------------------
Private Sub text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then
    BuscaCPBTE
    If Not EsNuevo Then Exit Sub
    If Text2.Visible Then
        Text2.SetFocus
     Else
        FG1.Enabled = True
        FG1.SetFocus
    End If
    Exit Sub
End If


If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'---------------------------------------------------------------------------------------
Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then
    If Not EsNuevo Then Exit Sub
        FG1.Enabled = True
        FG1.SetFocus
    Exit Sub
End If

End Sub

Sub VerSiAgrega()
Dim c1%

    SeccAnt = FG1.TextMatrix(FG1.Row, 0)
    
    FG1.Rows = FG1.Rows + 1
    FG1.Row = FG1.Rows - 1
    
                    
    FG1.TextMatrix(FG1.Row, 0) = SeccAnt
    FG1.Col = 1
   
    FG1.EditCell

End Sub

Private Sub BuscaCPBTE()
Dim Codi As String
Dim Seccio     As String

Seccio = FG1.TextMatrix(FG1.Row, 1)
EsNuevo = False

If EntrSali = 1 Then
    Select Case CodEntr
        Case 1
            CodMovs = "13"
        Case 2
            CodMovs = "14"
        Case 3
            CodMovs = "15"
        Case 4
            CodMovs = "18"
    End Select
End If

If EntrSali = 2 Then
    Select Case CodSali
        Case 1
            CodMovs = "23"
        Case 2
            CodMovs = "24"
        Case 3
            CodMovs = "25"
        Case 4
            CodMovs = "28"
    End Select
End If

    miSQL2 = "SELECT * FROM Movstock WHERE TIPO = '" & CodMovs & "' AND CPBTE = " & Text1.Text
    RgSTOK.Open miSQL2, dbSTOK, adOpenForwardOnly, adLockReadOnly
    
    If Not RgSTOK.EOF Then
        RgSTOK.Close
        MsgBox "Ya existe un movimiento con ese NUMERO", vbCritical + vbOKOnly, "Error de Ingreso"
        Exit Sub
    Else
        EsNuevo = True
    End If
        
        RgSTOK.Close

End Sub

Private Sub BuscaSeccion()
Dim Codi As String

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
    
    If RgTABL!ALF7 <> "UNID " And RgTABL!ALF7 <> "MTRS " And RgTABL!ALF7 <> "KG   " Then
        RgTABL.Close
        MsgBox "               La Sección  '" & FG1.TextMatrix(FG1.Row, 0) & "'" & _
               vbCrLf & " no se vende por unidades", _
                vbCritical + vbOKOnly, "Error en Tabla de Secciones"
                FG1.Col = 0
            FG1.EditCell
            If FG1.EditWindow = 0 Then FG1.EditCell
        Exit Sub
    End If

FG1.TextMatrix(FG1.Row, 1) = ""
FG1.TextMatrix(FG1.Row, 2) = ""
FG1.TextMatrix(FG1.Row, 3) = ""

If RgTABL!ALF1 <> "NRO  " And RgTABL!ALF1 <> "MM1  " Then
    RgTABL.Close
    MsgBox "La Sección " & FG1.TextMatrix(FG1.Row, 0) & " no se vende por UNIDADES", vbCritical + vbOKOnly, "Error de Ingreso"
            FG1.Col = 0
            FG1.EditCell
    If FG1.EditWindow = 0 Then FG1.EditCell
        Exit Sub
End If
        
RgTABL.Close

End Sub

Sub VerSiBusca()
Dim PCos          As Currency
Dim PVta          As Currency
Dim Stock         As Single
Dim Stmin         As Single
Dim Seccio     As String
Dim I1%

Seccio = FG1.TextMatrix(FG1.Row, 1)
        
If FG1.FindRow(Seccio, , 1) <> -1 And FG1.FindRow(Seccio, , 1) <> FG1.Row Then
    If FG1.TextMatrix(FG1.FindRow(Seccio, , 1), 0) = FG1.TextMatrix(FG1.Row, 0) Then
        MsgBox "Artículo ya Ingresado", vbExclamation + vbOKOnly, "Duplicación"
        SaleMal = True
        Exit Sub
    End If
End If
        
miSQL2 = "SELECT  * FROM Articulo WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         Seccio & "'"
RgART.Open miSQL2, dbART, adOpenForwardOnly, adLockReadOnly

    If RgART.EOF = True Then
        RgART.Close
        Busqueda.BuscaDesdeAfuera
        Exit Sub
    End If
        
        FG1.TextMatrix(FG1.Row, 2) = 0

miSQL2 = "SELECT  * FROM Stock WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         Seccio & "'"
RgSTOK.Open miSQL2, dbSTOK, adOpenForwardOnly, adLockReadOnly

    If Not RgSTOK.EOF Then
        FG1.TextMatrix(FG1.Row, 2) = RgSTOK!STUnid
    End If
    
DesArt = Trim(RgART!Descri)

Correa = RgART!COD1
RESTO = RgART!Cod2
        
RgART.Close
RgSTOK.Close

' ----------------   Despachos  ---------------

If EntrSali = 1 Then GoTo SiEsEntrada
'  Si es Salida
        
'---------------------------------------
'  Despachos
'--------------------------------------------
    Frame5.Visible = False
        
miSQL2 = "SELECT  * FROM Despachos WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         RESTO & "' and STOCK > 0"
RgART.Open miSQL2, dbART, adOpenForwardOnly, adLockReadOnly

    If RgART.EOF = True Then
        RgART.Close
        Exit Sub
    End If
    
FG2.Clear flexClearScrollable

For i = (FG2.Rows - 1) To 1 Step -1
FG2.RemoveItem (i)
Next i

Do Until RgART.EOF

FG2.AddItem RgART!NRODESP & vbTab & RgART!fecent & vbTab & RgART!Stock, FG2.Row + 1

RgART.MoveNext

Loop

RgART.Close
       
FG2.AutoSize 0, 2
    
'        Frame5.top = FG1.RowPos(FG1.Row)
        Frame5.Visible = True
        FG2.SetFocus

GoTo NoHayDespachos

SiEsEntrada:
'  Si es Entrada
'  por importación

If Text2.Text = "" Then GoTo NoHayDespachos
    
Frame5.Visible = False

' Busca si el despacho para ese art. ya se cargó
miSQL2 = "SELECT  * FROM Despachos WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         Seccio & "' and NRODESP = '" & Text2.Text & "'"
RgSTOK.Open miSQL2, dbSTOK, adOpenForwardOnly, adLockReadOnly

    If Not RgSTOK.EOF Then
        MsgBox "Artículo ya Ingresado para este DESPACHO", vbExclamation + vbOKOnly, "Duplicación"
        SaleMal = True
    End If
        
RgSTOK.Close

NoHayDespachos:

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

Sub LimpiaGrilla()
Dim I1%
   
FG1.Clear flexClearScrollable

For i = (FG1.Rows - 1) To 1 Step -1
FG1.RemoveItem (i)
Next i

FG1.Rows = 2
FG1.Col = 0
FG1.Row = 1

End Sub

Sub Graba()
Dim I0%, I1%, i2%, item As Single, CanItems As Integer
Dim NoAnduvo As Boolean

  vbMsgBoxTitle = "Grabación del Movimiento"
  vbMsgBoxText = "El siguiente proceso registrará en Movimiento en el Stock  " & vbCrLf & "Desea continuar ?  "
  vbMsgBoxResp = vbYesNoCancel + vbExclamation + vbApplicationModal + vbDefaultButton1 '
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

dbSTOK.BeginTrans


'---------------------------------------------------------
'         Graba Detalle - Stock,Movstck
'---------------------------------------------------------

If FG1.TextMatrix(FG1.Rows - 1, 3) <> "" Then
    FG1.Rows = FG1.Rows + 1
End If

CanItems = 0

For I1 = 1 To FG1.Rows

If FG1.TextMatrix(I1, 3) = "" Then GoTo Confirmar

CanItems = CanItems + 1

Correa = FG1.TextMatrix(I1, 0)
RESTO = FG1.TextMatrix(I1, 1)

item = 0

Graba_1:
    
'------------------------------------
'    Stock - Stock
'------------------------------------

miSQL1 = "SELECT * FROM Stock WHERE COD1 = '" & Correa & _
        "' and COD2 = '" & RESTO & "'"
        
RgSTOK.Open miSQL1, dbSTOK, adOpenDynamic, adLockPessimistic
    
    If RgSTOK.EOF Then
        RgSTOK.AddNew
        RgSTOK!COD1 = FG1.TextMatrix(I1, 0)
        RgSTOK!Cod2 = FG1.TextMatrix(I1, 1)
        RgSTOK!Stmin = 0
        RgSTOK!StmAX = 0
        RgSTOK!StREP = 0
        RgSTOK!DEP1 = 0
        RgSTOK!EST1 = 0
        RgSTOK!ESTT1 = 0
        RgSTOK!DEP2 = 0
        RgSTOK!EST2 = 0
        RgSTOK!ESTT2 = 0
        RgSTOK!STUnid = 0
        RgSTOK!Pulg = 0
        RgSTOK!Mtr = 0
        RgSTOK!Stant = 0
        RgSTOK!ENTMES = 0
        RgSTOK!SALMES = 0
        RgSTOK!AJEMES = 0
        RgSTOK!AJSMES = 0
        RgSTOK!STUnid = 0
    End If
    
    If EntrSali = 1 Then
        RgSTOK!STUnid = RgSTOK!STUnid + FG1.TextMatrix(I1, 3)
        RgSTOK!ENTMES = RgSTOK!ENTMES + FG1.TextMatrix(I1, 3)
    Else
        RgSTOK!STUnid = RgSTOK!STUnid - FG1.TextMatrix(I1, 3)
        RgSTOK!SALMES = RgSTOK!SALMES + FG1.TextMatrix(I1, 3)
    End If

    RgSTOK!FACTUAL = Date
    RgSTOK!USUARIO = Mid(Red_Usuario, 1, 6)
    
    RgSTOK.Update
    RgSTOK.Close
    
'------------------------------------
'   Movimientos de Stock - MovStock
'------------------------------------

'
' Cód.Movim.: 11-N/Créd.'A'       12-N/Créd.'B'          13-Fábricación
'             14-Compra           15-Aj.en Más           16-Cpra.'R'
'             17-N/Créd.'R'       18-Aj.Inventario       19-

'             21-Factura 'A'      22-Fact.'B'            23-Remito
'             24-Promoción        25-Aj.en Menos         26-
'             27-Fact.'R'         28-Aj.en menos Invent. 29-
'
If EntrSali = 1 Then
    Select Case CodEntr
        Case 1
            CodMovs = "13"
        Case 2
            CodMovs = "15"
        Case 3
            CodMovs = "14"
        Case 4
            CodMovs = "18"
    End Select
End If

If EntrSali = 2 Then
    Select Case CodSali
        Case 1
            CodMovs = "23"
        Case 2
            CodMovs = "25"
        Case 3
            CodMovs = "24"
        Case 4
            CodMovs = "28"
    End Select
End If

miSQL1 = "SELECT * FROM MovStock WHERE COD1= '" & Correa & "' and COD2 = '" & _
        RESTO & "' and CPBTE = " & Text1.Text & " AND PTOVTA = " & _
        PtoVta & " AND TIPO = '" & CodMovs & "'"
RgSTOK.Open miSQL1, dbSTOK, adOpenDynamic, adLockPessimistic
    
'    If Not RgSTOK.EOF Then
'        NoAnduvo = True
'        RgSTOK.Close
'        GoTo Confirmar
 '       GoTo SaltaAdd1
'    End If

    RgSTOK.AddNew
    RgSTOK!COD1 = FG1.TextMatrix(I1, 0)
    RgSTOK!Cod2 = FG1.TextMatrix(I1, 1)
    RgSTOK!TIPO = CodMovs
    RgSTOK!PtoVta = 1
    RgSTOK!Cpbte = Text1.Text
    
SaltaAdd1:
    RgSTOK!item = item
    RgSTOK!FECHA = Now
    RgSTOK!Cant = Val(FG1.TextMatrix(I1, 3))
        
        RgSTOK!Pulg = 0
        RgSTOK!Mtr = 0
    
    RgSTOK!Milim = 0
    RgSTOK!Telas = 0

    RgSTOK!USUARIO = Mid(Red_Usuario, 1, 6)
    
    RgSTOK.Update
    RgSTOK.Close


'--------  Despachos   ----------

miSQL1 = "SELECT  * FROM Despachos WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         RESTO & "' and NRODESP = '" & Text2.Text & "'"
RgSTOK.Open miSQL1, dbSTOK, adOpenDynamic, adLockPessimistic
    
    If RgSTOK.EOF Then
        RgSTOK.AddNew
        RgSTOK!COD1 = FG1.TextMatrix(I1, 0)
        RgSTOK!Cod2 = FG1.TextMatrix(I1, 1)
        RgSTOK!NRODESP = Text2.Text
        RgSTOK!Cpbte = Text1.Text
        RgSTOK!Stock = 0
        RgSTOK!ENTRADA = 0
        RgSTOK!salidaS = 0
        RgSTOK!fecent = Date
        RgSTOK!fecusal = Date
    End If
    
    If EntrSali = 1 Then
        RgSTOK!Stock = RgSTOK!Stock + FG1.TextMatrix(I1, 3)
        RgSTOK!ENTRADA = RgSTOK!ENTRADA + FG1.TextMatrix(I1, 3)
    Else
        RgSTOK!Stock = RgSTOK!Stock - FG1.TextMatrix(I1, 3)
        RgSTOK!salidaS = RgSTOK!salidaS + FG1.TextMatrix(I1, 3)
    End If

    RgSTOK.Update
    RgSTOK.Close

Next I1

'--------------------------------

Confirmar:

'  Si salió todo bien, Confirma las Transacciones
If NoAnduvo Then
    dbSTOK.RollbackTrans
    MsgBox "Hubo Errores durante la GRABACIÓN" & vbCrLf & "VERIFIQUE !!!", vbExclamation, "Aviso"
Else
    dbSTOK.CommitTrans
End If

Command3_Click
Option2(0).Value = True

Exit Sub

'---------------------------------------------
Salir1:

Unload Me

Salir2:

End Sub

