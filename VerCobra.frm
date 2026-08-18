VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form VerCobra 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cobranzas por Zona"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11520
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
   ScaleHeight     =   5100
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "Contraer"
      Height          =   285
      Index           =   1
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   300
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Expandir"
      Height          =   285
      Index           =   0
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   330
      Left            =   10170
      TabIndex        =   3
      Top             =   4590
      Width           =   1185
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00F9FADC&
      Height          =   315
      ItemData        =   "VerCobra.frx":0000
      Left            =   4800
      List            =   "VerCobra.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   1935
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   3915
      Left            =   60
      TabIndex        =   2
      Top             =   570
      Width           =   11325
      _cx             =   19976
      _cy             =   6906
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14089948
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   32768
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14548991
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"VerCobra.frx":0004
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Deuda Venc.:"
      Height          =   240
      Left            =   5610
      TabIndex        =   9
      Top             =   4680
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Deuda Total :"
      Height          =   240
      Left            =   1980
      TabIndex        =   8
      Top             =   4680
      Width           =   1590
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3570
      TabIndex        =   5
      Top             =   4620
      Width           =   1785
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6960
      TabIndex        =   4
      Top             =   4620
      Width           =   1785
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Zona :"
      Height          =   240
      Left            =   3810
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "VerCobra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbMAE As New ADODB.Connection
Dim RgCLTE As New ADODB.Recordset, RgCCTE As New ADODB.Recordset
Dim TotDeu As Currency, TotVenc As Currency

Private Sub Form_Load()

Me.Move 0, 0, 11600, 5500

dbMAE.ConnectionString = BDatos1
dbMAE.Open

CargaZona

Combo2.ListIndex = 0

'dover1

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbMAE.Close

End Sub

Private Sub Form_Deactivate()

'Unload Me

End Sub

Private Sub Combo2_Click()

Option1(1).Value = False

DoVer1

Option1(1).Value = True

End Sub

Private Sub Command1_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Option1_Click(Index As Integer)

If Index = 0 Then
       FG1.Outline -1
Else
       FG1.Outline 0
End If

End Sub

Sub DoVer1()
       
LimpiaGrilla

Ordenar = " ORDER BY NOMB"

miSQL1 = "SELECT * FROM Clientes WHERE zona = " & _
          Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & Ordenar
RgCLTE.Open miSQL1, dbMAE, adOpenForwardOnly, adLockReadOnly

If RgCLTE.EOF Then
    RgCLTE.Close
    Exit Sub
End If

CargaGrilla

    
'paraSTATUS = "Seleccionados : " & RgCCTE.RecordCount
'FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

  Label5.Caption = Format(TotDeu, "##,##0.00")
  Label6.Caption = Format(TotVenc, "##,##0.00")
    
    RgCLTE.Close


End Sub

Sub CargaGrilla()
Dim LosDias As String * 4, Dias As Integer, Estado As String * 5

Dim FECHA         As Date
Dim Cpbte         As String
Dim CLTE          As String * 5
Dim Nomb          As String * 25

Dim I1, i2, Tc
Dim i3 As Long

TotDeu = 0: TotVenc = 0

FG1.Col = 0

Do Until RgCLTE.EOF

miSQL = "Select * FROM CtasCtes WHERE  CLTE = " & RgCLTE!CODIGO & _
        " AND ( tipo = 1 or TIPO = 3 ) AND DEBE > 1 ORDER BY FECHA, CPBTE, TIPO"
RgCCTE.Open miSQL, dbMAE, adOpenForwardOnly, adLockReadOnly

If RgCCTE.EOF Then
    RgCCTE.Close
    GoTo OtroClte
End If
    
    Do Until RgCCTE.EOF
    
    i3 = i3 + 1
    
    FECHA = RgCCTE!FECHA
    RSet CLTE = RgCCTE!CLTE
    LSet Nomb = RgCLTE!Nomb
        
        Select Case RgCCTE!TIPO
            Case 1:      Cpbte = "Fact. "
            Case 3:      Cpbte = "N/D.  "
        End Select
    
    Cpbte = Cpbte & RgCCTE!Letra & " "
    Cpbte = Cpbte & String((2 - Len(RgCCTE!prefijo)), "0") & RgCCTE!prefijo & "-"
    Cpbte = Cpbte & String(6 - (Len(RgCCTE!Cpbte)), "0") & RgCCTE!Cpbte
        
    Dias = 0
    Estado = "": LosDias = ""
    TotDeu = TotDeu + RgCCTE!DEBE
    
        If RgCCTE!FECVTO < Date Then
            Dias = Date - RgCCTE!FECVTO
            Estado = "Venc."
            TotVenc = TotVenc + RgCCTE!DEBE
            RSet LosDias = Format(Dias, "###0")
        End If
    
    FG1.AddItem CLTE & " " & Nomb & vbTab & Cpbte & vbTab & RgCCTE!FECHA & vbTab & _
                RgCCTE!FECVTO & vbTab & RgCCTE!Impte & vbTab & RgCCTE!DEBE & vbTab & _
                LosDias & vbTab & Estado, FG1.Rows
        
    paraSTATUS = "Cargando : " & (FG1.Rows - 1)
    FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS
    
'    Me.Show
'    DoEvents
    
    RgCCTE.MoveNext
    
    Loop
    
RgCCTE.Close
    
OtroClte:

RgCLTE.MoveNext

Loop

FG1.MergeCells = flexMergeRestrictAll
FG1.MergeCol(-1) = True
FG1.Redraw = False

' set up outlining
'        FG1.OutlineBar = flexOutlineBarComplete
        FG1.OutlineBar = flexOutlineBarComplete
    

        ' sort datos desde first to last column

' Cuelta columnas visibles
Dim CVis As Integer
Dim G1 As Integer
CVis = 0
For G1 = 0 To FG1.Cols - 1
If Not FG1.ColHidden(G1) Then CVis = CVis + 1
Next
'        FG1.Select 0, 0, 1, (FG1.Cols - 1)
        FG1.Select 0, 0, 1, CVis

        FG1.Sort = flexSortGenericAscending
        FG1.Select 0, 0
        
        ' calcula subtotales
        FG1.SubtotalPosition = flexSTBelow
        FG1.SubTotal flexSTClear
        FG1.SubTotal flexSTSum, -1, 4, , 1, vbWhite, True, "Total", 5
        FG1.SubTotal flexSTSum, -1, 5, , 1, vbWhite, True, "Total", 5
        FG1.SubTotal flexSTSum, 0, 4, , vbWhite, vbBlue
        FG1.SubTotal flexSTSum, 0, 5, , vbWhite, vbBlue
        
FG1.Redraw = True
'FG1.AutoSize 0, 7

If FG1.Rows > 1 Then FG1.Row = 1


' FG1.SetFocus

End Sub

Sub LimpiaGrilla()
Dim I1%
   
FG1.Clear flexClearScrollable

For i = (FG1.Rows - 1) To 1 Step -1
FG1.RemoveItem (i)
Next i

End Sub

Private Sub CargaZona()
Dim I1

    ' Tabla de Zonas
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'ZN   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

Do Until RgTABL.EOF
Combo2.AddItem Mid(RgTABL!COD, 1, 3) & "-" & RgTABL!Descri
RgTABL.MoveNext
Loop

RgTABL.Close

End Sub
