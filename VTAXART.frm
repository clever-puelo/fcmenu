VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form VTAXART 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Ventas por Artículo"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "Expandir"
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
      Index           =   0
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   150
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Contraer"
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
      Index           =   1
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   450
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FEFADE&
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
      ItemData        =   "VTAXART.frx":0000
      Left            =   8100
      List            =   "VTAXART.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Año a Visualizar"
      Top             =   450
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
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
      Left            =   7860
      TabIndex        =   6
      Top             =   3990
      Width           =   1590
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
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
      Left            =   7290
      TabIndex        =   2
      Top             =   90
      Width           =   1890
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   4140
      TabIndex        =   1
      Top             =   90
      Width           =   1890
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FEFADE&
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
      ItemData        =   "VTAXART.frx":0004
      Left            =   5295
      List            =   "VTAXART.frx":002F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Seleccione el MES a visualizar"
      Top             =   450
      Width           =   2700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
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
      Left            =   9525
      TabIndex        =   5
      Top             =   3990
      Width           =   1590
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FEFADE&
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
      Left            =   1845
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Seleccione la Sección a Visualizar"
      Top             =   60
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid FG3 
      Height          =   3090
      Left            =   210
      TabIndex        =   7
      Top             =   825
      Width           =   10980
      _cx             =   19368
      _cy             =   5450
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
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"VTAXART.frx":0098
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
   Begin VB.Label Label9 
      Caption         =   "Desde :"
      Height          =   285
      Left            =   3345
      TabIndex        =   15
      Top             =   90
      Width           =   960
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta :"
      Height          =   285
      Left            =   6570
      TabIndex        =   14
      Top             =   90
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Mes  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3285
      TabIndex        =   13
      Top             =   510
      Width           =   1950
   End
   Begin VB.Label Label1 
      Caption         =   "Sección :"
      Height          =   285
      Left            =   885
      TabIndex        =   12
      Top             =   90
      Width           =   960
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Importe :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3750
      TabIndex        =   11
      Top             =   4080
      Width           =   1785
   End
   Begin VB.Label Label8 
      Caption         =   "Tot. Movim. :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   10
      Top             =   4065
      Width           =   1425
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1740
      TabIndex        =   9
      Top             =   4035
      Width           =   1905
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5550
      TabIndex        =   8
      Top             =   4050
      Width           =   1965
   End
End
Attribute VB_Name = "VTAXART"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbEST1 As New ADODB.Connection, RgEST1 As New ADODB.Recordset, RgCLTE As New ADODB.Recordset
Dim Nivel As Integer

Private Sub Combo3_Click()

DoVer3

End Sub

Private Sub Combo2_Click()

DoVer3

End Sub

Private Sub Combo4_Click()

DoVer3

End Sub

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

Imprimir

End Sub

Private Sub Form_Load()
Dim I1%, Aa1#

Combo3.ListIndex = Month(Date) - 1

Me.Move 0, 800, 11300, 4800

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

Text1.Text = 0: Text2.Text = 0

RgTABL.Close
Aa1 = Year(Date) - 6

For I1 = 0 To 6

Combo4.AddItem Aa1
Combo4.ItemData(I1) = I1
Aa1 = Aa1 + 1

Next I1

Combo4.ListIndex = 6
'DoVer3

End Sub

Private Sub Form_Unload(Cancel As Integer)
'FCMENU.Toolbar2.Buttons(7).Value = tbrUnpressed

End Sub

Private Sub FECHA_Change()

DoVer3

End Sub

Sub DoVer3()
Dim TIPO As Byte, elmes As String
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2

If Text1.Text = "" Then Exit Sub
elmes = Combo3.ListIndex + 1
If Len(elmes) = 1 Then elmes = "0" & elmes

FECHADsd = elmes & "/01/" & Combo4.Text
ElDia = 31
If elmes = 4 Or elmes = 6 Or elmes = 9 Or elmes = 11 Then ElDia = 30
If elmes = 2 Then ElDia = 28
FECHAHst = elmes & "/" & ElDia & "/" & Combo4.Text

'///           ----------------------------------------

laLOGIK1 = elmes & "/01/" & Combo4.Text
'miSQL = laselec4 & Combo1.Text & " * " & laselec8 & Correa & laselec6 & RESTO & laselec7 & Separa & FECHA.Value & Separa & Clasifica
miSQL = "SELECT  * FROM FCEstad1 WHERE COD1 = '" & Combo2.Text & _
        "' and FECHA >= " & Separa & FECHADsd & Separa & " AND FECHA < " & Separa & _
        FECHAHst & Separa & " order by fecha, cpbte"
RgEST1.Open miSQL, dbTABL, adOpenForwardOnly, adLockReadOnly
   
If RgEST1.EOF Then
    RgEST1.Close
    Exit Sub
End If

' paraSTATUS = "Seleccionados : " & RgEST1.RecordCount
'FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

       LimpiaGrilla
       CargaGrilla

End Sub

Sub LimpiaGrilla()
Dim I1%
   
FG3.Clear flexClearScrollable

For i = (FG3.Rows - 1) To 1 Step -1
FG3.RemoveItem (i)
Next i

End Sub

Sub CargaGrilla()
Dim CanReg As Double

Dim FECHA        As Date
Dim TIPO         As String
Dim TipoCpbt     As String * 6
Dim PtoVta       As String * 4
Dim Cpbte        As String * 6
Dim Letra        As String * 1
Dim ElClte       As String * 5
Dim ElNomb       As String * 20
Dim Secc         As String * 5
Dim PEsp         As Currency
Dim PVta         As Currency
Dim Impte        As Currency
Dim Total        As Currency
Dim Pulg         As Double
Dim Mtr          As Double
Dim MM           As Integer
Dim Telas        As Integer
Dim Cant         As Double

Dim I1, i2, Tc
Dim i3 As Long

FG3.Col = 0


Do Until RgEST1.EOF

If Val(RgEST1!Cod2) < Val(Text1.Text) Or Val(RgEST1!Cod2) > Val(Text2.Text) Then
    GoTo Salta
    End If

i3 = i3 + 1

FECHA = RgEST1!FECHA
Letra = RgEST1!Letra
PtoVta = String(4 - Len(RgEST1!PtoVta), "0") & RgEST1!PtoVta
Cpbte = String(6 - Len(RgEST1!Cpbte), "0") & RgEST1!Cpbte

Secc = RgEST1!COD1
Pulg = RgEST1!Pulg
Mtr = RgEST1!Mtr
MM = RgEST1!Milim
Telas = RgEST1!Telas
Cant = RgEST1!Cant

Impte = RgEST1!Impte
PVta = RgEST1!PVta
PEsp = RgEST1!PEsp

If RgEST1!TIPO = 1 Or RgEST1!TIPO = 3 Or RgEST1!TIPO = 7 Then
    Total = Total + RgEST1!Impte
Else
    Total = Total - RgEST1!Impte
End If
   
   
Select Case RgEST1!TIPO
    Case 1
        TipoCpbt = "Fact."
    Case 2
        TipoCpbt = "NCréd."
    Case 3
        TipoCpbt = "NDéb."
    Case 9
        TipoCpbt = "Anul."
    Case Else
        TipoCpbt = " -- "
End Select

RSet ElClte = RgEST1!CLTE
miSQL = "Select nomb FROM Clientes WHERE CODIGO = " & RgEST1!CLTE

RgCLTE.Open miSQL, dbTABL, adOpenForwardOnly, adLockReadOnly
If Not RgCLTE.EOF Then
    ElNomb = RgCLTE!Nomb
Else
    ElNomb = "*** Desconocido ***"
End If
    
    RgCLTE.Close

FG3.AddItem Pulg & vbTab & MM & vbTab & Telas & vbTab & Mtr & vbTab & Cant & vbTab & _
            PVta & vbTab & Impte & vbTab & FECHA & vbTab & TipoCpbt & " " & Letra & " " & _
            PtoVta & "-" & Cpbte & vbTab & ElClte & " " & ElNomb, FG3.Rows

CanReg = CanReg + 1

paraSTATUS = "Cargando : " & (FG3.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

Me.Show
DoEvents

Salta:

RgEST1.MoveNext
Loop

Salta2:

    RgEST1.Close

FG3.Col = 1
FG3.ColSel = 2
FG3.Sort = flexSortNumericAscending

Label5.Caption = Format(CanReg, "##,##0 ")
Label7.Caption = Format(Total, "$ ###,##0.00 ")

If FG3.Rows > 1 Then FG3.Row = 1


FG3.MergeCells = flexMergeRestrictAll
FG3.MergeCol(-1) = True
FG3.Redraw = False

' set up outlining
'        FG1.OutlineBar = flexOutlineBarComplete
        FG3.OutlineBar = flexOutlineBarComplete
    

        ' sort datos desde first to last column

 '       FG3.Select 0, 0, 1, (FG3.Cols - 1)

  '      FG3.Sort = flexSortGenericAscending
  '      FG3.Select 0, 0
        FG3.SubtotalPosition = flexSTBelow
        FG3.SubTotal flexSTClear
        FG3.SubTotal flexSTSum, -1, 3, , 1, vbWhite, True, "Total", 5
        FG3.SubTotal flexSTSum, -1, 4, , 1, vbWhite, True, "Total", 5
        FG3.SubTotal flexSTSum, -1, 6, , 1, vbWhite, True, "Total", 5
FG3.ColSel = 2
        
FG3.Redraw = True

        ' calcula subtotales
Select Case Combo2.Text
    Case "GPS       "

            FG3.Col = 1
            FG3.ColSel = 2
            FG3.Sort = flexSortNumericAscending
        FG3.SubTotal flexSTSum, 1, 3, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 1, 6, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 2, 3, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 2, 6, , vbWhite, vbBlue

        Nivel = 2
     Case "CT        "

            FG3.Col = 1
            FG3.ColSel = 2
            FG3.Sort = flexSortNumericAscending
        FG3.SubTotal flexSTSum, 1, 3, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 1, 6, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 2, 3, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 2, 6, , vbWhite, vbBlue

        Nivel = 2
       
    Case "PL        "
            FG3.Col = 0
            FG3.ColSel = 2
            FG3.Sort = flexSortNumericAscending
        FG3.SubTotal flexSTSum, 0, 3, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 0, 6, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 1, 3, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 1, 6, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 2, 3, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 2, 6, , vbWhite, vbBlue

        Nivel = 2
        
    Case "SF        "
            FG3.Col = 3
            FG3.ColSel = 3
            FG3.Sort = flexSortNumericAscending
        FG3.SubTotal flexSTSum, 3, 4, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 3, 6, , vbWhite, vbBlue

        Nivel = 3
        
    Case Else
            FG3.Col = 0
            FG3.ColSel = 0
            FG3.Sort = flexSortNumericAscending
        FG3.SubTotal flexSTSum, 0, 3, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 0, 4, , vbWhite, vbBlue
        FG3.SubTotal flexSTSum, 0, 6, , vbWhite, vbBlue
        
        Nivel = 0
        
End Select
        
        FG3.Outline Nivel

FG3.AutoSize 0, 9
Option1(1).Value = True

' FG1.SetFocus

End Sub

Private Sub Imprimir()

Dim s$, h$, d$, I1

Dim Cpbte As Double, FecCpbt As String
Dim Cliente As String, Nro As String, Mtr As String, MM As String
Dim Tela As String, Cant As String, Precio As String, Importe As String
Dim Total As Currency

IMPRE.Caption = " Impresión del Stock "
IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprA4
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"
        
'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 500

'Primer Rectángulo
        DibujaRect 13, 50, 35, 170, 30, 30
        DibujaRect 40, 3, 50, 205, 30, 30
        
'        DibujaLine 80, 7, 80, 205

Escribe "", 10, 0, 0, 3, 160, Format(Date, "ddd dd MMM YYYY")
'Escribe "", 12, 0, 0, 51, 153, Format(Time, "hh:mm")

        Escribe "ARIAL BLACK", 16, 1, 0, 15, 75, "Ventas de Artículos"
        Escribe "ARIAL BLACK", 16, 1, 0, 22, 80, "Sección :  " & Combo2.Text
        Escribe "ARIAL BLACK", 16, 1, 0, 27, 65, "Desde : " & Text1.Text
        Escribe "ARIAL BLACK", 16, 1, 0, 27, 110, "Hasta : " & Text2.Text

'  Ttulo
    Escribe "", 10, 1, 1, 43, 12, "Nro/Pulg"
    Escribe "", 10, 1, 1, 43, 30, "Mt/Kg"
    Escribe "", 10, 1, 1, 43, 40, "MM"
    Escribe "", 10, 1, 1, 43, 50, "T."
    Escribe "", 10, 1, 1, 43, 58, "Cant."
    Escribe "", 10, 1, 1, 43, 78, "Precio"
    Escribe "", 10, 1, 1, 43, 98, "Importe"
'    Escribe "", 10, 1, 1, 43, 120, "Fecha"
'    Escribe "", 10, 1, 1, 43, 140, "C L I E N T E"

    IMPRE.VP1.TableBorder = tbNone

s$ = ">+20mm|>+10mm|>+9mm|>+9mm|>+20mm|>+20mm|>+25mm|<+35mm|<+55mm"
h$ = "Nro/Pulg|Mtr/Kg|MM|T.|Cant.| Precio | Importe  |  Fecha  |        Cliente          "

With IMPRE.VP1
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Lucida Console"
    .CurrentY = "52mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "4mm"
 End With

'FG3.Outline 1

For I1 = 1 To FG3.Rows - 1

If Option1(1).Value Then
    If Mid(FG3.TextMatrix(I1, Nivel), 1, 5) <> "Total" Then GoTo Saltito
End If

If Mid(FG3.TextMatrix(I1, 1), 1, 5) = "Total" Then FG3.TextMatrix(I1, 1) = ">>"
If Mid(FG3.TextMatrix(I1, 2), 1, 5) = "Total" Then FG3.TextMatrix(I1, 2) = ">>"

Nro = Mid(FG3.TextMatrix(I1, 0), 6, 9)
Mtr = Format(FG3.TextMatrix(I1, 1), "##,###")
MM = Format(FG3.TextMatrix(I1, 2), "#0")
Tela = Format(FG3.TextMatrix(I1, 3), "#0")
Cant = Format(FG3.TextMatrix(I1, 4), "##,##0")
Precio = Format(FG3.TextMatrix(I1, 5), "###,##0.00")
Importe = Format(FG3.TextMatrix(I1, 6), "###,##0.00")
FecCpbt = Format(FG3.TextMatrix(I1, 7), "ddd dd MMM YYYY")
Cliente = FG3.TextMatrix(I1, 9)

d$ = Nro & "|" & Mtr & "|" & MM & "|" & Tela & "|" & Cant & "|" & _
     Precio & "|" & Importe & "|" & FecCpbt & "|" & Cliente

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 3
IMPRE.VP1.AddTable s$, h$, "||", , , True

Saltito:

Next I1

'    IMPRE.VP1.BrushStyle = bsSolid
'    IMPRE.VP1.BrushColor = &HE0E0E0
'    DibujaRect 260, 166, 268, 202, 30, 30
'    IMPRE.VP1.BrushStyle = 1
    
'    Escribe "Arial", 10, 1, 1, 255, 179, "TOTAL"

 IMPRE.VP1.EndDoc
 
IMPRE.Show 1

LimpiaGrilla

End Sub

Private Sub Option1_Click(Index As Integer)

If Index = 0 Then
       FG3.Outline -1
Else
       FG3.Outline Nivel
End If

End Sub

