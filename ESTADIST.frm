VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ESTADIST 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Ventas de Artículos por Cliente"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   11205
   ShowInTaskbar   =   0   'False
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
      Left            =   5100
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Seleccione la Sección a Visualizar"
      Top             =   60
      Width           =   1335
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
      Left            =   9180
      TabIndex        =   2
      Top             =   3990
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   330
      Left            =   6480
      TabIndex        =   1
      Top             =   450
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "ddd d MMM yyy"
      Format          =   21037059
      CurrentDate     =   36877
   End
   Begin VSFlex8Ctl.VSFlexGrid FG3 
      Height          =   3090
      Left            =   120
      TabIndex        =   5
      Top             =   810
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
      FormatString    =   $"ESTADIST.frx":0000
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
      Left            =   6930
      TabIndex        =   9
      Top             =   4020
      Width           =   1965
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
      Left            =   1650
      TabIndex        =   8
      Top             =   4020
      Width           =   1905
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
      Left            =   210
      TabIndex        =   7
      Top             =   4050
      Width           =   1425
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
      Left            =   5130
      TabIndex        =   6
      Top             =   4050
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Sección :"
      Height          =   285
      Left            =   4140
      TabIndex        =   4
      Top             =   90
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Muestra desde :"
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
      Left            =   4530
      TabIndex        =   3
      Top             =   510
      Width           =   1950
   End
End
Attribute VB_Name = "ESTADIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbEST1 As New ADODB.Connection, RgEST1 As New ADODB.Recordset, RgCLTE As New ADODB.Recordset

Private Sub Combo2_Click()

DoVer3

End Sub

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Form_Load()
Dim I1%

FECHA.Value = Date
FECHA.Day = 1

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

RgTABL.Close

'DoVer3

End Sub

Private Sub Form_Unload(Cancel As Integer)
'FCMENU.Toolbar2.Buttons(7).Value = tbrUnpressed

End Sub

Private Sub FECHA_Change()

DoVer3

End Sub

Sub DoVer3()
Dim TIPO As Byte
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2

FECHADsd = FECHA.Value
ElDia = 31
If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
FECHADsd = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year
FECHAHst = FECHA.Month & "/" & ElDia & "/" & FECHA.Year
  
If FECHADsd = FECHAHst Then
    If FECHA.Month = 12 Then
        FECHAHst = "01/01/" & FECHA.Year
    Else
        FECHAHst = FECHA.Month + 1 & "/01/" & FECHA.Year
    End If
End If

  If ORDENADOx <> "" Then
         Clasifica = elORDEN & ORDENADOx & ORDENADOad
  Else
         Clasifica = ""
End If
'///           ----------------------------------------

laLOGIK1 = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year
'miSQL = laselec4 & Combo1.Text & " * " & laselec8 & Correa & laselec6 & RESTO & laselec7 & Separa & FECHA.Value & Separa & Clasifica
'miSQL = "SELECT TOP " & Combo1.Text & " * FROM FCEstad1 WHERE COD1 = '" & Combo2.Text &
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

FG3.AddItem FECHA & vbTab & TipoCpbt & " " & Letra & " " & PtoVta & "-" & Cpbte & _
            vbTab & ElClte & " " & ElNomb & vbTab & Pulg & vbTab & Mtr & vbTab & MM & vbTab & _
            Telas & vbTab & Cant & vbTab & PVta & vbTab & Impte, FG3.Rows

CanReg = CanReg + 1

paraSTATUS = "Cargando : " & (FG3.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

Me.Show
DoEvents

Salta:

RgEST1.MoveNext
Loop

    RgEST1.Close

FG3.AutoSize 0, 9

Label5.Caption = Format(CanReg, "##,##0 ")
Label7.Caption = Format(Total, "$ ###,##0.00 ")

If FG3.Rows > 1 Then FG3.Row = 1

' FG1.SetFocus

End Sub


