VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CpraClte 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compras del Cliente"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11175
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
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
   ScaleHeight     =   4335
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "CpraClte.frx":0000
      Left            =   3690
      List            =   "CpraClte.frx":0019
      TabIndex        =   11
      Text            =   "50"
      ToolTipText     =   "Cantidad máximo de Renglones a Visualizar"
      Top             =   480
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9060
      TabIndex        =   10
      Top             =   3630
      Width           =   1590
   End
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   330
      Left            =   6480
      TabIndex        =   0
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
      Format          =   25493507
      CurrentDate     =   36877
   End
   Begin VSFlex8Ctl.VSFlexGrid FG3 
      Height          =   2670
      Left            =   30
      TabIndex        =   9
      Top             =   900
      Width           =   10770
      _cx             =   18997
      _cy             =   4710
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
      FormatString    =   $"CpraClte.frx":003A
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
      Left            =   5010
      TabIndex        =   8
      Top             =   3630
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
      Left            =   90
      TabIndex        =   7
      Top             =   3630
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
      Left            =   1530
      TabIndex        =   6
      Top             =   3600
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
      Left            =   6810
      TabIndex        =   5
      Top             =   3600
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2730
      TabIndex        =   4
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3675
      TabIndex        =   3
      Top             =   105
      Width           =   4950
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Renglones :"
      Height          =   240
      Left            =   2370
      TabIndex        =   2
      Top             =   510
      Width           =   1320
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Muestra desde:"
      Height          =   240
      Left            =   4530
      TabIndex        =   1
      Top             =   510
      Width           =   1950
   End
End
Attribute VB_Name = "CpraClte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbEST1 As New ADODB.Connection, RgEST1 As New ADODB.Recordset

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Combo1_Click()
DoVer3

End Sub

Private Sub Form_Load()

FECHA.Value = Date
FECHA.Day = 1

Me.Move 0, 800, 11100, 4800
Label2.Caption = CodCLTE & " - " & ClteNomb

DoVer3

End Sub

Private Sub Form_Unload(Cancel As Integer)
'FCMENU.Toolbar2.Buttons(7).Value = tbrUnpressed

End Sub

Private Sub FECHA_Change()

DoVer3

End Sub

Sub DoVer3()
Dim TIPO As Byte

  If ORDENADOx <> "" Then
         Clasifica = elORDEN & ORDENADOx & ORDENADOad
  Else
         Clasifica = ""
End If
'///           ----------------------------------------

laLOGIK1 = Format(FECHA.Month, "00") & "/" & Format(FECHA.Day, "00") & "/" & FECHA.Year
miSQL = "SELECT TOP " & Combo1.Text & " * FROM FCEstad1 WHERE CLTE = " & CodCLTE & _
        " and FECHA >= " & Separa & laLOGIK1 & Separa
RgEST1.Open miSQL, dbTABL, adOpenForwardOnly, adLockReadOnly
   
If RgEST1.EOF Then
    RgEST1.Close
    Exit Sub
End If

' paraSTATUS = "Seleccionados : " & RgEST1.RecordCount
'FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

       LimpiaGRILLA
       CargaGrilla

End Sub

Sub LimpiaGRILLA()
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

FG3.AddItem FECHA & vbTab & TipoCpbt & " " & Letra & " " & PtoVta & "-" & Cpbte & _
            vbTab & Secc & vbTab & Pulg & vbTab & Mtr & vbTab & MM & vbTab & _
            Telas & vbTab & Cant & vbTab & PVta & vbTab & Impte, FG3.Rows

CanReg = CanReg + 1

paraSTATUS = "Cargando : " & (FG3.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

Me.Show
DoEvents

SALTA:

RgEST1.MoveNext
Loop

    RgEST1.Close

FG3.AutoSize 0, 9

Label5.Caption = Format(CanReg, "##,##0 ")
Label7.Caption = Format(Total, "$ ###,##0.00 ")

If FG3.Rows > 1 Then FG3.Row = 1

' FG1.SetFocus

End Sub
