VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form DetPago 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso del Pago"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   3210
   ClientWidth     =   7350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0000C000&
      Height          =   3420
      Left            =   1725
      ScaleHeight     =   3360
      ScaleWidth      =   5025
      TabIndex        =   19
      Top             =   -45
      Visible         =   0   'False
      Width           =   5085
      Begin VB.CommandButton Command4 
         Caption         =   "Carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2910
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3015
         Width           =   930
      End
      Begin VB.CommandButton Command2 
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
         Height          =   315
         Left            =   3945
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3015
         Width           =   930
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "  Detalle del Cheque  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2910
         Left            =   45
         TabIndex        =   20
         Top             =   60
         Visible         =   0   'False
         Width           =   4920
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "DetPago.frx":0000
            Left            =   3270
            List            =   "DetPago.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Tipo de Cheque (Papel/Electr.)"
            Top             =   195
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   14
            TabIndex        =   9
            ToolTipText     =   "Observaciones"
            Top             =   2565
            Width           =   3615
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   8
            ToolTipText     =   "Concepto"
            Top             =   2280
            Width           =   3615
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Caption         =   "Orden   :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   135
            TabIndex        =   6
            ToolTipText     =   "Si es ""A la Orden"""
            Top             =   1680
            Visible         =   0   'False
            Width           =   1185
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            ToolTipText     =   "Fecha de Emisión"
            Top             =   795
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   503
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
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   160104451
            CurrentDate     =   38281
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   15
            TabIndex        =   7
            ToolTipText     =   "Datos Adicionales del Emisor"
            Top             =   1995
            Width           =   3615
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            ToolTipText     =   "Importe"
            Top             =   1380
            Width           =   1440
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   2
            ToolTipText     =   "Código del Banco"
            Top             =   510
            Visible         =   0   'False
            Width           =   630
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   1080
            TabIndex        =   4
            ToolTipText     =   "Fecha de Vencimiento"
            Top             =   1080
            Visible         =   0   'False
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   503
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
            CustomFormat    =   "ddd dd MMM yyyy"
            Format          =   160104451
            CurrentDate     =   38281
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   12
            TabIndex        =   0
            ToolTipText     =   "Número del Cheque"
            Top             =   210
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "DetPago.frx":002A
            Left            =   1080
            List            =   "DetPago.frx":0043
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   195
            Visible         =   0   'False
            Width           =   3645
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2430
            TabIndex        =   32
            Top             =   240
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Observ. :"
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
            Left            =   135
            TabIndex        =   29
            Top             =   2580
            Width           =   930
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto:"
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
            Left            =   135
            TabIndex        =   28
            Top             =   2295
            Width           =   930
         End
         Begin VB.Label Label8 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Datos Ad:"
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
            Left            =   135
            TabIndex        =   27
            Top             =   2010
            Width           =   930
         End
         Begin VB.Label Label7 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe:"
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
            Left            =   135
            TabIndex        =   26
            Top             =   1395
            Width           =   930
         End
         Begin VB.Label Label6 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fec.Vto.:"
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
            Left            =   135
            TabIndex        =   25
            Top             =   1095
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFC0&
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
            Left            =   1800
            TabIndex        =   24
            Top             =   540
            Width           =   2955
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fec.Emi.:"
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
            Left            =   135
            TabIndex        =   23
            Top             =   810
            Width           =   930
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Banco   :"
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
            Left            =   135
            TabIndex        =   22
            Top             =   525
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Operac.:"
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
            Left            =   135
            TabIndex        =   31
            Top             =   255
            Width           =   930
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Número:"
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
            Left            =   135
            TabIndex        =   21
            Top             =   225
            Visible         =   0   'False
            Width           =   930
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2820
      Left            =   0
      TabIndex        =   16
      Top             =   -60
      Width           =   6780
      Begin VSFlex8Ctl.VSFlexGrid FG1 
         Height          =   2625
         Left            =   30
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   6690
         _cx             =   11800
         _cy             =   4630
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
         ForeColorSel    =   15335153
         BackColorBkg    =   12632064
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"DetPago.frx":00CB
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5415
         TabIndex        =   18
         Top             =   2970
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   0
      TabIndex        =   17
      Top             =   2715
      Width           =   6750
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
         Left            =   1500
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   195
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Efectivo"
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
         Left            =   180
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
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
         Left            =   5355
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
   End
End
Attribute VB_Name = "DetPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LaRow As Integer, Col0Ant As String, Col10Ant As Integer
Dim ValorCelda As String, TotPago As Double
Dim LaTecla As Byte
Dim dbCYB As New ADODB.Connection, RgCYB As New ADODB.Recordset
'-------------------------------------------------
'    Grilla Valores
'
'  0-Nro.Cheque                6-Orden
'  1-Nomb.Bco./Descr.          7-Datos Adic.
'  2-Fecha Vto.                8-Concepto
'  3-Importe                   9-Observac.
'  4-Cod.Banco/Suc.           10-Cód. de Anticipo
'  5-Fecha Emisión            11-
'-------------------------------------------------

Private Sub Form_Load()

Me.Move 4380, 1650, 6900, 3720

dbCYB.ConnectionString = BDatos1
dbCYB.Open

LaRow = 1

Combo1.List(0) = "Retenc.Gan."
Combo1.List(1) = "Retenc.IB"
Combo1.List(2) = "Retenc.IVA"
Combo1.List(3) = "Retenc.SUS"
Combo1.List(4) = "Dep./Transf."
Combo1.List(5) = "Tarj.Déb/Cr."
Combo1.List(6) = "Cheques Vs."
Combo1.List(7) = "Cheq.Electr."
Combo1.List(8) = "Baja Incobr."

Combo2.List(0) = "Común (Papel)"
Combo2.List(1) = "Electrónico"

DTPicker1.Value = Date  ' Fec.Emision
DTPicker2.Value = Date  ' Fec.Vto.

End Sub

Private Sub Form_GotFocus()
                   
If Picture1.Visible = False Then
    FG1.SetFocus
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

PieRec.Label303.Caption = 0
PieRec.Label304.Caption = PieRec.Label302.Caption

dbCYB.Close

End Sub

Private Sub Form_Activate()

If NroRec = 0 Then
    CabRec.Show
End If

If FG1.Rows > 1 Then
    FG1.Row = 1
End If

End Sub

Private Sub Form_Astivate()
 
 FG1.BackColorFixed = &HE0E0E0
 FG1.ForeColorFixed = &HC00000
                   
If Picture1.Visible = False Then
    NroCheq = 0
    FG1.SetFocus
End If

If FG1.Rows <= 2 Then Exit Sub
If FG1.Row < 1 Then FG1.Row = 1

End Sub

Private Sub Form_Deactivate()
 
 FG1.BackColorFixed = &H8000000F
 FG1.ForeColorFixed = &H80000012

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'If KeyCode = 40 Then SendKeys "{TAB}"
'If KeyCode = 38 Then SendKeys "+{TAB}"
If KeyCode = 27 Then Unload Me

If KeyCode = vbKeyF9 And CCur(PieRec.Label304.Caption) = 0 Then
                    If FG1.Rows > 2 Then
                                FG1.TextMatrix(FG1.Row, 0) = ""
                               EmiRec.Show
                    End If
        End If

If KeyCode = vbKeyTab Then
    DetRec.Show
    DetRec.SetFocus
    Exit Sub
        End If
 
If KeyCode = vbKeyReturn Then
    SendKeys "{TAB}"
    Exit Sub
End If

If KeyCode = vbKeyF5 Then
    Command1_Click
    Exit Sub
        End If

End Sub

Private Sub Command1_Click()

If Text1.Visible Then
    Text1.Text = ""
    Text1.Visible = False
Else
    Text1.Text = 0
    Text1.Visible = True
    Text1.SetFocus
        End If
    
End Sub

Private Sub Command3_Click()

Unload Me
Importe = 0
Text1.Text = ""

End Sub

Private Sub Command2_Click()

Frame1.Enabled = True: Frame2.Enabled = True

NroCheq = Col0Ant
FG1.TextMatrix(FG1.Row, 10) = Col10Ant

Picture1.Visible = False
FG1.SetFocus
' FG1.EditCell

End Sub

Private Sub Command4_Click()

FG1.TextMatrix(FG1.Row, 0) = Text2.Text    'Nro.Cheque

FG1.TextMatrix(FG1.Row, 4) = Text3.Text    'Banco
FG1.TextMatrix(FG1.Row, 1) = Label5.Caption
FG1.TextMatrix(FG1.Row, 11) = Label5.Caption

FG1.TextMatrix(FG1.Row, 3) = Text4.Text    ' Importe

FG1.TextMatrix(FG1.Row, 6) = False
If Check1.Value = 1 Then                   ' Orden
    FG1.TextMatrix(FG1.Row, 6) = True
End If

FG1.TextMatrix(FG1.Row, 7) = Text5.Text    ' Datos Adic.
FG1.TextMatrix(FG1.Row, 8) = Text6.Text    ' Concepto
FG1.TextMatrix(FG1.Row, 9) = Text7.Text    ' Observac.
FG1.TextMatrix(FG1.Row, 12) = Combo2.ListIndex    ' Tipo de Cheque.

FG1.TextMatrix(FG1.Row, 5) = DTPicker1.Value ' Fec.Emision
FG1.TextMatrix(FG1.Row, 2) = DTPicker2.Value ' Fec.Vto.

Frame1.Enabled = True: Frame2.Enabled = True
Picture1.Visible = False

VerSiAgrega

End Sub

Private Sub Combo2_Click()
 If Text3.Visible Then Text3.SetFocus
End Sub

Private Sub Combo1_Click()

FG1.TextMatrix(FG1.Row, 1) = Combo1.List(Combo1.ListIndex)
Label5.Caption = Combo1.List(Combo1.ListIndex)

Select Case Combo1.ListIndex
    Case 0
        FG1.TextMatrix(FG1.Row, 0) = "Retenc.Gan."
        Text2.Text = "Retenc.Gan."
    Case 1
        FG1.TextMatrix(FG1.Row, 0) = "Retenc.IB"
        Text2.Text = "Retenc.IB"
    Case 2
        FG1.TextMatrix(FG1.Row, 0) = "Retenc.IVA"
        Text2.Text = "Retenc.IVA"
    Case 3
        FG1.TextMatrix(FG1.Row, 0) = "Retenc.SUS"
        Text2.Text = "Retenc.SUS"
    Case 4
        FG1.TextMatrix(FG1.Row, 0) = "Dep./Transf."
        Text2.Text = "Dep./Transf."
    Case 5
        FG1.TextMatrix(FG1.Row, 0) = "Tarj.Déb/Cr."
        Text2.Text = "Tarj.Déb/Cr."
    Case 6
        FG1.TextMatrix(FG1.Row, 0) = "Cheques Vs."
        Text2.Text = "Cheques Vs."
    Case 7
        FG1.TextMatrix(FG1.Row, 0) = "Cheq.Electr."
        Text2.Text = "Cheq.Electr."
    Case 8
        FG1.TextMatrix(FG1.Row, 0) = "Baja Incobr."
        Text2.Text = "Baja Incobr."
End Select


FG1.TextMatrix(FG1.Row, 10) = Combo1.ListIndex + 1
FG1.TextMatrix(FG1.Row, 12) = 0
    
    If Combo1.ListIndex = 5 Then
        Text3.Visible = True       '   Si es Tarj.
        Label3.Visible = True       '   Pide Banco
        Text3.SetFocus
    Else
        Text3.Visible = False
        Label3.Visible = False
        Label5.Visible = False
        If Combo1.ListIndex < 5 Then
            Text4.SetFocus
            DTPicker1.Enabled = False
        Else
            DTPicker1.Enabled = True
            DTPicker1.SetFocus
        End If
    End If

End Sub

Private Sub FG1_GotFocus()
            
                    If FG1.FindRow(NroCheq, , 0) > -1 Then
                       FG1.Row = FG1.FindRow(NroCheq, , 0)
'                       DetalleCheque
'                        Exit Sub
                    Else
                    
                        FG1.TextMatrix(FG1.Row, 0) = 0
                        FG1.TextMatrix(FG1.Row, 1) = ""
                        FG1.TextMatrix(FG1.Row, 2) = Date
                        FG1.TextMatrix(FG1.Row, 5) = Date
                        FG1.TextMatrix(FG1.Row, 3) = 0

                        End If
    
    FG1.Col = 0
    FG1.EditCell

End Sub

Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
 
 If Not (LaTecla = vbKeyUp Or LaTecla = vbKeyDown) Then
        If Col = 0 And Val(FG1.TextMatrix(FG1.Row, 0)) <> 0 Then
            Col0Ant = FG1.TextMatrix(FG1.Row, 0)
            Col10Ant = Val(FG1.TextMatrix(FG1.Row, 10))
            DetalleCheque
            End If
End If

End Sub

Private Sub FG1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

If KeyCode = vbKeyTab Then
    DetRec.Show
    DetRec.SetFocus
    Exit Sub
        End If

If KeyCode = vbKeyF8 Then
        Col0Ant = Val(FG1.TextMatrix(FG1.Row, 0))
        Col10Ant = Val(FG1.TextMatrix(FG1.Row, 10))
        If Val(FG1.TextMatrix(FG1.Row, 10)) = 0 Then
            FG1.TextMatrix(FG1.Row, 0) = "Retenc./Antic."
            FG1.TextMatrix(FG1.Row, 10) = 1
        End If
        DetalleCheque
End If

If Col = 0 Then LaTecla = KeyCode

If KeyCode = vbKeyDelete And FG1.Rows > 1 Then
    FG1.RemoveItem (Row)
    VerSiAgrega
End If

End Sub

Private Sub FG1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)

If NewRow = OldRow Then Exit Sub
If Not (NewCol = OldCol) Then
    If NewCol <> 1 And NewCol <> 4 Then
        FG1.Col = NewCol
        FG1.EditCell
    End If
End If

End Sub

Private Sub FG1_Click()

If FG1.Col = 0 Then
    FG1.EditCell
    End If

End Sub

Private Sub FG1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete And FG1.Rows > 1 Then
    FG1.RemoveItem (FG1.Row)
    VerSiAgrega
End If

'If KeyCode = vbKeyTab Then
 '   DetRec.Show
'    DetRec.FG1.SetFocus
'    Exit Sub
'        End If

If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
    If FG1.Col <> 1 And FG1.Col <> 4 Then
        FG1.EditCell
        End If
    End If
    
End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyDelete Then
    FG1.RemoveItem (FG1.Row)
End If

End Sub

Private Sub FG1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo SALIR

If FG1.MouseRow = -1 Then Exit Sub
If LaRow > (FG1.Rows - 1) Then LaRow = (FG1.Rows - 1)

If FG1.MouseRow > 0 Then
    Select Case FG1.MouseCol
        Case 0
            FG1.ToolTipText = " Nº Cheque ó [F8] Anticipos/Retenciones "
        Case Is > 0
            If FG1.TextMatrix(FG1.MouseRow, 10) = 0 Then
                FG1.ToolTipText = " Datos Ad..: " & FG1.TextMatrix(FG1.MouseRow, 7) & " "
                '        FG1.ToolTipText = " -- Cheque -- "
                Else
                        FG1.ToolTipText = " Anticipo "
                        End If
        Case Else
            FG1.ToolTipText = " Observ. : " & FG1.TextMatrix(FG1.MouseRow, 9) & " "
    End Select

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

SALIR:
End Sub

Private Sub Text1_Change()

If Not IsNumeric(Text1.Text) Then Text1.Text = 0

PieRec.Label303.Caption = Format(Importe + CCur(Text1.Text), "$ ###,##0.00")

End Sub

Private Sub text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)

'If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii = vbKeyBack Then Exit Sub

If KeyAscii = vbKeyReturn Then
'    Frame1.Enabled = True
    FG1.SetFocus
    KeyAscii = 0
    Exit Sub
End If

If KeyAscii = 46 Then KeyAscii = 44

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text1_LostFocus()

If Not IsNumeric(Text1.Text) Then
    Text1.Text = 0
    End If
    
Text1.Text = Format(Val(Replace(Text1.Text, ",", ".")), "###,##0.00")

End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyBack Then Exit Sub
   
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0   ' Cancela el carácter.
'      Beep            ' Sonido de aviso de error.
   End If

End Sub

Private Sub Text3_GotFocus()

Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyBack Then Exit Sub

If KeyAscii = vbKeyReturn Then
    CargaBco
'    DTPicker1.SetFocus
    Exit Sub
    End If
   
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0   ' Cancela el carácter.
'      Beep            ' Sonido de aviso de error.
   End If

End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub

If KeyAscii = 46 Then KeyAscii = 44

If KeyAscii = Asc(",") Then Exit Sub

    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If

End Sub

Private Sub Text4_LostFocus()

If Not IsNumeric(Text4.Text) Then
    Text4.Text = 0
    End If
    
Text4.Text = Format(Val(Replace(Text4.Text, ",", ".")), "###,##0.00")

End Sub

Private Sub text4_GotFocus()

Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub text5_GotFocus()

If CCur(Text4.Text) = 0 Then Text4.SetFocus

End Sub

Private Sub Text7_LostFocus()

Command4.SetFocus

End Sub

Private Sub Check1_GotFocus()

If CCur(Text4.Text) = 0 Then Text4.SetFocus

End Sub

Sub VerSiAgrega()
Dim c1%

Importe = 0
For c1 = 1 To FG1.Rows - 1
      If FG1.TextMatrix(c1, 3) <> "" Then
          Importe = Importe + FG1.TextMatrix(c1, 3)
          End If
Next

Label1.Caption = Format(Importe, "$ ###,##0.00")
PieRec.Label303.Caption = Format(Importe + Val(Text1.Text), "$ ###,##0.00")

If Val(FG1.TextMatrix((FG1.Rows - 1), 3)) > 0 Then
                    FG1.Rows = FG1.Rows + 1
                    FG1.Row = FG1.Rows - 1

                    FG1.TextMatrix(FG1.Row, 0) = 0
                    FG1.TextMatrix(FG1.Row, 1) = ""
                    FG1.TextMatrix(FG1.Row, 2) = Date
                    FG1.TextMatrix(FG1.Row, 5) = Date
                    FG1.TextMatrix(FG1.Row, 3) = 0
                    
                    NroCheq = 0
Else
    FG1.Row = FG1.Rows - 1
                    End If
                    
                    FG1.Col = 0
                    FG1.EditCell

End Sub

Private Sub DetalleCheque()
Dim CodAntic As Single

    CodAntic = Val(FG1.TextMatrix(FG1.Row, 10))    'Codigo de Anticipo
                        
    If Val(FG1.TextMatrix(FG1.Row, 3)) = 0 Then
                        FG1.TextMatrix(FG1.Row, 1) = ""
                        FG1.TextMatrix(FG1.Row, 2) = Date
                        FG1.TextMatrix(FG1.Row, 5) = Date
                        FG1.TextMatrix(FG1.Row, 3) = 0
    End If

Picture1.Visible = True
Frame3.Visible = True
Frame1.Enabled = False: Frame2.Enabled = False

If CodAntic > 0 Then
    Picture1.BackColor = vbWhite
    Frame3.BackColor = vbWhite
    Frame3.Caption = "  Anticipos / Retenciones  "
    Text2.Visible = False
    Label2.Visible = False
    Combo2.ListIndex = 0
    
    If CodAntic = 6 Then
        Text3.Visible = True       '   Si es Tarj.
        Label3.Visible = True      '   Pide Banco
    Else                            '
        Text3.Visible = False
        Label3.Visible = False
        Label5.Visible = False
    End If
    
    DTPicker2.Visible = False
    Label6.Visible = False
    Check1.Visible = False
    
    Combo1.Visible = True
    Label11.Visible = True
    
    Combo2.Visible = False               ' Combo de tipo de cheque
    Label12.Visible = False              '  Tipo
    
    Combo1.SetFocus
    Combo1.ListIndex = CodAntic - 1
Else
    Picture1.BackColor = &HC000&
    Frame3.BackColor = &HC0FFC0
    Frame3.Caption = "  Detalle del Cheque  "
    
    Combo1.Visible = False
    Label11.Visible = False
    
    Combo2.Visible = True               ' Combo de tipo de cheque
    Label12.Visible = True              '  Tipo
    Combo2.ListIndex = 0
    
    Text2.Visible = True
    Label2.Visible = True
    Text3.Visible = True
    Label3.Visible = True
    DTPicker2.Visible = True
    Label6.Visible = True
    Check1.Visible = True

    Text2.Text = FG1.TextMatrix(FG1.Row, 0)    'Nro.Cheque
    NroCheq = FG1.TextMatrix(FG1.Row, 0)
    
    DTPicker2.Value = FG1.TextMatrix(FG1.Row, 2)  ' Fec.Vto.
    
    Check1.Value = 0
    If Val(FG1.TextMatrix(FG1.Row, 6)) = 1 Then       ' Orden
        Check1.Value = 1
    End If

    Text3.SetFocus

End If
    
    Text3.Text = FG1.TextMatrix(FG1.Row, 4)    'Banco
    Label5.Caption = FG1.TextMatrix(FG1.Row, 1)

Text4.Text = FG1.TextMatrix(FG1.Row, 3)    ' Importe

Text5.Text = FG1.TextMatrix(FG1.Row, 7)    ' Datos Adic.
Text6.Text = FG1.TextMatrix(FG1.Row, 8)    ' Concepto
Text7.Text = FG1.TextMatrix(FG1.Row, 9)    ' Observac.

DTPicker1.Value = FG1.TextMatrix(FG1.Row, 5)  ' Fec.Emision

End Sub

Private Sub CargaBco()

miSQL1 = "Select * FROM Bancos where COD = " & Val(Text3.Text)
RgCYB.Open miSQL1, dbCYB, adOpenDynamic, adLockPessimistic

    If Not RgCYB.EOF Then
         Label5.Caption = RgCYB!NOMBRE
        RgCYB.Close
        Exit Sub
        End If
         
RgCYB.Close
   
 Label5.Caption = " *** No Existe en Tabla *** "
 
  vbMsgBoxTitle = "Carga de Cheques - Banco no existe"
  vbMsgBoxText = "Desea dar de alta el Banco " & Val(Text3.Text) & " ?"
  vbMsgBoxResp = vbYesNo + vbQuestion + vbApplicationModal + vbDefaultButton2
  vbResponse = MsgBox(vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle)
 
  If vbResponse = vbYes Then
              ParaBcos = True
              ABMBcos.Show
              ABMBcos.Text101.Text = Val(Text3.Text)
              ABMBcos.Text101.Enabled = False
              Exit Sub
  End If

End Sub
