VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form ABMTablas 
   Caption         =   "ABMTablas - Version 1.0"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   9615
   Begin VB.Frame Frame1 
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
      Height          =   2955
      Left            =   1545
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   6495
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Zonas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1530
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Secciones Correa"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   690
         Value           =   -1  'True
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Datos Varios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1950
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Motivos de N/C y D"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1530
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Condic. de Venta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1110
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "U. de Medida Prod."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   690
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Provincias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1950
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendedores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1110
         Width           =   2500
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "  Actualización  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   600
      TabIndex        =   11
      Top             =   4020
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton Command5 
         Caption         =   "Cambio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2850
         TabIndex        =   14
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1560
         TabIndex        =   13
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Alta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   12
         Top             =   210
         Width           =   1100
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Menú"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6030
      TabIndex        =   10
      Top             =   4230
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   4230
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   2940
      Left            =   1935
      TabIndex        =   15
      Top             =   975
      Width           =   5775
      _cx             =   10186
      _cy             =   5186
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
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
      SelectionMode   =   0
      GridLines       =   1
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
      FormatString    =   ""
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
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tablas Varias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   1560
      TabIndex        =   9
      Top             =   240
      Width           =   6525
   End
End
Attribute VB_Name = "ABMTablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

Me.Move 0, 0, 9800, 5400

Option1(1) = False


Frame1.Visible = True
Frame2.Visible = False
FG1.Visible = False
Command1.Move 6880
Command2.Visible = False

Habilita = False

End Sub

Private Sub Command1_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Command2_Click()

FG1.Visible = False
Frame1.Visible = True
Frame2.Visible = False
Command1.Move 6880
Command2.Visible = False
Label1.Caption = "Tablas Varias"

End Sub

Private Sub Command3_Click()
CodItem = 0

TIPOMov = 1
ActTabla.Show

End Sub

Private Sub Command4_Click()

CodItem = FG1.TextMatrix(FG1.Row, 0)

TIPOMov = 2
ActTabla.Show

End Sub

Private Sub Command5_Click()
    
 CodItem = FG1.TextMatrix(FG1.Row, 0)

TIPOMov = 3
ActTabla.Show

End Sub

Private Sub Option1_Click(Index As Integer)

Select Case Index
     
     Case 1
            CodTabla = "SC "
            Label1.Caption = "Tabla de Secciones de Correas"
     Case 2
            CodTabla = "VD "
            Label1.Caption = "Tabla de Vendedores"
     Case 3
            CodTabla = "ZN "
            Label1.Caption = "Tabla de Zonas"
     Case 4
            CodTabla = "PV "
            Label1.Caption = "Tabla de Provincias"
     Case 5
            CodTabla = "UM "
            Label1.Caption = "Tabla de Unid. Medida de Prod."
     Case 6
            CodTabla = "CV "
            Label1.Caption = "Tabla de Condic. de Venta"
     Case 7
            CodTabla = "MT "
            Label1.Caption = "Tabla de Motivos de NC y D."
     Case 8
            CodTabla = "VS "
            Label1.Caption = "Tabla de Datos Varios"

End Select

LaOpcion = Index

Option1(Index) = False

CargaGrilla

FG1.Visible = True
    
If FG1.Rows > 1 Then FG1.Row = 1
    
Frame1.Visible = False
Frame2.Visible = True
Command1.Move 7580
Command2.Visible = True

End Sub

Sub CargaGrilla()

FG1.Rows = 1
FG1.Row = 0

miSQL2 = "SELECT COD, DESCRI FROM FCTabla1 WHERE CTAB = '" & CodTabla & "' " _
                                    & "ORDER BY DESCRI DESC"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

If RgTABL.EOF Then
    RgTABL.Close
    Command3_Click
    Exit Sub
    End If
    
Do Until RgTABL.EOF

FG1.AddItem RgTABL!COD & vbTab & RgTABL!Descri, FG1.Row + 1

RgTABL.MoveNext

Loop

FG1.AutoSize 0, 1
FG1.Row = 1
FG1.TabIndex = 0

RgTABL.Close


End Sub
Private Sub FG1_DBLClick()

Command5_Click

End Sub
