VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form BusClte 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Busqueda de Clientes   "
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8160
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   37
      Top             =   15
      Visible         =   0   'False
      Width           =   7995
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cliente Ocacional"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   465
         Left            =   2070
         TabIndex        =   38
         Top             =   270
         Width           =   4185
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cliente Ocasional"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5070
      TabIndex        =   36
      Top             =   3150
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Alta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4335
      TabIndex        =   18
      Top             =   3690
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
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
      Height          =   345
      Left            =   6795
      TabIndex        =   17
      Top             =   3690
      Width           =   1155
   End
   Begin VB.ListBox List3 
      BackColor       =   &H00FEFADE&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      ItemData        =   "BusClte.frx":0000
      Left            =   4590
      List            =   "BusClte.frx":0002
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   16
      ToolTipText     =   "Condición de Venta"
      Top             =   4455
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Frame Espera 
      BackColor       =   &H00FCE7C7&
      Caption         =   "   Carga de Datos   "
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
      Height          =   2325
      Left            =   1440
      TabIndex        =   13
      Top             =   1170
      Visible         =   0   'False
      Width           =   5235
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Top             =   930
         Width           =   4755
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
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
      Height          =   345
      Left            =   5580
      TabIndex        =   3
      Top             =   3690
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "    Busqueda de Clientes   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7995
      Begin VB.OptionButton Option1 
         Caption         =   "Vendedor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   690
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Localidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   690
         Width           =   1155
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   2790
         ScaleHeight     =   435
         ScaleWidth      =   4935
         TabIndex        =   12
         Top             =   270
         Width           =   4995
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            MaxLength       =   11
            TabIndex        =   0
            Top             =   60
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   1
            Top             =   60
            Width           =   4425
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   450
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Razón Soc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   450
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "CUIT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1230
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Datos del Cliente  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2640
      Left            =   0
      TabIndex        =   19
      Top             =   990
      Visible         =   0   'False
      Width           =   7995
      Begin VB.CommandButton Command5 
         Caption         =   "Continuar"
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
         Left            =   6180
         TabIndex        =   28
         Top             =   2100
         Width           =   1410
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FEFADE&
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
         ItemData        =   "BusClte.frx":0004
         Left            =   1680
         List            =   "BusClte.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   27
         ToolTipText     =   "Seleccione la Condición de IVA"
         Top             =   1485
         Width           =   2130
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FEFADE&
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
         Left            =   5685
         LinkTimeout     =   5
         MaxLength       =   2
         TabIndex        =   26
         Text            =   "B"
         ToolTipText     =   "Código de Provincia"
         Top             =   1215
         Width           =   525
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FEFADE&
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   25
         ToolTipText     =   "Razón Social del Cliente"
         Top             =   390
         Width           =   4530
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   24
         ToolTipText     =   "Dirección del Cliente"
         Top             =   675
         Width           =   4530
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         LinkTimeout     =   5
         MaxLength       =   30
         TabIndex        =   23
         ToolTipText     =   "Localidad del Cliente"
         Top             =   930
         Width           =   2820
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         LinkTimeout     =   5
         MaxLength       =   10
         TabIndex        =   22
         ToolTipText     =   "Código Postal"
         Top             =   1215
         Width           =   1020
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FEFADE&
         Enabled         =   0   'False
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
         Left            =   5685
         MaxLength       =   11
         TabIndex        =   21
         ToolTipText     =   "Número CUIT/CUIL/DNI (sin guiones)"
         Top             =   1530
         Width           =   1935
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4830
         TabIndex        =   20
         ToolTipText     =   "Vuelve a la Busqueda"
         Top             =   2100
         Width           =   1155
      End
      Begin VB.Label Label107a 
         Caption         =   "Provincia          :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4380
         TabIndex        =   35
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label Label102 
         Caption         =   "Razón Social :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   34
         Top             =   405
         Width           =   1395
      End
      Begin VB.Label Label104 
         Caption         =   "Dirección           :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label105 
         Caption         =   "Localidad           :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   32
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label106 
         Caption         =   "Cod. Postal        :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   31
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label Label109 
         Caption         =   "Condic. I.V.A.    :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   30
         Top             =   1575
         Width           =   1335
      End
      Begin VB.Label Label110 
         Caption         =   "C.U.I.T.              :"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4380
         TabIndex        =   29
         Top             =   1575
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2580
      Left            =   0
      TabIndex        =   10
      Top             =   1035
      Visible         =   0   'False
      Width           =   8010
      Begin VSFlex8Ctl.VSFlexGrid FG1 
         Height          =   2385
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   7890
         _cx             =   13917
         _cy             =   4207
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
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   17
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
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   45
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   4185
   End
End
Attribute VB_Name = "BusClte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LaRow As Integer, ParaSQL As String
Private Const laselec20 = "SELECT * FROM Clientes WHERE "
Dim IND As Integer
Dim YaEsta As Boolean

Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset

Private Sub Form_Activate()

unload CtaCte

End Sub

Private Sub Form_Load()
Dim I1%, HAY As String

HAY = ""
LaRow = 1
VerDetClt = False

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

Me.Move 0, 1610, 8300, 4500
'FCMENU.Toolbar2.Buttons(1).Value = tbrPressed

Text1.Visible = False
Text2.Visible = False

    ' Tabla de Condiciones de Venta
    
miSQL2 = "SELECT cod,descri FROM FCTabla1 WHERE CTAB = 'CV   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

Do Until RgTABL.EOF
List3.AddItem RgTABL!COD & "-" & RgTABL!Descri
RgTABL.MoveNext
Loop

RgTABL.Close

miSQL3 = "SELECT * FROM Clientes WHERE "

    
If HAY = " AND " Then DoVer

Option1(0).Value = False
Option1(0).Value = True

Select Case DeDonde
  
  Case 1
'    Command4.Visible = True
    CabFact.Picture2.Visible = True
  
  Case 2
'    Command4.Visible = True
'    CabRec.Picture2.Visible = True

End Select


End Sub

Private Sub Form_Unload(Cancel As Integer)

dbCLTE.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 39 Then SendKeys "{TAB}"
If KeyCode = 38 Or KeyCode = 37 Then SendKeys "+{TAB}"

If KeyCode = vbKeyF5 Then
    CabFact.Option1(1).Value = True
        End If
If KeyCode = vbKeyF6 Then
    CabFact.Option1(2).Value = True
        End If

If KeyCode = vbKeyF7 Then
    CabFact.Option1(3).Value = True
    End If

If KeyCode = vbKeyF12 Then
    CabFact.Option1(4).Value = True
        End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Then
    If Frame1.Visible Then
        Command2_Click
    Else
        Command6_Click
    End If
End If

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Command2_Click()

'Unload Me

Call SlideWindow(Me, 50)

If CodCLTE = 0 Then
    Select Case DeDonde
        Case 1
            unload CabFact
            unload DetFact
            unload PieFact
        Case 2
            unload CabRec
            unload DetRec
            unload DetPago
            unload PieRec
        Case 7
            unload NCInterna
    End Select
End If

End Sub

Private Sub Command3_Click()

TipoMov3 = 1

ABMClte.Show

End Sub

Private Sub Command4_Click()


Frame1.Visible = False
Frame3.Visible = True
Frame4.Visible = True

Text4.Text = "": Text5.Text = "": Text6.Text = "": Text7.Text = ""
Text8.Text = "B": Text9.Text = ""
Combo1.ListIndex = 2

Command1.Visible = False
Command3.Visible = False
Command4.Visible = False

Text4.SetFocus

End Sub

Private Sub Command5_Click()

CodCLTE = 0

SaleBien

End Sub

Private Sub Command6_Click()

Frame1.Visible = True
Frame3.Visible = False
Frame4.Visible = False

Command1.Visible = True
Command3.Visible = True
    
End Sub

Private Sub FG1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And YaEsta Then SaleBien

End Sub

Private Sub FG1_Click()

CodCLTE = FG1.TextMatrix(FG1.Row, 0)

SaleBien

End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then SaleBien

End Sub

Private Sub FG1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Final

If FG1.MouseRow < 1 Then Exit Sub

If FG1.MouseRow > 0 Then
       FG1.ToolTipText = FG1.TextMatrix(FG1.MouseRow, 1)
    End If

If FG1.MouseRow = FG1.Row Then
   FG1.Cell(7, LaRow, 0, LaRow, 10) = &H80000012
   FG1.BackColorSel = &HFF0000
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 10) = &HFCE7C7
   LaRow = FG1.MouseRow
   Else
   FG1.BackColorSel = &H800000
   End If

If FG1.MouseRow > 0 And FG1.MouseRow < FG1.Rows Then
   FG1.Cell(7, LaRow, 0, LaRow, 10) = &H80000012
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 10) = &HFF0000
   LaRow = FG1.MouseRow
   End If

Final:

End Sub

Private Sub Option1_Click(Index As Integer)

Text1.Text = "": Text2.Text = ""

Text1.Visible = False
Text2.Visible = False

Select Case Index


        Case 0
                ParaSQL = laselec20 & "CODIGO = "
 
        Case 1
                ParaSQL = laselec20 & "CUIT like '%"
       
        Case 2
                ParaSQL = laselec20 & "NOMB  LIKE '%"
        
        Case 3
                ParaSQL = laselec20 & "PCIA  LIKE '%"
        
        Case 4
                ParaSQL = laselec20 & "LOC  LIKE '%"
        
        Case 5
                ParaSQL = laselec20 & "VEND  = "

        
End Select
        
IND = Index
        
If (Index > 1 And Index < 5) Then
                Text1.Visible = True
    '            Text1.SetFocus
    Else
                Text2.Visible = True
    '            Text2.SetFocus
                End If
                
 If Text1.Visible = True Then Text1.SetFocus
 If Text2.Visible = True Then Text2.SetFocus
                
End Sub

Sub Command1_Click()
Dim CUIT1 As Double, CUIT2 As Double

If Text1.Text = "" And Text2.Text = "" Then Exit Sub

If IND = 1 Then
    CUIT1 = Val(Text2.Text)
'    CUIT2 = Val(Text2.Text) + 1
'    CUIT1 = 11 - Len(Text2.Text)
'    CUIT2 = CUIT2 * 10 ^ CUIT1
'    CUIT1 = Val(Text2.Text) * 10 ^ CUIT1
    miSQL3 = ParaSQL & CUIT1 & "%'"
Else
    If (IND > 1 And IND <> 5) Then
        miSQL3 = ParaSQL & Text1.Text & "%'"
    Else
        miSQL3 = ParaSQL & Text2.Text
    End If
End If

If IND = 0 Then
    CodCLTE = Val(Text2.Text)
    SaleBien
    Exit Sub
End If

LaRow = 1
Command1.Default = False

DoVer

End Sub

Private Sub text1_Change()

Command1.Enabled = True

If DeDonde = 1 Then
    If Len(Text1.Text) > 0 Then
'        Command1.Default = True
 '       Command4.Visible = False
    Else
'        Command4.Visible = True
'        Command4.Default = True
    End If
End If
    
Command1_Click

End Sub

Private Sub text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Command1.Default = True

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub
If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If
    
End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)
Command1.Default = True

End Sub

Private Sub Text2_Change()

Command1.Enabled = True

If DeDonde = 1 Then
    If Len(Text1.Text) > 0 Then
'        Command1.Default = True
'        Command4.Visible = False
    Else
'        Command4.Visible = True
'        Command4.Default = True
    End If
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text9_lostfocus()

If Not Cuit(Text9.Text) Then
        vbMsgBoxText = " Error en el CUIT " & vbCrLf & vbCrLf & "Reingresa ?"
        vbMsgBoxResp = MsgBox(vbMsgBoxText, vbExclamation + vbYesNo, " Error de Ingreso ")
        If vbMsgBoxResp = vbYes Then
            Text9.SetFocus
            Exit Sub
        Else
            Combo1.SetFocus
        End If
Else
    Command5.SetFocus
End If

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub
If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Combo1_Click()

If Combo1.ListIndex < 2 Then
    Label110.Enabled = True
    Text9.Enabled = True
    Text9.SetFocus
Else
    Label110.Enabled = False
    Text9.Text = ""
    Text9.Enabled = False
    Command5.SetFocus
End If

End Sub

Private Sub DoVer()
Dim CanClte As Integer

Frame2.Visible = False
Label2.Visible = False
    
    Espera.BackColor = &HFCE7C7
    Espera.ForeColor = &HC00000
    Label1.ForeColor = &HC00000
    Espera.Caption = "   Carga de Datos  "
    Label1 = "Buscando Datos, Por Favor Espere ..."
    
Espera.Visible = True

'Show
'DoEvents

RgCLTE.Open miSQL3, dbCLTE, adOpenForwardOnly, adLockReadOnly

' CanClte = RgCLTE.RecordCount
Label2.Caption = "Clientes Encontrados :" & CanClte
Label2.Visible = True

If RgCLTE.EOF Then
    Espera.Visible = True
    Espera.BackColor = &HC0&
    Espera.ForeColor = &HFFFF&
    Label1.ForeColor = &HFFFF&
    Espera.Caption = "   Error de Busqueda  "
    Label1.Caption = " No existen datos para esta Selección "
    RgCLTE.Close
    Exit Sub
    End If

       LimpiaGrilla

Do Until RgCLTE.EOF

FG1.AddItem RgCLTE!CODIGO & vbTab & RgCLTE!Nomb & vbTab & RgCLTE!Dir & vbTab & _
            RgCLTE!Loc & vbTab & RgCLTE!PCIA & vbTab & RgCLTE!tel1 & vbTab & _
            RgCLTE!CIVA & vbTab & RgCLTE!Cuit, FG1.Row + 1


RgCLTE.MoveNext

Loop

RgCLTE.Close
       
FG1.AutoSize 0, 5
       
Espera.Visible = False
Frame2.Visible = True
YaEsta = True
FG1.Row = 1

'FG1.SetFocus

End Sub

Private Sub SaleBien()
Dim LaConstAFIP As String, DescCVTA As String, ClteCanal As Integer

If CodCLTE = 0 Then
    ClteNomb = Text4.Text
    ClteDir = Text5.Text
    ClteLoc = Text6.Text
    ClteCPos = Text7.Text
    CltePCIA = Text8.Text
    ClteCIVA = Combo1.ListIndex + 1
    ClteCUIT = Text9.Text
    ClteCVta = 4
    ClteCanal = 0
    
Else
    
    miSQL2 = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE
    RgCLTE.Open miSQL2, dbCLTE, adOpenForwardOnly, adLockReadOnly
    If RgCLTE.EOF Then
        RgCLTE.Close
        Exit Sub
    End If
    
    ClteNomb = RgCLTE!Nomb
    ClteDir = RgCLTE!Dir
    ClteLoc = RgCLTE!Loc
    ClteCPos = RgCLTE!CP
    ClteTel1 = RgCLTE!tel1
    CltePCIA = RgCLTE!PCIA
    ClteCIVA = RgCLTE!CIVA
    ClteCUIT = RgCLTE!Cuit
    ClteCVta = RgCLTE!CVTA
    ClteVend = RgCLTE!VEND
    ClteCanal = RgCLTE!CANAL
    ClteDeuda = RgCLTE!DEUDA
    ClteCredit = RgCLTE!CREDIT
        
        RgCLTE.Close

End If
    
Select Case DeDonde
    Case 0
        TipoMov3 = 3
        VerDetClt = True
        ABMClte.Caption = "Detalle del Cliente "
        ABMClte.Frame1.Enabled = False
        ABMClte.Frame2.Enabled = False
        ABMClte.Frame3.Enabled = False
        ABMClte.Frame5.Enabled = False
'        ABMClte.Command2.Visible = False
'        ABMClte.Command4.Visible = False
'        ABMClte.Command5.Visible = False
           ABMClte.Show
           
    Case 1
'  **********************************************************************
'             Para la Factura
'  **********************************************************************
        BuscaEnCombo CabFact.Combo2, (String(3 - Len(ClteVend), "0") & ClteVend)

        CabFact.Label4.Caption = "Constancia Vigente"
        CabFact.Label4.BackColor = vbBlue
        
        If ClteCIVA > 2 And ClteCIVA < 6 Then
            CabFact.Label4.Caption = "No requiere Constancia"
            CabFact.Label4.BackColor = vbBlack
        End If
        
        If ClteCIVA < 3 Then
            miSQL2 = "SELECT * FROM Constancias WHERE CUIT = " & Val(ClteCUIT)
            RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
            
            LaConstAFIP = ""
            
            If RgTABL.EOF Then
                LaConstAFIP = "No tiene Constancia Cargada"
                    CabFact.Label4.Caption = "Falta Constancia"
                    CabFact.Label4.BackColor = vbRed
            Else
                If RgTABL("FecVto") < Date Then
                    LaConstAFIP = "Tiene la Constancia Vencida"
                        CabFact.Label4.Caption = "Constancia Vencida"
                        CabFact.Label4.BackColor = vbRed
                End If
            End If
            
            RgTABL.Close
            
            If LaConstAFIP <> "" Then
                vbMsgBoxText = "           Este Cliente " & vbCrLf & LaConstAFIP & vbCrLf & vbCrLf & "Continúa ?"
                vbMsgBoxResp = MsgBox(vbMsgBoxText, vbExclamation + vbYesNo, "  Constancia de AFIP  ")
                If vbMsgBoxResp = vbNo Then
                    If Text1.Visible Then Text1.SetFocus
                    Exit Sub
                End If
            End If
            
        End If
         
        CabFact.Label101 = "( " & CodCLTE & " )  " & ClteNomb
        CodCLTEf = CodCLTE
        CabFact.Label102 = ClteDir & " ( " & ClteCPos & " )  " & ClteLoc & " (" & CltePCIA & ")"
        CabFact.Label103 = "IVA " & CabFact.Combo1.List(ClteCIVA - 1) & "     CUIT : " & ClteCUIT
        CabFact.Label5 = "Tel: " & ClteTel1
        
            
'        If CabFact.Combo3.ListCount >= (ClteCVta - 1) Then
                CabFact.Combo3.ListIndex = BuscaEnCombo(CabFact.Combo3, ClteCVta)
        '        Combo3.ListIndex = RgCLTE!CVTA - 1
                DescCVTA = Mid(CabFact.Combo3.List(BuscaEnCombo(CabFact.Combo3, ClteCVta)), 4, 40)
'        Else
'                DescCVTA = "Cond. Vta."
'               End If

        CabFact.Label104 = "Condición de Venta : " & ClteCVta & " - " & DescCVTA
        
        CabFact.Picture2.Visible = False
        If TipoFac = 1 Or TipoFac = 4 Then
            DetFact.FG1.TabIndex = 0
        End If
        If TipoFac = 2 Or TipoFac = 3 Then
            CabFact.Combo5.TabIndex = 0
        End If
        
        If TieneDeuda Then
            ConDeuda.Show 1
            If ClteCanal <> 9 Then Exit Sub
        End If
                
        PieFact.Label4.Caption = vbCrLf & "Saldo: " & Format(SaldoClte, "$###,##0.00")

        CabFact.BuscaUltima
        
        unload Me
        
        If Not NOTACLTE.Visible Then
            DetFact.Enabled = True
            DetFact.FG1.SetFocus
        End If
                
        VeNotaC = False

    Case 2
'  **********************************************************************
'             Para el Recibo
'  **********************************************************************
        DetPago.Show
        DetRec.Show
        PieRec.Show

    ' Busca Próximo Nro. de Recibo
    miSQL1 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
    RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockOptimistic

CabRec.Text1.Text = 0

If Not RgTABL.EOF Then
    CabRec.Text1.Text = Format((RgTABL!nume8 + 1), "###,##0")
End If

RgTABL.Close
         
        CabRec.Label101 = "( " & CodCLTE & " )  " & ClteNomb
'        CodCLTEf = CodCLTE
        CabRec.Label102 = ClteDir & " ( " & ClteCPos & " )  " & ClteLoc & " (" & CltePCIA & ")"
        CabRec.Label103 = "IVA " & CabRec.Combo1.List(ClteCIVA - 1) & "     CUIT : " & ClteCUIT
'       CabRec.Label5 = "Tel: " & ClteTel1
'
        CabRec.Label3 = Format(ClteCredit, "###,##0.00")
        CabRec.Label5 = Format(ClteDeuda, "###,##0.00")
        
        If Not NOTACLTE.Visible Then
            DetRec.Enabled = True
            DetRec.FG1.SetFocus
        End If
                
        VeNotaC = False
        unload Me
    
    Case 3
        CtaCte.Show
        CtaCte.SetFocus
        
    Case 4
        Listados.Label2.Caption = "Cliente : " & CodCLTE & "  " & ClteNomb
        unload Me
    
    Case 5
        DtosxClte.Show
        
    Case 6
        CpraClte.Show
    Case 7
        NCInterna.Label2 = "( " & CodCLTE & " )  " & ClteNomb
        NCInterna.Command1.Enabled = False
        NCInterna.BuscaDeuda
        NCInterna.Command3.Visible = True
        unload Me
        
End Select

' Unload Me

End Sub

Sub LimpiaGrilla()
Dim I1%
   
LaRow = 0
FG1.Clear flexClearScrollable

For i = (FG1.Rows - 1) To 1 Step -1
FG1.RemoveItem (i)
Next i

End Sub

Private Function TieneDeuda() As Boolean
Dim Dif As Integer
     
SaldoClte = 0

miSQL = "SELECT * FROM CtasCtes WHERE CLTE = " & CodCLTE & _
        " ORDER BY FECHA"
RgTABL.Open miSQL, dbTABL, adOpenForwardOnly, adLockReadOnly

If RgTABL.EOF Then
    RgTABL.Close
    GoTo aDeuda
    End If

Do Until RgTABL.EOF


If RgTABL!TIPO = 0 Or RgTABL!TIPO = 1 Or RgTABL!TIPO = 3 Or RgTABL!TIPO = 7 Then
    SaldoClte = SaldoClte + RgTABL!Impte
Else
    SaldoClte = SaldoClte - RgTABL!Impte
End If

RgTABL.MoveNext

Loop
    
RgTABL.Close

'-------------------------------------------------------------
aDeuda:

miSQL = "SELECT * FROM CtasCtes WHERE CLTE = " & CodCLTE & _
        " and TIPO = 1 and DEBE > 1 ORDER BY FECHA, CPBTE, TIPO"
RgTABL.Open miSQL, dbTABL, adOpenForwardOnly, adLockReadOnly

If RgTABL.EOF Then
    RgTABL.Close
    Exit Function
    End If
'RgTABL!DEBE = 0
Dif = Now - RgTABL!FECHA
RgTABL.Close

If Dif < DiasDeuda Then Exit Function

TieneDeuda = True

End Function
