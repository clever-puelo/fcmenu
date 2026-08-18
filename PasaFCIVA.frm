VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PasaFCIVA 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pasa Facturas Manuales"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00F9FADC&
      Height          =   315
      ItemData        =   "PasaFCIVA.frx":0000
      Left            =   6120
      List            =   "PasaFCIVA.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   5580
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar Pasaje"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   195
      TabIndex        =   1
      Top             =   4740
      Width           =   1845
   End
   Begin VB.CommandButton Command2 
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
      Height          =   390
      Left            =   9525
      TabIndex        =   0
      Top             =   4725
      Width           =   1455
   End
   Begin VSFlex8Ctl.VSFlexGrid GRILLA2 
      Height          =   3555
      Left            =   105
      TabIndex        =   2
      Top             =   1125
      Width           =   10935
      _cx             =   19288
      _cy             =   6271
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
      FormatString    =   $"PasaFCIVA.frx":0052
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0080FF80&
         Height          =   3300
         Left            =   3090
         ScaleHeight     =   3240
         ScaleWidth      =   7680
         TabIndex        =   3
         Top             =   105
         Visible         =   0   'False
         Width           =   7740
         Begin VB.CommandButton Command3 
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
            Height          =   270
            Left            =   6270
            TabIndex        =   32
            Top             =   2925
            Width           =   1185
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "  Comprobante  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   60
            TabIndex        =   14
            Top             =   1170
            Width           =   7530
            Begin VB.Label Label28 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Unids. :"
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
               Left            =   2535
               TabIndex        =   31
               Top             =   495
               Width           =   780
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
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
               Left            =   3375
               TabIndex        =   30
               Top             =   495
               Width           =   825
            End
            Begin VB.Label Label26 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Facturas :"
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
               Left            =   150
               TabIndex        =   29
               Top             =   1065
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label25 
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
               Left            =   1200
               TabIndex        =   28
               Top             =   1065
               Visible         =   0   'False
               Width           =   2970
            End
            Begin VB.Label Label24 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Tot. Neto  :"
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
               Left            =   4425
               TabIndex        =   27
               Top             =   1365
               Width           =   1125
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
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
               Left            =   5580
               TabIndex        =   26
               Top             =   1365
               Width           =   1770
            End
            Begin VB.Label Label22 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Flete       :"
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
               Left            =   150
               TabIndex        =   25
               Top             =   795
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
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
               Left            =   1200
               TabIndex        =   24
               Top             =   795
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label Label20 
               BackColor       =   &H00C0FFC0&
               Caption         =   "IVA N/Ins.:"
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
               Left            =   4425
               TabIndex        =   23
               Top             =   1065
               Width           =   1125
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
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
               Left            =   5580
               TabIndex        =   22
               Top             =   1065
               Width           =   1770
            End
            Begin VB.Label Label18 
               BackColor       =   &H00C0FFC0&
               Caption         =   "IVA Ins.    :"
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
               Left            =   4425
               TabIndex        =   21
               Top             =   795
               Width           =   1125
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
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
               Left            =   5580
               TabIndex        =   20
               Top             =   795
               Width           =   1770
            End
            Begin VB.Label Label17 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Tot. Bruto :"
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
               Left            =   4425
               TabIndex        =   19
               Top             =   495
               Width           =   1125
            End
            Begin VB.Label Label16 
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
               Left            =   1155
               TabIndex        =   18
               Top             =   195
               Width           =   5820
            End
            Begin VB.Label Label15 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Items      :"
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
               Left            =   150
               TabIndex        =   17
               Top             =   495
               Width           =   1020
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
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
               Left            =   1215
               TabIndex        =   16
               Top             =   495
               Width           =   750
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
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
               Left            =   5580
               TabIndex        =   15
               Top             =   495
               Width           =   1770
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "  Cliente  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Left            =   60
            TabIndex        =   4
            Top             =   15
            Width           =   7530
            Begin VB.Label Label31 
               BackColor       =   &H00C0FFC0&
               Caption         =   "CUIT :"
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
               Left            =   3675
               TabIndex        =   13
               Top             =   825
               Width           =   630
            End
            Begin VB.Label Label30 
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
               Left            =   4350
               TabIndex        =   12
               Top             =   825
               Width           =   1815
            End
            Begin VB.Label Label29 
               BackColor       =   &H00C0FFC0&
               Caption         =   "IVA :"
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
               Left            =   150
               TabIndex        =   11
               Top             =   825
               Width           =   525
            End
            Begin VB.Label Label11 
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
               Left            =   720
               TabIndex        =   10
               Top             =   825
               Width           =   1815
            End
            Begin VB.Label Label10 
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
               Left            =   5940
               TabIndex        =   9
               Top             =   525
               Width           =   1500
            End
            Begin VB.Label Label9 
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
               Left            =   3675
               TabIndex        =   8
               Top             =   525
               Width           =   2175
            End
            Begin VB.Label Label8 
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
               Left            =   150
               TabIndex        =   7
               Top             =   525
               Width           =   3435
            End
            Begin VB.Label Label2 
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
               Left            =   1155
               TabIndex        =   6
               Top             =   225
               Width           =   5820
            End
            Begin VB.Label Label1 
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
               Left            =   150
               TabIndex        =   5
               Top             =   225
               Width           =   840
            End
         End
      End
   End
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   330
      Left            =   5535
      TabIndex        =   33
      Top             =   300
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
      CalendarBackColor=   16777152
      CustomFormat    =   "ddd d MMM yyy"
      Format          =   59244547
      CurrentDate     =   36877
   End
   Begin VB.Label Label32 
      Height          =   405
      Left            =   8295
      TabIndex        =   40
      Top             =   165
      Width           =   2685
   End
   Begin VB.Label Label3 
      Height          =   405
      Left            =   8295
      TabIndex        =   39
      Top             =   585
      Width           =   2685
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
      Height          =   240
      Left            =   5295
      TabIndex        =   38
      Top             =   4755
      Width           =   1185
   End
   Begin VB.Label Label7 
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
      Height          =   240
      Left            =   7875
      TabIndex        =   37
      Top             =   4755
      Width           =   1185
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
      Height          =   240
      Left            =   6600
      TabIndex        =   36
      Top             =   4755
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha:"
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
      Left            =   4050
      TabIndex        =   34
      Top             =   360
      Width           =   1410
   End
End
Attribute VB_Name = "PasaFCIVA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Registro As String, Campo As String, I1%
Dim PClte As Integer, PPtoVta As Integer, PCpbte As Double, PFecha As Date, PTipo As Integer
Dim PSECC As String, PNRO As String, Secc
Dim NomDSN2 As String, DescDSN2 As String, BDPath2 As String

Dim Cpbte As Double, TIPO As String * 2, Letra As String * 1, PtoVta As Integer

Dim TOTNETO As Double, TOTBRUT As Double, TOTIVA As Double, CanItems As Integer
Dim LaSuc As Integer, HuboErrores As Boolean

Dim dbFCIVA As New ADODB.Connection, dbCLTE As New ADODB.Connection
Dim dbFCIV  As New ADODB.Connection
Dim RgFCIVA As New ADODB.Recordset, RgCLTE As New ADODB.Recordset
Dim RgFCIV  As New ADODB.Recordset

Private Sub Command3_Click()

    Picture1.Visible = False

End Sub

Private Sub Form_Load()

FECHA.Value = Date

Me.Move 0, 0, 11300, 5600

DescDSN2 = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Descrip")
NomDSN2 = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Nombre")

BDPath2 = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Carpeta") & NomDSN2 & ".mdb"

Inicio.GenSystemDSN TipoDB, NomDSN2, DescDSN2, BDPath2
BDatos2 = "dsn=" & NomDSN2

dbFCIVA.ConnectionString = BDatos2
dbFCIVA.Open

dbFCIV.ConnectionString = BDatos1
dbFCIV.Open

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

DoVer2

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbCLTE.Close
dbFCIV.Close
dbFCIVA.Close

End Sub

Private Sub Command1_Click()
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2

  vbMsgBoxTitle = " Pasaje de Facturas Manuales "
  vbMsgBoxText = " Desea continuar ?  "
  vbMsgBoxResp = vbYesNoCancel + vbExclamation + vbApplicationModal + vbDefaultButton1
 
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

If RgFCIVA.State = 1 Then Exit Sub
          
miSQL = "Select * FROM FCIVAVTA WHERE FECHA >= " & Separa & _
        FECHADsd & Separa & " AND FECHA < " & Separa & FECHAHst & Separa & " order by fecha, cpbte"

RgFCIVA.Open miSQL, dbFCIVA, adOpenForwardOnly, adLockReadOnly
If RgFCIVA.EOF Then
    RgFCIVA.Close
    Exit Sub
End If


Do Until RgFCIVA.EOF

PCpbte = RgFCIVA!Cpbte
PPtoVta = RgFCIVA!PtoVta
PFecha = RgFCIVA!FECHA
 
Show
DoEvents

Label3.Caption = "Cpbte.=" & PCpbte & " Fecha=" & PFecha

miSQL1 = "SELECT * FROM fcivavta WHERE PTOVTA = " & PPtoVta & " and CPBTE = " & PCpbte
RgFCIV.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

If Not RgFCIV.EOF Then
'    RgFCIV.Close
    GoTo OtroReg
    End If

RgFCIV.AddNew

'RgFCIV!USUAR = "Pasaje"
   
RgFCIV!FECHA = RgFCIVA!FECHA
RgFCIV!PtoVta = RgFCIVA!PtoVta
RgFCIV!Cpbte = RgFCIVA!Cpbte
RgFCIV!Letra = RgFCIVA!Letra
RgFCIV!TIPO = RgFCIVA!TIPO

RgFCIV!CLTE = RgFCIVA!CLTE
RgFCIV!Nomb = RgFCIVA!Nomb
RgFCIV!PCIA = RgFCIVA!PCIA
RgFCIV!CVTA = RgFCIVA!CVTA

    RgFCIV!moti = RgFCIVA!moti

RgFCIV!CIVA = RgFCIVA!CIVA
RgFCIV!Cuit = RgFCIVA!Cuit
RgFCIV!VEND = RgFCIVA!VEND
RgFCIV!ZONA = RgFCIVA!ZONA
RgFCIV!TotCan = RgFCIVA!TotCan
RgFCIV!GRINS = RgFCIVA!GRINS
RgFCIV!GRNOINs = RgFCIVA!GRNOINs
RgFCIV!IVAIns = RgFCIVA!IVAIns
RgFCIV!IVANOINS = RgFCIVA!IVANOINS
RgFCIV!EXENTO = RgFCIVA!EXENTO
RgFCIV!totcos = RgFCIVA!totcos
RgFCIV!ITEMS = RgFCIVA!ITEMS
RgFCIV!bon = RgFCIVA!bon
RgFCIV!NOIMPR = RgFCIVA!NOIMPR
RgFCIV!COMIS = RgFCIVA!COMIS

    RgFCIV.Update

I1 = I1 + 1
Label32.Caption = "Grabados : " & I1
    
OtroReg:
    
    RgFCIV.Close

RgFCIVA.MoveNext

Loop

Label3.Caption = ""
Label32.Caption = ""

Close #1

Salir1:

Unload Me

Salir2:

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub FECHA_Change()

DoVer2

End Sub

Sub DoVer2()
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
       
       LimpiaGrilla

If RgFCIVA.State = 1 Then Exit Sub
          
miSQL = " fecha, tipo, letra, ptovta, cpbte, grINS, ivains, ivanoins, clte, nomb "

If LaSuc = 0 Then
    miSQL = "Select " & miSQL & " FROM FCIVAVTA WHERE FECHA >= " & Separa & _
                        FECHADsd & Separa & " AND FECHA < " & Separa & FECHAHst & Separa & " order by fecha, cpbte"
Else
    miSQL = "Select " & miSQL & " FROM FCIVAVTA WHERE FECHA >= " & Separa & _
                        FECHADsd & Separa & " AND FECHA < " & Separa & FECHAHst & Separa & " order by fecha, cpbte"
End If

RgFCIVA.Open miSQL, dbFCIVA, adOpenForwardOnly, adLockReadOnly
If RgFCIVA.EOF Then
    RgFCIVA.Close
    Exit Sub
End If
    
paraSTATUS = "Seleccionados : " & RgFCIVA.RecordCount
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

       CargaGrilla

  Label5.Caption = Format(TOTBRUT, "##,##0.00")
  Label6.Caption = Format(TOTIVA, "##,##0.00")
  Label7.Caption = Format(TOTNETO, "##,##0.00")
    
    RgFCIVA.Close


End Sub

Sub LimpiaGrilla()
Dim I1%
   
GRILLA2.Clear flexClearScrollable

For i = (GRILLA2.Rows - 1) To 1 Step -1
GRILLA2.RemoveItem (i)
Next i

End Sub

Sub CargaGrilla()

Dim NETO          As Double
Dim BRUTO         As Double
Dim ELIVA         As Double

Dim FECHA         As Date
Dim TipoCpbt      As String * 6
Dim Cpbte         As String
Dim CLTE          As String * 5
Dim Nomb          As String * 25

Dim I1, i2, Tc
Dim i3 As Long

TOTNETO = 0: TOTBRUT = 0: TOTIVA = 0

GRILLA2.Col = 0

Do Until RgFCIVA.EOF

i3 = i3 + 1

Tc = RgFCIVA!TIPO
FECHA = RgFCIVA!FECHA
RSet CLTE = RgFCIVA!CLTE
LSet Nomb = RgFCIVA!Nomb

Cpbte = RgFCIVA!Letra & " "
Cpbte = Cpbte & String((4 - Len(RgFCIVA!PtoVta)), "0") & RgFCIVA!PtoVta & "-"
Cpbte = Cpbte & String(8 - (Len(RgFCIVA!Cpbte)), "0") & RgFCIVA!Cpbte

NETO = 0: BRUTO = 0: ELIVA = 0
    
    NETO = RgFCIVA!GRINS + RgFCIVA!IVAIns
    BRUTO = RgFCIVA!GRINS
    ELIVA = RgFCIVA!IVAIns

If Tc = 1 Or Tc = 3 Or Tc = 7 Then
    TOTNETO = TOTNETO + NETO
    TOTBRUT = TOTBRUT + BRUTO
    TOTIVA = TOTIVA + ELIVA
Else
    TOTNETO = TOTNETO - NETO
    TOTBRUT = TOTBRUT - BRUTO
    TOTIVA = TOTIVA - ELIVA
End If
   
Select Case RgFCIVA!TIPO
    
    Case 0
        TipoCpbt = "Sdo.A."
    Case 1
        TipoCpbt = "Fact."
    Case 2
        TipoCpbt = "NCréd."
    Case 3
        TipoCpbt = "NDéb."
    Case 4
        TipoCpbt = "Recibo"
    Case 5
        TipoCpbt = "PACta."
    Case 6
        TipoCpbt = "Descto"
    Case 7
        TipoCpbt = "F.Most."
    Case 8
        TipoCpbt = "Dev.M."
    Case 9
        TipoCpbt = "Anul."

End Select


GRILLA2.AddItem FECHA & vbTab & TipoCpbt & vbTab & Cpbte & vbTab & BRUTO & vbTab & _
            ELIVA & vbTab & NETO & vbTab & CLTE & " " & Nomb & vbTab & _
             RgFCIVA!TIPO & vbTab & RgFCIVA!Letra & vbTab & _
             RgFCIVA!PtoVta & vbTab & RgFCIVA!Cpbte, GRILLA2.Rows
    
paraSTATUS = "Cargando : " & (GRILLA2.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

Me.Show
DoEvents

Salta:

RgFCIVA.MoveNext

Loop

GRILLA2.AutoSize 0, 7

If GRILLA2.Rows > 1 Then GRILLA2.Row = 1


' FG1.SetFocus

End Sub

Private Sub GRILLA2_Click()

If GRILLA2.Rows < 1 Then Exit Sub

MuestraDetalle

End Sub

Private Sub MuestraDetalle()
Dim FILA As Integer, ElTotal As Currency, I4%

FILA = GRILLA2.Row
TIPO = GRILLA2.TextMatrix(FILA, 7)
Letra = GRILLA2.TextMatrix(FILA, 8)
PtoVta = GRILLA2.TextMatrix(FILA, 9)
Cpbte = GRILLA2.TextMatrix(FILA, 10)

miSQL = " CPBTE = " & Cpbte & " AND TIPO = '" & TIPO & "' AND LETRA = '" & _
        Letra & "' AND PTOVTA = " & PtoVta
miSQL = "Select * FROM FCIVAVTA WHERE " & miSQL

RgFCIVA.Open miSQL, dbFCIVA, adOpenForwardOnly, adLockReadOnly
If RgFCIVA.EOF Then
    RgFCIVA.Close
    Exit Sub
End If


Label16.Caption = GRILLA2.TextMatrix(FILA, 1) & "  -  " & GRILLA2.TextMatrix(FILA, 2)
Label14.Caption = RgFCIVA!ITEMS

CanItems = RgFCIVA!ITEMS

ElTotal = RgFCIVA!GRINS + RgFCIVA!IVAIns + RgFCIVA!IVANOINS

Label13.Caption = Format(RgFCIVA!GRINS, "###,##0.00")
Label12.Caption = Format(RgFCIVA!IVAIns, "###,##0.00")
Label19.Caption = Format(RgFCIVA!IVANOINS, "###,##0.00")
Label23.Caption = Format(ElTotal, "###,##0.00")

'Label21.Caption = Format(RgFCIVA!flete, "##,##0.00")
Label27.Caption = RgFCIVA!TotCan

Label1.Caption = RgFCIVA!CLTE
Label2.Caption = RgFCIVA!Nomb
Label11.Caption = RgFCIVA!CIVA
Label30.Caption = RgFCIVA!Cuit
Label10.Caption = RgFCIVA!PCIA

      Picture1.Visible = True

miSQL = " CPBTE = " & Cpbte & " AND TIPO = " & TIPO & " AND LETRA = '" & _
        Letra & "' AND PTOVTA = " & PtoVta
miSQL = "Select * FROM Clientes WHERE Codigo = " & RgFCIVA!CLTE

RgCLTE.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly
If RgCLTE.EOF Then
    GoTo SALIR
End If

Label1.Caption = RgCLTE!CODIGO
Label2.Caption = RgCLTE!Nomb
Label8.Caption = RgCLTE!Dir
Label9.Caption = "(" & RgCLTE!CP & ") " & RgCLTE!Loc

I4 = Asc(RgCLTE!PCIA)
Label10.Caption = Mid(PCIA(I4 - 65), 2, 10)

Label11.Caption = Combo2.List(RgCLTE!CIVA - 1)
Label30.Caption = RgCLTE!Cuit

SALIR:

RgCLTE.Close
RgFCIVA.Close

End Sub

