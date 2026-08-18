VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VerFact 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Facturas Emitidas"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11160
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
   ScaleHeight     =   6345
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00F9FADC&
      Height          =   315
      ItemData        =   "VerFact.frx":0000
      Left            =   5280
      List            =   "VerFact.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Top             =   5790
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "VerFact.frx":0052
      Left            =   5685
      List            =   "VerFact.frx":006B
      TabIndex        =   17
      Text            =   "50"
      Top             =   495
      Width           =   780
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   330
      Left            =   9870
      TabIndex        =   16
      Top             =   4500
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   240
      TabIndex        =   9
      Top             =   -30
      Visible         =   0   'False
      Width           =   3885
      Begin VB.OptionButton Option1 
         Caption         =   "Todas"
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   150
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cta. Cte."
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   14
         Top             =   390
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mostrador"
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   645
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Manuales"
         Height          =   210
         Index           =   3
         Left            =   2100
         TabIndex        =   12
         Top             =   150
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Empresas"
         Height          =   210
         Index           =   4
         Left            =   2100
         TabIndex        =   11
         Top             =   390
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Exportación"
         Height          =   210
         Index           =   5
         Left            =   2100
         TabIndex        =   10
         Top             =   645
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H000040C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   255
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   5040
      Begin VB.Frame Frame4 
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   135
         TabIndex        =   1
         Top             =   465
         Width           =   4770
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3480
            TabIndex        =   4
            Top             =   810
            Width           =   1065
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Anular"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1260
            TabIndex        =   3
            Top             =   1335
            Width           =   1140
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Cerrar"
            Height          =   285
            Left            =   2475
            TabIndex        =   2
            Top             =   1335
            Width           =   1140
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "Código de Autorización :"
            Height          =   285
            Left            =   465
            TabIndex        =   7
            Top             =   405
            Width           =   2520
         End
         Begin VB.Label Label33 
            BackColor       =   &H000080FF&
            Caption         =   "Ingrese el Cód. de Autorización :"
            Height          =   270
            Left            =   240
            TabIndex        =   6
            Top             =   855
            Width           =   3270
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            TabIndex        =   5
            Top             =   405
            Width           =   1350
         End
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Anulación de Comprobantes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   135
         TabIndex        =   8
         Top             =   180
         Width           =   4770
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid GRILLA2 
      Height          =   3555
      Left            =   240
      TabIndex        =   18
      Top             =   915
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
      FormatString    =   $"VerFact.frx":008C
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
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00AFF8CC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   2295
         ScaleHeight     =   2775
         ScaleWidth      =   8415
         TabIndex        =   19
         Top             =   375
         Visible         =   0   'False
         Width           =   8475
         Begin VB.CommandButton Command4 
            Caption         =   "Cerrar"
            Height          =   270
            Left            =   7140
            TabIndex        =   20
            Top             =   2460
            Width           =   1185
         End
         Begin VSFlex8Ctl.VSFlexGrid FG2 
            Height          =   2400
            Left            =   30
            TabIndex        =   21
            Top             =   0
            Width           =   8340
            _cx             =   14711
            _cy             =   4233
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   13695957
            ForeColor       =   -2147483640
            BackColorFixed  =   12648384
            ForeColorFixed  =   32768
            BackColorSel    =   16777215
            ForeColorSel    =   32768
            BackColorBkg    =   12648384
            BackColorAlternate=   15335153
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"VerFact.frx":01DD
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
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3270
         Left            =   3030
         ScaleHeight     =   3210
         ScaleWidth      =   7680
         TabIndex        =   22
         Top             =   -60
         Visible         =   0   'False
         Width           =   7740
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "  Cliente  "
            Height          =   1155
            Left            =   60
            TabIndex        =   44
            Top             =   15
            Width           =   7530
            Begin VB.Label Label1 
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   150
               TabIndex        =   53
               Top             =   225
               Width           =   840
            End
            Begin VB.Label Label2 
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   1155
               TabIndex        =   52
               Top             =   225
               Width           =   5820
            End
            Begin VB.Label Label8 
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   150
               TabIndex        =   51
               Top             =   525
               Width           =   3435
            End
            Begin VB.Label Label9 
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   3675
               TabIndex        =   50
               Top             =   525
               Width           =   2175
            End
            Begin VB.Label Label10 
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   5940
               TabIndex        =   49
               Top             =   525
               Width           =   1500
            End
            Begin VB.Label Label11 
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   720
               TabIndex        =   48
               Top             =   825
               Width           =   1815
            End
            Begin VB.Label Label29 
               BackColor       =   &H00C0FFC0&
               Caption         =   "IVA :"
               Height          =   240
               Left            =   150
               TabIndex        =   47
               Top             =   825
               Width           =   525
            End
            Begin VB.Label Label30 
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   4350
               TabIndex        =   46
               Top             =   825
               Width           =   1815
            End
            Begin VB.Label Label31 
               BackColor       =   &H00C0FFC0&
               Caption         =   "CUIT :"
               Height          =   240
               Left            =   3675
               TabIndex        =   45
               Top             =   825
               Width           =   630
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "  Comprobante  "
            Height          =   1650
            Left            =   60
            TabIndex        =   26
            Top             =   1170
            Width           =   7530
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   5580
               TabIndex        =   43
               Top             =   480
               Width           =   1770
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   1215
               TabIndex        =   42
               Top             =   495
               Width           =   750
            End
            Begin VB.Label Label15 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Items      :"
               Height          =   240
               Left            =   150
               TabIndex        =   41
               Top             =   495
               Width           =   1020
            End
            Begin VB.Label Label16 
               BackColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   1155
               TabIndex        =   40
               Top             =   195
               Width           =   5820
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Tot. Bruto :"
               Height          =   240
               Left            =   4425
               TabIndex        =   39
               Top             =   480
               Width           =   1125
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   5565
               TabIndex        =   38
               Top             =   750
               Width           =   1785
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "IVA Ins.:"
               Height          =   240
               Left            =   4410
               TabIndex        =   37
               Top             =   750
               Width           =   1125
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   5580
               TabIndex        =   36
               Top             =   1020
               Width           =   1770
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "I.B. Bs.As.:"
               Height          =   240
               Left            =   4425
               TabIndex        =   35
               Top             =   1020
               Width           =   1125
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   1200
               TabIndex        =   34
               Top             =   795
               Visible         =   0   'False
               Width           =   1845
            End
            Begin VB.Label Label22 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Flete       :"
               Height          =   240
               Left            =   150
               TabIndex        =   33
               Top             =   795
               Visible         =   0   'False
               Width           =   1020
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   5580
               TabIndex        =   32
               Top             =   1305
               Width           =   1770
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Tot. Neto :"
               Height          =   240
               Left            =   4425
               TabIndex        =   31
               Top             =   1305
               Width           =   1125
            End
            Begin VB.Label Label25 
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   1200
               TabIndex        =   30
               Top             =   1065
               Visible         =   0   'False
               Width           =   2970
            End
            Begin VB.Label Label26 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Facturas :"
               Height          =   240
               Left            =   150
               TabIndex        =   29
               Top             =   1065
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFC0&
               Height          =   240
               Left            =   3375
               TabIndex        =   28
               Top             =   495
               Width           =   825
            End
            Begin VB.Label Label28 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Unids. :"
               Height          =   240
               Left            =   2535
               TabIndex        =   27
               Top             =   495
               Width           =   780
            End
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cerrar"
            Height          =   270
            Left            =   6285
            TabIndex        =   25
            Top             =   2880
            Width           =   1185
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Ver Items"
            Height          =   270
            Left            =   5010
            TabIndex        =   24
            Top             =   2880
            Width           =   1185
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Anular"
            Height          =   270
            Left            =   165
            TabIndex        =   23
            Top             =   2925
            Width           =   1185
         End
      End
   End
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   330
      Left            =   5625
      TabIndex        =   54
      Top             =   90
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
      Format          =   53215235
      CurrentDate     =   36877
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
      Left            =   6825
      TabIndex        =   59
      Top             =   4545
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
      Left            =   8100
      TabIndex        =   58
      Top             =   4545
      Width           =   1185
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
      Left            =   5520
      TabIndex        =   57
      Top             =   4545
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "Muestra desde"
      Height          =   240
      Left            =   4140
      TabIndex        =   56
      Top             =   150
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Los Primeros"
      Height          =   240
      Left            =   4350
      TabIndex        =   55
      Top             =   525
      Width           =   1365
   End
End
Attribute VB_Name = "VerFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cpbte As Double, TIPO As String * 2, Letra As String * 1, PtoVta As Integer

Dim TOTNETO As Double, TOTBRUT As Double, TOTIVA As Double, CanItems As Integer
Dim LaSuc As Integer, HuboErrores As Boolean

Dim dbMAE As New ADODB.Connection, dbFACT As New ADODB.Connection
Dim dbCCTE As New ADODB.Connection, dbSTOCK As New ADODB.Connection
Dim dbEXPED As New ADODB.Connection, dbESTAD As New ADODB.Connection

Dim RgFCIVA As New ADODB.Recordset, RgPARM As New ADODB.Recordset
Dim RgCLTE As New ADODB.Recordset, RgDESPA As New ADODB.Recordset
Dim RgEST1 As New ADODB.Recordset

Private Sub Form_Load()

FECHA.Value = Date

Me.Move 0, 0, 11500, 5300

dbFACT.ConnectionString = BDatos1
dbFACT.Open

dbMAE.ConnectionString = BDatos1
dbMAE.Open

dbESTAD.ConnectionString = BDatos1
dbESTAD.Open

Option1(1).Value = False
Option1(1).Value = True

'DoVer2

End Sub

Private Sub Form_Unload(Cancel As Integer)


dbMAE.Close
dbFACT.Close
dbESTAD.Close

End Sub

Private Sub Form_Deactivate()

'Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then Unload Me

End Sub

Private Sub Combo1_Click()

DoVer2

End Sub

Private Sub Command1_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Command2_Click()

    Picture1.Visible = False

End Sub

Private Sub Command4_Click()

    Picture2.Visible = False

End Sub

Private Sub Command5_Click()

    AnulaFactura

End Sub

Private Sub Command6_Click()

    GrabaAnul

End Sub

Private Sub Command7_Click()

Frame5.Visible = False

End Sub

Private Sub Command3_Click()
Dim Cpbte As Double, TIPO As Integer, Letra As String * 1, PtoVta As Integer
Dim FILA As Integer, ElTotal As Currency, i3%

FILA = GRILLA2.Row
TIPO = GRILLA2.TextMatrix(FILA, 7)
Letra = GRILLA2.TextMatrix(FILA, 8)
PtoVta = GRILLA2.TextMatrix(FILA, 9)
Cpbte = GRILLA2.TextMatrix(FILA, 10)

miSQL = " CPBTE = " & Cpbte & " AND TIPO = " & TIPO & " AND LETRA = '" & _
        Letra & "' AND PTOVTA = " & PtoVta
miSQL = "Select * FROM FCEstad1 WHERE " & miSQL
RgEST1.Open miSQL, dbESTAD, adOpenForwardOnly, adLockReadOnly

If RgEST1.EOF Then
    RgEST1.Close
    Exit Sub
End If
   
FG2.Clear flexClearScrollable

For i3 = (FG2.Rows - 1) To 1 Step -1
FG2.RemoveItem (i3)
Next i3

'Set FG2.DataSource = RgEST1

Do Until RgEST1.EOF

i3 = i3 + 1

FG2.AddItem RgEST1!COD1 & vbTab & RgEST1!Pulg & vbTab & RgEST1!Mtr & vbTab & _
            RgEST1!Milim & vbTab & RgEST1!Telas & vbTab & RgEST1!Cant & _
            vbTab & RgEST1!PEsp & vbTab & RgEST1!PCos & _
            vbTab & RgEST1!PVta & vbTab & RgEST1!Impte, FG2.Rows
    
paraSTATUS = "Cargando : " & (FG2.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

RgEST1.MoveNext

Loop

FG2.AutoSize 0, 5

    Picture2.Visible = True

RgEST1.Close

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

RgFCIVA.Open miSQL, dbFACT, adOpenForwardOnly, adLockReadOnly
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

Do Until RgFCIVA.EOF Or i3 >= Combo1.Text

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

RgFCIVA.Open miSQL, dbFACT, adOpenForwardOnly, adLockReadOnly
If RgFCIVA.EOF Then
    RgFCIVA.Close
    Exit Sub
End If


Label16.Caption = GRILLA2.TextMatrix(FILA, 1) & "  -  " & GRILLA2.TextMatrix(FILA, 2)
Label14.Caption = RgFCIVA!ITEMS

CanItems = RgFCIVA!ITEMS

        If Not RgFCIVA!TotIB Then
            ElTotal = RgFCIVA!GRINS + RgFCIVA!IVAIns + RgFCIVA!TotIB
        Else
            ElTotal = RgFCIVA!GRINS + RgFCIVA!IVAIns
        End If

Label13.Caption = Format(RgFCIVA!GRINS, "###,##0.00")
Label12.Caption = Format(RgFCIVA!IVAIns, "###,##0.00")
Label19.Caption = Format(RgFCIVA!TotIB, "###,##0.00")
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

RgCLTE.Open miSQL, dbMAE, adOpenForwardOnly, adLockReadOnly
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

Private Sub Option1_Click(Index As Integer)

LaSuc = Index

DoVer2

End Sub

Private Sub AnulaFactura()
Dim I1%, i2%, CLV As String * 5

Randomize
Text1.Text = ""

For I1 = 1 To 5
        
        i2 = Int((26 * Rnd(Mid(Time, 7, 2))) + 65) ' Genera valores aleatorios entre 65 y 90.
            
        Mid(CLV, I1, 1) = Chr(i2)
            
Next I1
Label32.Caption = CLV

Frame5.Visible = True
Text1.SetFocus


End Sub

Private Sub text1_Change()

If Text1.Text = Label32.Caption Then
    Command6.Enabled = True
    Else
    Command6.Enabled = False
End If

End Sub

'***************************************************************
'            ANULACIÓM DE COMPROBANTES
'---------------------------------------------------------------
'                   Totales
'--------------------------------------------------------------

Private Sub GrabaAnul()
Dim RgCCTE As New ADODB.Recordset, RgFCIVA As New ADODB.Recordset
Dim RgTOTL As New ADODB.Recordset, RgNENV As New ADODB.Recordset, RgDESPA As New ADODB.Recordset

Respuesta = MsgBox("Realmente desea ANULAR ? ", vbQuestion + vbYesNo, " ANULACIÓN - Proceso Final ")

If Respuesta = vbNo Then Exit Sub

HuboErrores = False
    
    dbMAE.BeginTrans
    dbFACT.BeginTrans
    dbESTAD.BeginTrans

'----------------------------------------------
'   Relee el Registro del Subdiario

miSQL = " CPBTE = " & Cpbte & " AND TIPO = '" & TIPO & "' AND LETRA = '" & _
        Letra & "' AND PTOVTA = " & PtoVta
miSQL = "Select * FROM FCIVAVTA WHERE " & miSQL

RgFCIVA.Open miSQL, dbFACT, adOpenDynamic, adLockPessimistic

If RgFCIVA.EOF Then
    HuboErrores = True
    RgFCIVA.Close
    GoTo Sale6
End If

BorraDetalle

'----------------------------------------------
'   Borra en la Cuenta Cte.

miSQL2 = "SELECT * FROM CtasCtes WHERE CPBTE = " & RgFCIVA!Cpbte & _
      " AND TIPO = " & RgFCIVA!TIPO & _
      " AND FECHA = '" & Format(RgFCIVA!FECHA, "dd/mm/yyyy") & _
      "' AND PREFIJO = " & RgFCIVA!PtoVta & _
      " AND CLTE = " & RgFCIVA!CLTE
' Debug.Print miSQL2
RgCLTE.Open miSQL2, dbFACT, adOpenDynamic, adLockPessimistic

If RgCLTE.EOF Then
    RgCLTE.Close
    GoTo Sale1
End If

RgCLTE.Delete: RgCLTE.Close

Sale1:

'----------------------------------------------
'   Actualiza Totales - Borra el importe de la Factura
laLOGIK1 = Month(RgFCIVA!FECHA) & "/" & Day(RgFCIVA!FECHA) & "/" & Year(RgFCIVA!FECHA)
miSQL2 = "SELECT * FROM Totales WHERE FECHA = " & Separa & laLOGIK1 & Separa
RgTOTL.Open miSQL2, dbFACT, adOpenDynamic, adLockPessimistic
      

If RgTOTL.EOF Then
    RgTOTL.Close
    GoTo Sale2
End If

If Letra = "A" Then
   RgTOTL!PESPA = RgTOTL!PESPA - RgFCIVA!GRINS
   RgTOTL!UNIDA = RgTOTL!UNIDA - RgFCIVA!TotCan
   RgTOTL!faca = RgTOTL!faca - 1
Else
   RgTOTL!PESPb = RgTOTL!PESPb - RgFCIVA!GRINS
   RgTOTL!UNIDB = RgTOTL!UNIDB - RgFCIVA!TotCan
   RgTOTL!facb = RgTOTL!facb - 1
End If
   
   RgTOTL!PCos = RgTOTL!PCos - RgFCIVA!totcos
   RgTOTL!PVta = RgTOTL!PVta - RgFCIVA!GRINS
   RgTOTL!PEsp = RgTOTL!PEsp - RgFCIVA!GRINS

RgTOTL("Factual") = Date
RgTOTL("USUARIO") = Mid$(Red_Usuario, 1, 6)

RgTOTL.Update: RgTOTL.Close

Sale2:
'----------------------------------------------
'   Actualiza el Saldo en CLIENTES

miSQL2 = "SELECT * FROM Clientes WHERE CODIGO = " & RgFCIVA!CLTE
RgCLTE.Open miSQL2, dbMAE, adOpenDynamic, adLockPessimistic

If RgCLTE.EOF Then
    RgCLTE.Close
    GoTo Sale4
End If

RgCLTE!DEUDA = RgCLTE!DEUDA - (RgFCIVA!GRINS + RgFCIVA!IVAIns + RgFCIVA!IVANOINS)

RgCLTE!FACTUAL = Date
RgCLTE!USUARIO = Mid$(Red_Usuario, 1, 6)

RgCLTE.Update

Sale4:

RgCLTE.Close

'----------------------------------------------
'   Actualiza el FCIVA - Pone 9 en Tipo

RgFCIVA!TIPO = 9

RgFCIVA!TotCan = 0
RgFCIVA!GRINS = 0
RgFCIVA!GRNOINs = 0
RgFCIVA!IVAIns = 0
RgFCIVA!IVANOINS = 0
RgFCIVA!EXENTO = 0
RgFCIVA!totcos = 0
RgFCIVA!ITEMS = 0
RgFCIVA!bon = 0

'RgFCIVA!USUARIO = "      "

RgFCIVA.Update

Sale6:

'  Si salió todo bien, Confirma las Transacciones
If HuboErrores Then
    dbMAE.RollbackTrans
    dbFACT.RollbackTrans
    dbESTAD.RollbackTrans
    MsgBox "Hubo Errores durante la GRABACIÓN" & vbCrLf & "VERIFIQUE !!!", vbExclamation, "Aviso"

Else
    dbMAE.CommitTrans
    dbFACT.CommitTrans
    dbESTAD.CommitTrans
End If

'FECHA.Value = RgFCIVA!FECHA

RgFCIVA.Close

 Frame5.Visible = False
 Frame4.Visible = False
 
 
End Sub

'**********************************************************
'                    Anulación del Detalle
'----------------------------------------------------------

Private Sub BorraDetalle()
Dim I1%, TipMOVS As Integer

Dim RgEST1 As New ADODB.Recordset, RgSTOCK As New ADODB.Recordset
Dim RgMOVS As New ADODB.Recordset, RgSALI As New ADODB.Recordset

miSQL = " CPBTE = " & Cpbte & " AND TIPO = " & TIPO & " AND LETRA = '" & _
        Letra & "' AND PTOVTA = " & PtoVta
miSQL = "Select * FROM FCEstad1 WHERE " & miSQL
RgEST1.Open miSQL, dbESTAD, adOpenKeyset, adLockPessimistic

If RgEST1.EOF Then
'    huboerrores=True
    RgEST1.Close
    Exit Sub
End If

'-----------------------------------
'     Comienza el Loop

Do Until RgEST1.EOF

'--------------------------------------
' Borra Estadistica y Lee el Próximo
RgEST1.Delete

RgEST1.MoveNext

Loop

RgEST1.Close

End Sub


