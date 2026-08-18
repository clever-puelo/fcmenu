VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VerCheq 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Consulta de Cheques  "
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11550
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
   ScaleHeight     =   6255
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "  Opciones  "
      Height          =   825
      Left            =   8190
      TabIndex        =   61
      Top             =   90
      Width           =   3135
      Begin VB.OptionButton Option1 
         Caption         =   "Otros"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1470
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   480
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rechazados"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1470
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   210
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Egresados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   480
         Width           =   1305
      End
      Begin VB.OptionButton Option1 
         Caption         =   "En Cartera"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   210
         Width           =   1305
      End
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00F9FADC&
      Height          =   315
      ItemData        =   "VerCheq.frx":0000
      Left            =   2430
      List            =   "VerCheq.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   120
      TabIndex        =   39
      Top             =   2220
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
         TabIndex        =   40
         Top             =   465
         Width           =   4770
         Begin VB.CommandButton Command7 
            Caption         =   "Cerrar"
            Height          =   285
            Left            =   2475
            TabIndex        =   43
            Top             =   1335
            Width           =   1140
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Anular"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1260
            TabIndex        =   42
            Top             =   1335
            Width           =   1140
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0FFFF&
            Height          =   315
            Left            =   3480
            TabIndex        =   41
            Top             =   810
            Width           =   1065
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
            TabIndex        =   46
            Top             =   405
            Width           =   1350
         End
         Begin VB.Label Label33 
            BackColor       =   &H000080FF&
            Caption         =   "Ingrese el Cód. de Autorización :"
            Height          =   270
            Left            =   240
            TabIndex        =   45
            Top             =   855
            Width           =   3270
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "Código de Autorización :"
            Height          =   285
            Left            =   465
            TabIndex        =   44
            Top             =   405
            Width           =   2520
         End
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         Caption         =   "Borrado de Cheques"
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
         TabIndex        =   47
         Top             =   180
         Width           =   4770
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   330
      Left            =   10050
      TabIndex        =   34
      Top             =   4620
      Width           =   1185
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "VerCheq.frx":0052
      Left            =   4380
      List            =   "VerCheq.frx":006B
      TabIndex        =   0
      Text            =   "50"
      Top             =   525
      Width           =   780
   End
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   330
      Left            =   4320
      TabIndex        =   1
      Top             =   120
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
      Format          =   51576835
      CurrentDate     =   36877
   End
   Begin VSFlex8Ctl.VSFlexGrid GRILLA2 
      Height          =   3555
      Left            =   90
      TabIndex        =   4
      Top             =   990
      Width           =   11235
      _cx             =   19817
      _cy             =   6271
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"VerCheq.frx":008C
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
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3300
         Left            =   0
         ScaleHeight     =   3240
         ScaleWidth      =   10920
         TabIndex        =   5
         Top             =   210
         Visible         =   0   'False
         Width           =   10980
         Begin VB.CommandButton Command4 
            Caption         =   "Grabar"
            Enabled         =   0   'False
            Height          =   270
            Left            =   9570
            TabIndex        =   56
            Top             =   2925
            Width           =   1185
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Egreso"
            Height          =   270
            Left            =   4980
            TabIndex        =   51
            Top             =   2940
            Width           =   1185
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "  Egreso  "
            Height          =   2850
            Left            =   7680
            TabIndex        =   50
            Top             =   15
            Width           =   3195
            Begin VB.TextBox Text2 
               BackColor       =   &H00FEFADE&
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   1
               Left            =   90
               LinkTimeout     =   5
               MaxLength       =   50
               MultiLine       =   -1  'True
               TabIndex        =   54
               ToolTipText     =   "Detalle del Egreso"
               Top             =   1170
               Width           =   3000
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FEFADE&
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   1320
               LinkTimeout     =   5
               MaxLength       =   8
               TabIndex        =   53
               ToolTipText     =   "Comprobante de Respaldo del Egreso (Recibo del Prov. ó Bol. de Depósito)"
               Top             =   630
               Width           =   1050
            End
            Begin VB.ComboBox Combo3 
               BackColor       =   &H00F9FADC&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               ItemData        =   "VerCheq.frx":0177
               Left            =   1320
               List            =   "VerCheq.frx":0190
               Style           =   2  'Dropdown List
               TabIndex        =   52
               ToolTipText     =   "Destino del Egreso"
               Top             =   300
               Width           =   1755
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   330
               Left            =   930
               TabIndex        =   55
               ToolTipText     =   "Fecha del Egreso"
               Top             =   1950
               Width           =   2160
               _ExtentX        =   3810
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
               Format          =   51576835
               CurrentDate     =   36877
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Fecha  :"
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
               Index           =   3
               Left            =   60
               TabIndex        =   60
               Top             =   2010
               Width           =   795
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Descripción :"
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
               Index           =   2
               Left            =   90
               TabIndex        =   59
               Top             =   960
               Width           =   1125
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Nº Cpbte. :"
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
               Index           =   1
               Left            =   90
               TabIndex        =   58
               Top             =   660
               Width           =   1125
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Destino :"
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
               Index           =   0
               Left            =   90
               TabIndex        =   57
               Top             =   360
               Width           =   1125
            End
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Borrar"
            Height          =   270
            Left            =   165
            TabIndex        =   33
            Top             =   2925
            Width           =   1185
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cerrar"
            Height          =   270
            Left            =   6270
            TabIndex        =   32
            Top             =   2925
            Width           =   1185
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0FFFF&
            Caption         =   " Cheque  "
            Height          =   1695
            Left            =   60
            TabIndex        =   16
            Top             =   1170
            Width           =   7530
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Número :"
               Height          =   240
               Left            =   150
               TabIndex        =   48
               Top             =   210
               Width           =   1020
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   3240
               TabIndex        =   38
               Top             =   180
               Width           =   1290
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   5595
               TabIndex        =   37
               Top             =   180
               Width           =   1770
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Recibo :"
               Height          =   240
               Left            =   4440
               TabIndex        =   36
               Top             =   180
               Width           =   1125
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Datos Ad.:"
               Height          =   240
               Left            =   150
               TabIndex        =   31
               Top             =   1065
               Width           =   1005
            End
            Begin VB.Label Label25 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   1230
               TabIndex        =   30
               Top             =   1065
               Width           =   2970
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Importe :"
               Height          =   240
               Left            =   2445
               TabIndex        =   29
               Top             =   1365
               Width           =   1125
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
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
               Height          =   270
               Left            =   3600
               TabIndex        =   28
               Top             =   1365
               Width           =   1830
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Concepto :"
               Height          =   240
               Left            =   150
               TabIndex        =   27
               Top             =   795
               Width           =   1020
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   1230
               TabIndex        =   26
               Top             =   795
               Width           =   2985
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Vencimto.:"
               Height          =   240
               Left            =   4425
               TabIndex        =   25
               Top             =   1065
               Width           =   1125
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   5580
               TabIndex        =   24
               Top             =   1065
               Width           =   1770
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Emisión :"
               Height          =   240
               Left            =   4425
               TabIndex        =   23
               Top             =   795
               Width           =   1125
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   5580
               TabIndex        =   22
               Top             =   795
               Width           =   1770
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Ingreso :"
               Height          =   240
               Left            =   4425
               TabIndex        =   21
               Top             =   525
               Width           =   1125
            End
            Begin VB.Label Label16 
               BackColor       =   &H00FFFFFF&
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
               Height          =   270
               Left            =   1230
               TabIndex        =   20
               Top             =   195
               Width           =   1770
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               Caption         =   "Banco :"
               Height          =   240
               Left            =   150
               TabIndex        =   19
               Top             =   525
               Width           =   1020
            End
            Begin VB.Label Label14 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   1230
               TabIndex        =   18
               Top             =   525
               Width           =   2970
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   5580
               TabIndex        =   17
               Top             =   525
               Width           =   1770
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "  Cliente  "
            Height          =   1155
            Left            =   60
            TabIndex        =   6
            Top             =   15
            Width           =   7530
            Begin VB.Label Label31 
               BackColor       =   &H00C0FFFF&
               Caption         =   "CUIT :"
               Height          =   240
               Left            =   3675
               TabIndex        =   15
               Top             =   825
               Width           =   630
            End
            Begin VB.Label Label30 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   4350
               TabIndex        =   14
               Top             =   825
               Width           =   1815
            End
            Begin VB.Label Label29 
               BackColor       =   &H00C0FFFF&
               Caption         =   "IVA :"
               Height          =   240
               Left            =   150
               TabIndex        =   13
               Top             =   825
               Width           =   525
            End
            Begin VB.Label Label11 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   720
               TabIndex        =   12
               Top             =   825
               Width           =   1815
            End
            Begin VB.Label Label10 
               BackColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   5940
               TabIndex        =   11
               Top             =   525
               Width           =   1500
            End
            Begin VB.Label Label9 
               BackColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   3105
               TabIndex        =   10
               Top             =   525
               Width           =   2745
            End
            Begin VB.Label Label8 
               BackColor       =   &H00C0FFFF&
               Height          =   240
               Left            =   150
               TabIndex        =   9
               Top             =   525
               Width           =   2865
            End
            Begin VB.Label Label2 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   1155
               TabIndex        =   8
               Top             =   225
               Width           =   5820
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   1  'Fixed Single
               Height          =   240
               Left            =   150
               TabIndex        =   7
               Top             =   225
               Width           =   840
            End
         End
      End
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
      Height          =   300
      Left            =   7860
      TabIndex        =   35
      Top             =   4650
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Renglones"
      Height          =   240
      Left            =   3270
      TabIndex        =   3
      Top             =   555
      Width           =   1230
   End
   Begin VB.Label Label4 
      Caption         =   "Muestra desde"
      Height          =   240
      Left            =   2850
      TabIndex        =   2
      Top             =   180
      Width           =   1515
   End
End
Attribute VB_Name = "VerCheq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cpbte As Double, TIPO As String * 2, Letra As String * 1, PtoVta As Integer
Dim Ubi3X As Double, Ubi3Y As Double, SeMueve3 As Boolean, NroCheq As String
Dim ElEstado As Integer
Dim TOTCHEQ As Double, HuboErrores As Boolean

Dim dbMAE As New ADODB.Connection

Dim RgCHEQ As New ADODB.Recordset, RgPARM As New ADODB.Recordset
Dim RgCLTE As New ADODB.Recordset, RgCYB As New ADODB.Recordset


Private Sub Combo3_Click()

If Text2(0).Enabled = True Then Text2(0).SetFocus

End Sub

Private Sub Form_Load()

FECHA.Value = Date

Me.Move 0, 0, 11650, 5500

dbMAE.ConnectionString = BDatos1
dbMAE.Open

Option1(1).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbMAE.Close

End Sub

Private Sub Form_Deactivate()

'Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then Unload Me

If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
'If KeyCode = 40 Then SendKeys "{TAB}"
'If KeyCode = 38 Then SendKeys "+{TAB}"
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Combo1_Click()

DoVer2

End Sub

'--------------------------------------
'   Cierra la Ventana de Cheques - Salida Total
Private Sub Command1_Click()

Unload Me

End Sub

'------------------------------------
'  Cierra Detalle del Cheque

Private Sub Command2_Click()

    Picture1.Visible = False

End Sub
'----------------------------------------
' Abre ventana para el Egreso
Private Sub Command3_Click()

Text2(0).Text = ""
Text2(1).Text = ""

DTPicker1.Value = Date

Command4.Enabled = False

Picture1.width = 11000

Combo3.SetFocus

End Sub

'---------------------------------------
'Graba el Egreso del Cheque
Private Sub Command4_Click()
     
  vbMsgBoxTitle = "Grabación del Egreso"
  vbMsgBoxText = "El siguiente proceso registrará el Egreso del Cheque  " & vbCrLf & "Desea continuar ?  "
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

miSQL1 = "SELECT * FROM Cheques WHERE NROCHEQ = " & NroCheq
RgCHEQ.Open miSQL1, dbMAE, adOpenDynamic, adLockPessimistic

If RgCHEQ.EOF Then
    RgCHEQ.Close
    GoTo Salir1
    End If
    
    RgCHEQ!tipegr = Combo3.ListIndex
    RgCHEQ!cpbegr = Text2(0).Text
    RgCHEQ!destino = Text2(1).Text
    RgCHEQ!fecegr = DTPicker1.Value
    RgCHEQ!Estado = 2
    
RgCHEQ.Update
RgCHEQ.Close

Salir1:

Picture1.width = 7700
Picture1.Visible = False
DoVer2

Salir2:

End Sub

'------------------------------------
'  Borrado del Cheque
Private Sub Command5_Click()

    BorraCheque

End Sub

'---------------------------------------------
'  Acepta la Clave de Anulación (Borrado del Cheque)
Private Sub Command6_Click()

    GrabaAnul

End Sub

'---------------------------------------------
'   Cierra la Ventana de Clave
Private Sub Command7_Click()

 Frame5.Visible = False
 
End Sub


Private Sub FECHA_Change()

DoVer2

End Sub

Sub DoVer2()
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2, CantReg As Integer
Dim PaFecha1 As Date, PaFecha2 As Date

FECHADsd = FECHA.Value
ElDia = 31
If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
FECHADsd = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year
FECHAHst = FECHA.Month & "/" & ElDia & "/" & FECHA.Year
  
  If ORDENADOx <> "" Then
         Clasifica = elORDEN & ORDENADOx & ORDENADOad
  Else
         Clasifica = ""
End If

'///           ----------------------------------------
       
       LimpiaGrilla

If RgCHEQ.State = 1 Then Exit Sub
          
'miSQL = " nrocheq, fecha, orden,  cpbte, grINS, ivains, ivanoins, clte, nomb "
'miSQL = Year(FECHADsd)
PaFecha1 = FECHADsd: PaFecha2 = FECHAHst
miSQL = "SELECT count(*) as CanReg FROM Cheques WHERE FECing >= " & Separa & PaFecha1 & _
        Separa & " and ESTADO = '" & ElEstado & "'"
        
RgCHEQ.Open miSQL, dbMAE, adOpenForwardOnly, adLockReadOnly


If RgCHEQ.EOF Then
    RgCHEQ.Close
    Exit Sub
End If

CantReg = RgCHEQ!CanReg
RgCHEQ.Close
    
paraSTATUS = "Seleccionados : " & CantReg
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

miSQL = "SELECT * FROM Cheques WHERE FECing >= " & Separa & PaFecha1 & _
        Separa & " and ESTADO = '" & ElEstado & "' ORDER BY FECVTO, NROCHEQ ASC"

RgCHEQ.Open miSQL, dbMAE, adOpenForwardOnly, adLockReadOnly
If RgCHEQ.EOF Then
    RgCHEQ.Close
    Exit Sub
End If

       CargaGrilla

  Label7.Caption = Format(TOTCHEQ, "$ ####,##0.00")
    
    RgCHEQ.Close


End Sub

Sub LimpiaGrilla()
Dim I1%
   
GRILLA2.Clear flexClearScrollable

For i = (GRILLA2.Rows - 1) To 1 Step -1
GRILLA2.RemoveItem (i)
Next i

End Sub

Sub CargaGrilla()

Dim LaFECHA       As Date
Dim TipoCpbt      As String
Dim ElCheque      As String * 10
Dim CLTE          As String * 5
Dim Nomb          As String * 20
Dim BcoSuc        As String
Dim ElRecibo      As String

Dim I1, i2, Tc As Boolean
Dim i3 As Long

TOTCHEQ = 0

GRILLA2.Col = 0

Do Until RgCHEQ.EOF Or i3 >= Combo1.Text

i3 = i3 + 1

LaFECHA = RgCHEQ!FECVTO
RSet CLTE = RgCHEQ!CLTE
miSQL = "Select nomb FROM Clientes WHERE CODIGO = " & RgCHEQ!CLTE

RgCLTE.Open miSQL, dbMAE, adOpenForwardOnly, adLockReadOnly
If Not RgCLTE.EOF Then
    Nomb = RgCLTE!Nomb
Else
    Nomb = "*** Desconocido ***"
End If
    
    RgCLTE.Close

'Nomb = RgCHEQ!Nomb
BcoSuc = RgCHEQ!BcoSuc
ElRecibo = Format(RgCHEQ!cpbing, "###,##0")

RSet ElCheque = RgCHEQ!NroCheq

GRILLA2.AddItem ElCheque & vbTab & LaFECHA & vbTab & BcoSuc & vbTab & ElRecibo & vbTab & _
            CLTE & " " & Nomb & vbTab & _
             RgCHEQ!Importe & vbTab & RgCHEQ!orden, GRILLA2.Rows
    
paraSTATUS = "Cargando : " & (GRILLA2.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

TOTCHEQ = TOTCHEQ + RgCHEQ!Importe

Me.Show
DoEvents

Salta:

RgCHEQ.MoveNext

Loop

GRILLA2.AutoSize 0, 6

If GRILLA2.Rows > 1 Then GRILLA2.Row = 1


' FG1.SetFocus

End Sub

Private Sub GRILLA2_Click()

If GRILLA2.Rows < 1 Then Exit Sub

MuestraDetalle

End Sub

Private Sub MuestraDetalle()
Dim FILA As Integer, ElTotal As Currency
Dim ElCheque As String * 10, I4%

FILA = GRILLA2.Row
ElCheque = GRILLA2.TextMatrix(FILA, 0)
NroCheq = GRILLA2.TextMatrix(FILA, 0)

miSQL = "Select * FROM Cheques WHERE NROCHEQ = " & ElCheque

RgCHEQ.Open miSQL, dbMAE, adOpenForwardOnly, adLockReadOnly
If RgCHEQ.EOF Then
    RgCHEQ.Close
    Exit Sub
End If

Label1.Caption = RgCHEQ!CLTE


Label16.Caption = GRILLA2.TextMatrix(FILA, 0)
Label23.Caption = Format(RgCHEQ!Importe, "$ ###,##0.00")
Label27.Caption = ""
If RgCHEQ!orden Then Label27.Caption = "N/Orden"

Label6.Caption = Format(RgCHEQ!cpbing, "####,##0")
Label14.Caption = RgCHEQ!BcoSuc & " - *** No Existe en Tabla *** "
Label13.Caption = Format(RgCHEQ!FECING, "ddd dd mmm yyyy")
Label12.Caption = Format(RgCHEQ!fecemi, "ddd dd mmm yyyy")
Label19.Caption = Format(RgCHEQ!FECVTO, "ddd dd mmm yyyy")
Label21.Caption = RgCHEQ!concep
Label25.Caption = RgCHEQ!datosad

miSQL1 = "Select * FROM Bancos where COD = " & RgCHEQ!BcoSuc
RgCYB.Open miSQL1, dbMAE, adOpenDynamic, adLockPessimistic

    If Not RgCYB.EOF Then
         Label14.Caption = RgCHEQ!BcoSuc & " - " & RgCYB!NOMBRE
        End If
         
RgCYB.Close

If ElEstado = 1 Then
      Combo3.Enabled = True
      Text2(0).Enabled = True
      Text2(1).Enabled = True
      DTPicker1.Enabled = True
      Command4.Visible = True
      Picture1.width = 7700
Else
      Combo3.Enabled = False
      Text2(0).Enabled = False
      Text2(1).Enabled = False
      DTPicker1.Enabled = False
      Command4.Visible = False
      Combo3.ListIndex = RgCHEQ!tipegr
      Text2(0).Text = RgCHEQ!cpbegr
      Text2(1).Text = RgCHEQ!destino
      DTPicker1.Value = RgCHEQ!fecegr
      Picture1.width = 11000
End If

      Picture1.Visible = True

miSQL1 = "Select * FROM Clientes WHERE CODIGO = " & RgCHEQ!CLTE
RgCLTE.Open miSQL1, dbMAE, adOpenForwardOnly, adLockReadOnly
If RgCLTE.EOF Then
    RgCLTE.Close
    GoTo SaltaClte
End If

Label1.Caption = RgCLTE!CODIGO
Label2.Caption = RgCLTE!Nomb
Label8.Caption = RgCLTE!Dir
Label9.Caption = "(" & RgCLTE!CP & ") " & RgCLTE!Loc

I4 = Asc(RgCLTE!PCIA)
Label10.Caption = Mid(PCIA(I4 - 65), 2, 10)

Label11.Caption = Combo2.List(RgCLTE!CIVA - 1)
Label30.Caption = RgCLTE!Cuit

SaltaClte:

RgCLTE.Close
RgCHEQ.Close

End Sub

Private Sub BorraCheque()
Dim I1%, i2%, CLV As String * 5, i3%, I4%

Text1.Text = ""
Randomize

For I1 = 1 To 5
        i3 = Mid(Time, 7, 2)
        i2 = Int((26 * Rnd(i3)) + 65) ' Genera valores aleatorios entre 65 y 90.
        
        Mid(CLV, I1, 1) = Chr(i2)
            
Next I1
Label32.Caption = CLV

Frame5.Visible = True
Text1.SetFocus


End Sub

Private Sub Option1_Click(Index As Integer)

ElEstado = Index
If Index = 1 Then
    Command3.Visible = True
Else
    Command3.Visible = False
End If
    
DoVer2

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

Respuesta = MsgBox("Realmente desea Borrar ? ", vbQuestion + vbYesNo, " Borrado - Proceso Final ")

If Respuesta = vbNo Then Exit Sub

HuboErrores = False
    
    dbMAE.BeginTrans

'----------------------------------------------
'   Relee el Registro de Cheques

miSQL = "Select * FROM Cheques WHERE NROCHEQ = " & NroCheq
RgCHEQ.Open miSQL, dbMAE, adOpenDynamic, adLockPessimistic

If RgCHEQ.EOF Then
    HuboErrores = True
    RgCHEQ.Close
    GoTo Sale6
End If

RgCHEQ.Delete

Sale6:

'  Si salió todo bien, Confirma las Transacciones
If HuboErrores Then
    dbMAE.RollbackTrans
    MsgBox "Hubo Errores durante la GRABACIÓN" & vbCrLf & "VERIFIQUE !!!", vbExclamation, "Aviso"
Else
    dbMAE.CommitTrans
End If

RgCHEQ.Close
 
End Sub

'--------------------------------------------
'    Movimiento de Ventanas

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

SeMueve3 = True
Ubi3X = X
Ubi3Y = Y

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If SeMueve3 Then
    Picture1.left = Picture1.left - (Ubi3X - X)
    Picture1.top = Picture1.top - (Ubi3Y - Y)
End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

SeMueve3 = False

End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

SeMueve3 = True
Ubi3X = X
Ubi3Y = Y

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If SeMueve3 Then
    Picture1.left = Picture1.left - (Ubi3X - X)
    Picture1.top = Picture1.top - (Ubi3Y - Y)
End If

End Sub

Private Sub Frame2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

SeMueve3 = False

End Sub


Private Sub Text2_GotFocus(Index As Integer)

Text2(Index).SelStart = 0
Text2(Index).SelLength = Len(Text2(Index).Text)

End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> vbKeyReturn Then Exit Sub

If Index > 0 Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub ValidaEgreso()

If Combo3.Text = "" Then
    Combo3.SetFocus
    Exit Sub
End If

If Text2(0).Text = "" Then
    Text2(0).SetFocus
    Exit Sub
End If

If Text2(1).Text = "" Then
    Text2(1).SetFocus
    Exit Sub
End If

Command4.Enabled = True

End Sub

Private Sub Text2_LostFocus(Index As Integer)

ValidaEgreso

End Sub
