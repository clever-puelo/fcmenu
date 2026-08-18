VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form CtaCte 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas Corrientes"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13710
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
   ScaleHeight     =   5625
   ScaleWidth      =   13710
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   106
      Top             =   60
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "   Saldos de Cuenta Corriente   "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5115
      Left            =   360
      TabIndex        =   26
      Top             =   420
      Visible         =   0   'False
      Width           =   8535
      Begin VSFlex8Ctl.VSFlexGrid FG8 
         Height          =   4395
         Left            =   840
         TabIndex        =   27
         Top             =   270
         Width           =   6630
         _cx             =   11695
         _cy             =   7752
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
         BackColor       =   14089948
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"CtaCte.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   1110
         TabIndex        =   28
         Top             =   4680
         Width           =   1380
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   6000
         TabIndex        =   30
         Top             =   4710
         Width           =   1365
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         Caption         =   "SALDO :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   5010
         TabIndex        =   29
         Top             =   4710
         Width           =   1005
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Otro Cliente"
      Height          =   285
      Left            =   11130
      TabIndex        =   104
      Top             =   60
      Width           =   2160
   End
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
      Left            =   120
      ScaleHeight     =   2775
      ScaleWidth      =   8415
      TabIndex        =   63
      Top             =   2100
      Visible         =   0   'False
      Width           =   8475
      Begin VB.CommandButton Command4 
         Caption         =   "Cerrar"
         Height          =   270
         Left            =   7140
         TabIndex        =   64
         Top             =   2460
         Width           =   1185
      End
      Begin VSFlex8Ctl.VSFlexGrid FG2 
         Height          =   2400
         Left            =   30
         TabIndex        =   65
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
         FormatString    =   $"CtaCte.frx":006B
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
   Begin VB.PictureBox Picture3 
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
      Height          =   3300
      Left            =   840
      ScaleHeight     =   3240
      ScaleWidth      =   7680
      TabIndex        =   31
      Top             =   450
      Visible         =   0   'False
      Width           =   7740
      Begin VB.CommandButton Command6 
         Caption         =   "Ver Items"
         Height          =   270
         Left            =   4995
         TabIndex        =   61
         Top             =   2925
         Width           =   1185
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cerrar"
         Height          =   270
         Left            =   6270
         TabIndex        =   60
         Top             =   2925
         Width           =   1185
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "  Comprobante  "
         Height          =   1695
         Left            =   60
         TabIndex        =   42
         Top             =   1170
         Width           =   7530
         Begin VB.Label Label50 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Unids. :"
            Height          =   240
            Left            =   2535
            TabIndex        =   59
            Top             =   495
            Width           =   780
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   3375
            TabIndex        =   58
            Top             =   495
            Width           =   825
         End
         Begin VB.Label Label48 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Facturas :"
            Height          =   240
            Left            =   150
            TabIndex        =   57
            Top             =   1065
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label Label47 
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   1200
            TabIndex        =   56
            Top             =   1065
            Visible         =   0   'False
            Width           =   2970
         End
         Begin VB.Label Label46 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Tot. Neto  :"
            Height          =   240
            Left            =   4425
            TabIndex        =   55
            Top             =   1365
            Width           =   1125
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   5580
            TabIndex        =   54
            Top             =   1365
            Width           =   1770
         End
         Begin VB.Label Label44 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Flete       :"
            Height          =   240
            Left            =   150
            TabIndex        =   53
            Top             =   795
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   1200
            TabIndex        =   52
            Top             =   795
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label Label42 
            BackColor       =   &H00C0FFC0&
            Caption         =   "IVA N/Ins.:"
            Height          =   240
            Left            =   4425
            TabIndex        =   51
            Top             =   1065
            Width           =   1125
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   5580
            TabIndex        =   50
            Top             =   1065
            Width           =   1770
         End
         Begin VB.Label Label40 
            BackColor       =   &H00C0FFC0&
            Caption         =   "IVA Ins.    :"
            Height          =   240
            Left            =   4425
            TabIndex        =   49
            Top             =   795
            Width           =   1125
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   5580
            TabIndex        =   48
            Top             =   795
            Width           =   1770
         End
         Begin VB.Label Label38 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Tot. Bruto :"
            Height          =   240
            Left            =   4425
            TabIndex        =   47
            Top             =   495
            Width           =   1125
         End
         Begin VB.Label Label37 
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   1155
            TabIndex        =   46
            Top             =   195
            Width           =   5820
         End
         Begin VB.Label Label36 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Items      :"
            Height          =   240
            Left            =   150
            TabIndex        =   45
            Top             =   495
            Width           =   1020
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   1215
            TabIndex        =   44
            Top             =   495
            Width           =   750
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   5580
            TabIndex        =   43
            Top             =   495
            Width           =   1770
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "  Cliente  "
         Height          =   1155
         Left            =   60
         TabIndex        =   32
         Top             =   15
         Width           =   7530
         Begin VB.Label Label33 
            BackColor       =   &H00C0FFC0&
            Caption         =   "CUIT :"
            Height          =   240
            Left            =   3675
            TabIndex        =   41
            Top             =   825
            Width           =   630
         End
         Begin VB.Label Label32 
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   4350
            TabIndex        =   40
            Top             =   825
            Width           =   1815
         End
         Begin VB.Label Label31 
            BackColor       =   &H00C0FFC0&
            Caption         =   "IVA :"
            Height          =   240
            Left            =   150
            TabIndex        =   39
            Top             =   825
            Width           =   525
         End
         Begin VB.Label Label30 
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   720
            TabIndex        =   38
            Top             =   825
            Width           =   1815
         End
         Begin VB.Label Label29 
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   5940
            TabIndex        =   37
            Top             =   525
            Width           =   1500
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   3675
            TabIndex        =   36
            Top             =   525
            Width           =   2175
         End
         Begin VB.Label Label27 
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   150
            TabIndex        =   35
            Top             =   525
            Width           =   3435
         End
         Begin VB.Label Label26 
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   1155
            TabIndex        =   34
            Top             =   225
            Width           =   5820
         End
         Begin VB.Label Label25 
            BackColor       =   &H00C0FFC0&
            Height          =   240
            Left            =   150
            TabIndex        =   33
            Top             =   225
            Width           =   840
         End
      End
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00F9FADC&
      Height          =   315
      ItemData        =   "CtaCte.frx":01B0
      Left            =   3960
      List            =   "CtaCte.frx":01C3
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   6060
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Cancel          =   -1  'True
      Caption         =   "Ver Saldos"
      Height          =   375
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5130
      Width           =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   11160
      TabIndex        =   15
      Top             =   5100
      Width           =   2310
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   11040
      ScaleHeight     =   4515
      ScaleWidth      =   2385
      TabIndex        =   2
      Top             =   390
      Width           =   2445
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ver Cheques"
         Height          =   315
         Left            =   90
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   4170
         Width           =   2250
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Tot. Debe :"
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
         Left            =   0
         TabIndex        =   23
         Top             =   3075
         Width           =   1005
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Tot. Haber :"
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
         Left            =   0
         TabIndex        =   22
         Top             =   3315
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         TabIndex        =   21
         Top             =   3060
         Width           =   1365
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Left            =   1005
         TabIndex        =   20
         Top             =   3315
         Width           =   1365
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         Caption         =   "SALDO :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   0
         TabIndex        =   19
         Top             =   3630
         Width           =   1005
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000002&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   990
         TabIndex        =   18
         Top             =   3630
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Sdo.Venc.:"
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
         Left            =   0
         TabIndex        =   17
         Top             =   3900
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   990
         TabIndex        =   16
         Top             =   3900
         Width           =   1365
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000013&
         Height          =   240
         Left            =   15
         TabIndex        =   14
         Top             =   1260
         Width           =   2265
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000013&
         Caption         =   "Teléfono"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   690
         TabIndex        =   13
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000013&
         Height          =   240
         Left            =   15
         TabIndex        =   12
         Top             =   705
         Width           =   2265
      End
      Begin VB.Label Label19 
         BackColor       =   &H80000013&
         Caption         =   "Contacto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   690
         TabIndex        =   11
         Top             =   465
         Width           =   975
      End
      Begin VB.Label Label18 
         BackColor       =   &H80000013&
         Height          =   240
         Left            =   1020
         TabIndex        =   10
         Top             =   150
         Width           =   1260
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000013&
         Caption         =   "Fec. Alta :"
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
         Left            =   30
         TabIndex        =   9
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000013&
         Height          =   240
         Left            =   1020
         TabIndex        =   8
         Top             =   2370
         Width           =   1275
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000013&
         Caption         =   "Val.Pend.:"
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
         Left            =   30
         TabIndex        =   7
         Top             =   2370
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000013&
         Height          =   240
         Left            =   1065
         TabIndex        =   6
         Top             =   2085
         Width           =   1230
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000013&
         Height          =   240
         Left            =   1050
         TabIndex        =   5
         Top             =   1860
         Width           =   1245
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000013&
         Caption         =   "U.Vta.Ctd.:"
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
         Left            =   30
         TabIndex        =   4
         Top             =   2085
         Width           =   1050
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000013&
         Caption         =   "U.Vta.C.C.:"
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
         Left            =   30
         TabIndex        =   3
         Top             =   1860
         Width           =   1035
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid GRILLA5 
      Height          =   4575
      Left            =   60
      TabIndex        =   24
      Top             =   420
      Width           =   10950
      _cx             =   19315
      _cy             =   8070
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
      BackColor       =   14089948
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"CtaCte.frx":0202
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
   Begin VB.PictureBox Picture4 
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
      Height          =   4575
      Left            =   30
      ScaleHeight     =   4515
      ScaleWidth      =   10935
      TabIndex        =   66
      Top             =   390
      Visible         =   0   'False
      Width           =   10995
      Begin VB.CommandButton Command7 
         Caption         =   "Cerrar"
         Height          =   270
         Left            =   9360
         TabIndex        =   67
         Top             =   4170
         Width           =   1185
      End
      Begin VSFlex8Ctl.VSFlexGrid GRILLA2 
         Height          =   4065
         Left            =   30
         TabIndex        =   68
         Top             =   30
         Width           =   10905
         _cx             =   19235
         _cy             =   7170
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
         FormatString    =   $"CtaCte.frx":0392
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
         Begin VB.PictureBox Picture5 
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
            Left            =   3060
            ScaleHeight     =   3240
            ScaleWidth      =   7680
            TabIndex        =   69
            Top             =   210
            Visible         =   0   'False
            Width           =   7740
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0FFFF&
               Caption         =   "  Cliente  "
               Height          =   1155
               Left            =   60
               TabIndex        =   91
               Top             =   15
               Width           =   7530
               Begin VB.Label Label78 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   150
                  TabIndex        =   100
                  Top             =   225
                  Width           =   840
               End
               Begin VB.Label Label77 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   1155
                  TabIndex        =   99
                  Top             =   225
                  Width           =   5820
               End
               Begin VB.Label Label76 
                  BackColor       =   &H00C0FFFF&
                  Height          =   240
                  Left            =   150
                  TabIndex        =   98
                  Top             =   525
                  Width           =   2865
               End
               Begin VB.Label Label75 
                  BackColor       =   &H00C0FFFF&
                  Height          =   240
                  Left            =   3105
                  TabIndex        =   97
                  Top             =   525
                  Width           =   2745
               End
               Begin VB.Label Label74 
                  BackColor       =   &H00C0FFFF&
                  Height          =   240
                  Left            =   5940
                  TabIndex        =   96
                  Top             =   525
                  Width           =   1500
               End
               Begin VB.Label Label73 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   720
                  TabIndex        =   95
                  Top             =   825
                  Width           =   1815
               End
               Begin VB.Label Label72 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "IVA :"
                  Height          =   240
                  Left            =   150
                  TabIndex        =   94
                  Top             =   825
                  Width           =   525
               End
               Begin VB.Label Label71 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   4350
                  TabIndex        =   93
                  Top             =   825
                  Width           =   1815
               End
               Begin VB.Label Label70 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "CUIT :"
                  Height          =   240
                  Left            =   3675
                  TabIndex        =   92
                  Top             =   825
                  Width           =   630
               End
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0FFFF&
               Caption         =   " Cheque  "
               Height          =   1695
               Left            =   60
               TabIndex        =   71
               Top             =   1170
               Width           =   7530
               Begin VB.Label Label69 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   5580
                  TabIndex        =   90
                  Top             =   525
                  Width           =   1770
               End
               Begin VB.Label Label68 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   1230
                  TabIndex        =   89
                  Top             =   525
                  Width           =   2970
               End
               Begin VB.Label Label67 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Banco :"
                  Height          =   240
                  Left            =   150
                  TabIndex        =   88
                  Top             =   525
                  Width           =   1020
               End
               Begin VB.Label Label66 
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
                  TabIndex        =   87
                  Top             =   195
                  Width           =   1770
               End
               Begin VB.Label Label65 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Ingreso :"
                  Height          =   240
                  Left            =   4425
                  TabIndex        =   86
                  Top             =   525
                  Width           =   1125
               End
               Begin VB.Label Label64 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   5580
                  TabIndex        =   85
                  Top             =   795
                  Width           =   1770
               End
               Begin VB.Label Label63 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Emisión :"
                  Height          =   240
                  Left            =   4425
                  TabIndex        =   84
                  Top             =   795
                  Width           =   1125
               End
               Begin VB.Label Label62 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   5580
                  TabIndex        =   83
                  Top             =   1065
                  Width           =   1770
               End
               Begin VB.Label Label61 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Vencimto.:"
                  Height          =   240
                  Left            =   4425
                  TabIndex        =   82
                  Top             =   1065
                  Width           =   1125
               End
               Begin VB.Label Label60 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   1230
                  TabIndex        =   81
                  Top             =   795
                  Width           =   2985
               End
               Begin VB.Label Label59 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Concepto :"
                  Height          =   240
                  Left            =   150
                  TabIndex        =   80
                  Top             =   795
                  Width           =   1020
               End
               Begin VB.Label Label58 
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
                  TabIndex        =   79
                  Top             =   1365
                  Width           =   1830
               End
               Begin VB.Label Label57 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Importe :"
                  Height          =   240
                  Left            =   2445
                  TabIndex        =   78
                  Top             =   1365
                  Width           =   1125
               End
               Begin VB.Label Label56 
                  BackColor       =   &H00C0FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   1230
                  TabIndex        =   77
                  Top             =   1065
                  Width           =   2970
               End
               Begin VB.Label Label55 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Datos Ad.:"
                  Height          =   240
                  Left            =   150
                  TabIndex        =   76
                  Top             =   1065
                  Width           =   1005
               End
               Begin VB.Label Label54 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Recibo :"
                  Height          =   240
                  Left            =   4440
                  TabIndex        =   75
                  Top             =   180
                  Width           =   1125
               End
               Begin VB.Label Label53 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   1  'Fixed Single
                  Height          =   240
                  Left            =   5595
                  TabIndex        =   74
                  Top             =   180
                  Width           =   1770
               End
               Begin VB.Label Label52 
                  Alignment       =   2  'Center
                  BackColor       =   &H00C0FFFF&
                  Height          =   240
                  Left            =   3240
                  TabIndex        =   73
                  Top             =   180
                  Width           =   1290
               End
               Begin VB.Label Label51 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "Número :"
                  Height          =   240
                  Left            =   150
                  TabIndex        =   72
                  Top             =   210
                  Width           =   1020
               End
            End
            Begin VB.CommandButton Command9 
               Caption         =   "Cerrar"
               Height          =   270
               Left            =   6270
               TabIndex        =   70
               Top             =   2925
               Width           =   1185
            End
         End
      End
      Begin VB.Label Label80 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Left            =   5130
         TabIndex        =   103
         Top             =   4170
         Width           =   2415
      End
      Begin VB.Label Label79 
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
         Height          =   300
         Left            =   7590
         TabIndex        =   102
         Top             =   4170
         Width           =   1635
      End
   End
   Begin VB.Label Label81 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   105
      Top             =   90
      Width           =   750
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
      Height          =   270
      Left            =   3540
      TabIndex        =   1
      Top             =   90
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
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
      Left            =   2580
      TabIndex        =   0
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "CtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim VeNotaC As Boolean, HaySaldo As Boolean
Dim Ubi3X As Double, Ubi3Y As Double, SeMueve3 As Boolean
Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset
Dim RgCCTE As New ADODB.Recordset, RgFCIVA As New ADODB.Recordset
Dim RgCHEQ As New ADODB.Recordset, RgCYB As New ADODB.Recordset
Dim Cpbte As Double, TIPO As String * 2, Letra As String * 1, PtoVtaC As Integer
Dim SdoInicDb As Currency, SdoInicCr As Currency

Dim TOTCHEQ As Double

Private Sub Combo1_Click()

DoVer3

End Sub

Private Sub Command7_Click()

    Picture4.Visible = False

End Sub

Private Sub Command8_Click()

DeDonde = 3
BusClte.Show
BusClte.SetFocus

End Sub

Private Sub Command9_Click()

    Picture5.Visible = False

End Sub

Private Sub Form_Load()

Me.Move 0, 0, 13800, 6000

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

VeNotaC = False

CargaAnos

DoVer3

End Sub

Private Sub Form_Activate()

If VeNotaC Then Exit Sub

VeNotaC = True

miSQL2 = "SELECT * FROM Notaclte WHERE CLTE = " & CodCLTE
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
   
If Not RgTABL.EOF Then
    RgTABL.Close
    NOTACLTE.Show
    Exit Sub
End If

RgTABL.Close

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button < 2 Then Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)

'FCMENU.Toolbar2.Buttons(2).Value = tbrUnpressed

dbCLTE.Close

End Sub

Private Sub FECHA_Change()

DoVer3

End Sub

Private Sub Command1_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Command2_Click()

Frame1.Visible = True
DoVer8

End Sub

Private Sub Command3_Click()

' RgCLTE.Close
'RgCCTE.Close

Frame1.Visible = False

End Sub

Private Sub Command5_Click()

    Picture3.Visible = False

End Sub

Private Sub Command6_Click()

Dim Cpbte As Double, TIPO As Integer, Letra As String * 1, PtoVtaC1 As Integer
Dim FILA As Integer, ElTotal As Currency, i3%

FILA = GRILLA5.Row
TIPO = GRILLA5.TextMatrix(FILA, 11)
PtoVtaC1 = GRILLA5.TextMatrix(FILA, 12)
Cpbte = GRILLA5.TextMatrix(FILA, 2)

miSQL = " CPBTE = " & Cpbte & " AND TIPO = " & TIPO & _
        " AND PTOVTA = " & PtoVtaC1
miSQL = "Select * FROM FCEstad1 WHERE " & miSQL
RgFCIVA.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgFCIVA.EOF Then
    RgFCIVA.Close
    Exit Sub
End If
   
FG2.Clear flexClearScrollable

For i3 = (FG2.Rows - 1) To 1 Step -1
FG2.RemoveItem (i3)
Next i3

'Set FG2.DataSource = RGFCIVA

Do Until RgFCIVA.EOF

i3 = i3 + 1

FG2.AddItem RgFCIVA!COD1 & vbTab & RgFCIVA!Pulg & vbTab & RgFCIVA!Mtr & vbTab & _
            RgFCIVA!Milim & vbTab & RgFCIVA!Telas & vbTab & RgFCIVA!Cant & _
            vbTab & RgFCIVA!PEsp & vbTab & RgFCIVA!PCos & _
            vbTab & RgFCIVA!PVta & vbTab & RgFCIVA!Impte, FG2.Rows
    
paraSTATUS = "Cargando : " & (FG2.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

RgFCIVA.MoveNext

Loop

FG2.AutoSize 0, 5

    Picture2.Visible = True

RgFCIVA.Close

End Sub

Private Sub Command4_Click()

    Picture2.Visible = False
    
End Sub

Private Sub Command10_Click()

DoVer9

End Sub

Private Sub FG8_Click()

Frame1.Visible = False
CodCLTE = FG8.TextMatrix(FG8.Row, 0)
ClteNomb = FG8.TextMatrix(FG8.Row, 1)

DoVer3

End Sub

Sub DoVer3()
Dim PLIS As Double, DEBE As Currency, HABER As Currency, Saldo As Currency

Me.Label2.Caption = CodCLTE & "  " & ClteNomb

DEBE = 0: HABER = 0: Saldo = 0

  If ORDENADOx <> "" Then
         Clasifica = elORDEN & ORDENADOx & ORDENADOad
  Else
         Clasifica = ""
End If


'///           ----------------------------------------
SdoInicDb = 0: SdoInicCr = 0
     
miSQL = "SELECt sum(iif(tipo=1or tipo=3,impte,0)) as saldeb, " & _
        "sum(iif(tipo=4or tipo=2,impte,0)) as salcred " & _
        "from CtasCtes " & _
        "WHERE CLTE = " & CodCLTE & " and year(fecha)<" & Combo1.ItemData(Combo1.ListIndex)
RgCCTE.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgCCTE.EOF Then
    RgCCTE.Close
    Exit Sub
    End If
    
SdoInicDb = IIf(IsNull(RgCCTE!saldeb), 0, RgCCTE!saldeb)
SdoInicCr = IIf(IsNull(RgCCTE!salcred), 0, RgCCTE!salcred)

    RgCCTE.Close
    
'///           ----------------------------------------
     
     
     
miSQL = "SELECT * FROM CtasCtes WHERE CLTE = " & CodCLTE & _
        " and YEAR(FECHA) >= " & Combo1.ItemData(Combo1.ListIndex) & _
        " ORDER BY FECHA, CPBTE, TIPO"
RgCCTE.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgCCTE.EOF Then
    RgCCTE.Close
    Exit Sub
    End If


' paraSTATUS = "Seleccionados : " & RgCCTE.RecordCount
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

       LimpiaGrilla
       CargaGrilla

  
        DEBE = GRILLA5.Aggregate(flexSTSum, _
        1, 8, (GRILLA5.Rows - 1), 8)
        HABER = GRILLA5.Aggregate(flexSTSum, _
        1, 9, (GRILLA5.Rows - 1), 9)
        Saldo = GRILLA5.Aggregate(flexSTSum, _
        1, 7, (GRILLA5.Rows - 1), 7)
        Saldo = DEBE - HABER
        
GRILLA5.TopRow = GRILLA5.Rows

Label7.Caption = Format(Saldo, "###,##0.00 ")
Label5.Caption = Format(DEBE, "###,##0.00 ")
Label9.Caption = Format(HABER, "###,##0.00 ")

End Sub

Sub LimpiaGrilla()
Dim I1%
   
GRILLA5.Clear flexClearScrollable

For i = (GRILLA5.Rows - 1) To 1 Step -1
GRILLA5.RemoveItem (i)
Next i

End Sub

Sub CargaGrilla()

Dim DEBE          As String
Dim HABER         As String
Dim Saldo         As Double
Dim SALDOed       As String
Dim RESTO         As String
Dim Cpbte         As Single
Dim Im1            As Single
Dim Im2            As Single
Dim Im3            As Single
Dim Im4            As Single
Dim Im5            As Single
Dim Im6            As Single
Dim FECHA         As Date
Dim TipoCpbt      As String

Dim I1, i2, Tc
Dim i3 As Long

GRILLA5.Col = 0: i3 = 1

FECHA = "31-12-" & Combo1.ItemData(Combo1.ListIndex) - 1
GRILLA5.AddItem FECHA & vbTab & "Sdo.Ant." & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & _
            "" & vbTab & "" & vbTab & "" & vbTab & _
             SdoInicDb & vbTab & SdoInicCr & vbTab & _
            SdoInicDb - SdoInicCr & vbTab & 0, GRILLA5.Rows

Saldo = SdoInicDb - SdoInicCr

Do Until RgCCTE.EOF

i3 = i3 + 1

Tc = RgCCTE!TIPO
FECHA = RgCCTE!FECHA
Cpbte = RgCCTE!Cpbte
Im1 = Val(RgCCTE!imput1)
Im2 = Val(RgCCTE!imput2)
Im3 = Val(RgCCTE!imput3)
Im4 = Val(RgCCTE!imput4)
Im5 = Val(RgCCTE!imput5)

DEBE = 0: HABER = 0: RESTO = 0
If Tc = 0 Or Tc = 1 Or Tc = 3 Or Tc = 7 Then
    DEBE = Format(RgCCTE!Impte, "$ ###,##0.00")
    HABER = ""
    Saldo = Saldo + DEBE
Else
    DEBE = ""
    HABER = Format(RgCCTE!Impte, "$ ###,##0.00")
    Saldo = Saldo - HABER
End If

RESTO = ""
If Tc = 1 Or Tc = 3 Then
    If RgCCTE!DEBE = 0 Then
        RESTO = "  ---      "
    Else
        RESTO = Format(RgCCTE!DEBE, "$ ###,##0.00")
    End If
End If
   
   
Select Case RgCCTE!TIPO
    
    Case 0:        TipoCpbt = "Sdo.A."
    Case 1:        TipoCpbt = "Fact."
    Case 2:        TipoCpbt = "NCréd."
    Case 3:        TipoCpbt = "NDéb."
    Case 4:        TipoCpbt = "Recibo"
    Case 5:        TipoCpbt = "PACta."
    Case 6:        TipoCpbt = "Descto"
    Case 7:        TipoCpbt = "RecAn."
    Case 8:        TipoCpbt = "NCInt"
    Case 9:        TipoCpbt = "Anul."

End Select

SALDOed = Format(Saldo, "$ ###,##0.00")
GRILLA5.AddItem FECHA & vbTab & TipoCpbt & vbTab & Cpbte & vbTab & Im1 & vbTab & Im2 & vbTab & _
            Im3 & vbTab & Im4 & vbTab & RESTO & vbTab & _
            DEBE & vbTab & HABER & vbTab & SALDOed & vbTab & _
            RgCCTE!TIPO & vbTab & RgCCTE!prefijo, GRILLA5.Rows
    
paraSTATUS = "Cargando : " & (GRILLA5.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

'Me.Show
'DoEvents

Salta:

RgCCTE.MoveNext
Loop

    RgCCTE.Close

GRILLA5.AutoSize 0, 10

If GRILLA5.Rows > 1 Then GRILLA5.Row = 1


' FG1.SetFocus

End Sub

'******************************************************************
'               C H E Q U E S
'******************************************************************

Private Sub DoVer9()
Dim CantReg As Integer

GRILLA2.Clear flexClearScrollable

For i = (GRILLA2.Rows - 1) To 1 Step -1
GRILLA2.RemoveItem (i)
Next i

If RgCHEQ.State = 1 Then Exit Sub
          
miSQL = "SELECT count(*) as CanReg FROM Cheques WHERE CLTE = " & CodCLTE & _
         " and ESTADO = '1'"
        
RgCHEQ.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly


If RgCHEQ.EOF Then
    RgCHEQ.Close
    Exit Sub
End If

CantReg = RgCHEQ!CanReg
RgCHEQ.Close
    
paraSTATUS = "Seleccionados : " & CantReg
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

miSQL = "SELECT * FROM Cheques WHERE CLTE = " & CodCLTE & _
         " and ESTADO = '1'"

RgCHEQ.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly
If RgCHEQ.EOF Then
    RgCHEQ.Close
    Exit Sub
End If
       
Picture4.Visible = True
       
       CargaCheq

  Label79.Caption = Format(TOTCHEQ, "##,##0.00")
  Label80.Caption = "Hay " & CantReg & " cheques, por "
    
    RgCHEQ.Close


End Sub

Sub CargaCheq()

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

Do Until RgCHEQ.EOF Or i3 >= 50

i3 = i3 + 1

LaFECHA = RgCHEQ!FECVTO
RSet CLTE = RgCHEQ!CLTE


miSQL = "Select nomb FROM Clientes WHERE CODIGO = " & RgCHEQ!CLTE

RgCLTE.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly
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

MuestraCheq

End Sub

Private Sub MuestraCheq()
Dim FILA As Integer, ElTotal As Currency
Dim ElCheque As String * 10, I4%

FILA = GRILLA2.Row
ElCheque = GRILLA2.TextMatrix(FILA, 0)

miSQL = "Select * FROM Cheques WHERE NROCHEQ = " & ElCheque

RgCHEQ.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly
If RgCHEQ.EOF Then
    RgCHEQ.Close
    Exit Sub
End If

Label1.Caption = RgCHEQ!CLTE


Label66.Caption = GRILLA2.TextMatrix(FILA, 0)
Label58.Caption = Format(RgCHEQ!Importe, "$ ###,##0.00")
Label52.Caption = ""
If RgCHEQ!orden Then Label52.Caption = "N/Orden"

Label53.Caption = Format(RgCHEQ!cpbing, "####,##0")
Label68.Caption = RgCHEQ!BcoSuc & " - *** No Existe en Tabla *** "
Label69.Caption = Format(RgCHEQ!FECING, "ddd dd mmm yyyy")
Label64.Caption = Format(RgCHEQ!fecemi, "ddd dd mmm yyyy")
Label62.Caption = Format(RgCHEQ!FECVTO, "ddd dd mmm yyyy")
Label60.Caption = RgCHEQ!concep
Label56.Caption = RgCHEQ!datosad

miSQL1 = "Select * FROM Bancos where COD = " & RgCHEQ!BcoSuc
RgCYB.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

    If Not RgCYB.EOF Then
         Label68.Caption = RgCHEQ!BcoSuc & " - " & RgCYB!NOMBRE
        End If
         
RgCYB.Close

      Picture5.Visible = True

miSQL1 = "Select * FROM Clientes WHERE CODIGO = " & RgCHEQ!CLTE
RgCLTE.Open miSQL1, dbCLTE, adOpenForwardOnly, adLockReadOnly
If RgCLTE.EOF Then
    RgCLTE.Close
    GoTo SaltaClte
End If

Label78.Caption = RgCLTE!CODIGO
Label77.Caption = RgCLTE!Nomb
Label76.Caption = RgCLTE!Dir
Label75.Caption = "(" & RgCLTE!CP & ") " & RgCLTE!Loc

I4 = Asc(RgCLTE!PCIA)
Label74.Caption = Mid(PCIA(I4 - 65), 2, 10)

Label73.Caption = Combo2.List(RgCLTE!CIVA - 1)
Label71.Caption = RgCLTE!Cuit

SaltaClte:

RgCLTE.Close
RgCHEQ.Close

End Sub
'******************************************************************
'               S A L D O S
'******************************************************************

Private Sub DoVer8()
Dim Saldo As Currency

FG8.Enabled = False
     
miSQL = "SELECT CODIGO, NOMB FROM Clientes  ORDER BY NOMB"
RgCLTE.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgCLTE.EOF Then
    RgCLTE.Close
    Exit Sub
    End If

Me.Show


' paraSTATUS = "Seleccionados : " & RgCCTE.RecordCount
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

       LimpiaFG8

FG8.Col = 0

Do Until RgCLTE.EOF

DoEvents
       
CargaFG8
    
paraSTATUS = "Cargando : " & (FG8.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

RgCLTE.MoveNext

If HaySaldo Then
        Saldo = FG8.Aggregate(flexSTSum, _
        1, 2, (FG8.Rows - 1), 2)
        Label24.Caption = Format(Saldo, " $ ###,##0.00 ")
End If

HaySaldo = False

Loop

RgCLTE.Close

FG8.AutoSize 0, 2

If FG8.Rows > 1 Then FG8.Row = 1

FG8.Enabled = True

End Sub

Sub CargaFG8()
Dim Cliente       As String
Dim Saldo         As Double

Dim I1, i2, Tc
Dim i3 As Long
     
Saldo = 0

miSQL = "SELECT * FROM CtasCtes WHERE CLTE = " & RgCLTE!CODIGO & _
        " ORDER BY FECHA"
RgCCTE.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgCCTE.EOF Then
    RgCCTE.Close
    Exit Sub
    End If


Do Until RgCCTE.EOF


If RgCCTE!TIPO = 0 Or RgCCTE!TIPO = 1 Or RgCCTE!TIPO = 3 Or RgCCTE!TIPO = 7 Then
    Saldo = Saldo + RgCCTE!Impte
Else
    Saldo = Saldo - RgCCTE!Impte
End If

RgCCTE.MoveNext

Loop
       
FG8.AddItem RgCLTE!CODIGO & vbTab & RgCLTE!Nomb & vbTab & Format(Saldo, " $ ###,##0.00 "), FG8.Rows

HaySaldo = True

    RgCCTE.Close

End Sub

Sub LimpiaFG8()
Dim I1%
   
FG8.Clear flexClearScrollable

For i = (FG8.Rows - 1) To 1 Step -1
FG8.RemoveItem (i)
Next i

End Sub

Private Sub GRILLA5_Click()

If GRILLA5.Rows < 1 Then Exit Sub

MuestraDetalle

End Sub

Private Sub MuestraDetalle()
Dim FILA As Integer, ElTotal As Currency, I4%, ElPtoVta As String * 4, ElCpbte As String * 8

FILA = GRILLA5.Row
TIPO = GRILLA5.TextMatrix(FILA, 11)
PtoVtaC = GRILLA5.TextMatrix(FILA, 12)
Cpbte = GRILLA5.TextMatrix(FILA, 2)

miSQL = "Select * FROM FCIVAVTA WHERE  CPBTE = " & Cpbte & " AND TIPO = '" & _
        TIPO & "' AND CLTE = " & CodCLTE & " AND PTOVTA = " & PtoVtaC

RgFCIVA.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly
If RgFCIVA.EOF Then
    RgFCIVA.Close
    Exit Sub
End If

ElPtoVta = String(4 - Len(RgFCIVA!PtoVta), "0") & RgFCIVA!PtoVta
ElCpbte = String(8 - Len(RgFCIVA!Cpbte), "0") & RgFCIVA!Cpbte

Label37.Caption = GRILLA5.TextMatrix(FILA, 1) & "  " & RgFCIVA!Letra & "  -  " & _
                   ElPtoVta & " - " & ElCpbte
Label35.Caption = RgFCIVA!ITEMS

ElTotal = RgFCIVA!GRINS + RgFCIVA!IVAIns + RgFCIVA!IVANOINS

Label34.Caption = Format(RgFCIVA!GRINS, "###,##0.00")
Label39.Caption = Format(RgFCIVA!IVAIns, "###,##0.00")
Label41.Caption = Format(RgFCIVA!IVANOINS, "###,##0.00")
Label45.Caption = Format(ElTotal, "###,##0.00")

'Label21.Caption = Format(RgFCIVA!flete, "##,##0.00")
Label49.Caption = RgFCIVA!TotCan

Label25.Caption = RgFCIVA!CLTE
Label26.Caption = RgFCIVA!Nomb

      Picture3.Visible = True

miSQL = " CPBTE = " & Cpbte & " AND TIPO = " & TIPO & " AND LETRA = '" & _
        Letra & "' AND PTOVTA = " & PtoVtaC
miSQL = "Select * FROM Clientes WHERE Codigo = " & RgFCIVA!CLTE

RgCLTE.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly
If RgCLTE.EOF Then
    GoTo SALIR
End If

Label25.Caption = RgCLTE!CODIGO
Label26.Caption = RgCLTE!Nomb
Label27.Caption = RgCLTE!Dir
Label28.Caption = "(" & RgCLTE!CP & ") " & RgCLTE!Loc

I4 = Asc(RgCLTE!PCIA)
Label29.Caption = Mid(PCIA(I4 - 65), 2, 10)

Label30.Caption = Combo2.List(RgCLTE!CIVA - 1)
Label32.Caption = RgCLTE!Cuit

SALIR:

RgCLTE.Close
RgFCIVA.Close

End Sub


'--------------------------------------------
'    Movimiento de Ventanas

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = True
Ubi3X = x
Ubi3Y = y

End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If SeMueve3 Then
    Picture3.Left = Picture3.Left - (Ubi3X - x)
    Picture3.top = Picture3.top - (Ubi3Y - y)
End If

End Sub

Private Sub Picture3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = False

End Sub

Private Sub Frame2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = True
Ubi3X = x
Ubi3Y = y

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If SeMueve3 Then
    Picture3.Left = Picture3.Left - (Ubi3X - x)
    Picture3.top = Picture3.top - (Ubi3Y - y)
End If

End Sub

Private Sub Frame2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = False

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = True
Ubi3X = x
Ubi3Y = y

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If SeMueve3 Then
    Picture2.Left = Picture2.Left - (Ubi3X - x)
    Picture2.top = Picture2.top - (Ubi3Y - y)
End If

End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = False

End Sub

Private Sub FG2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = True
Ubi3X = x
Ubi3Y = y

End Sub

Private Sub FG2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If SeMueve3 Then
    Picture2.Left = Picture2.Left - (Ubi3X - x)
    Picture2.top = Picture2.top - (Ubi3Y - y)
End If

End Sub

Private Sub FG2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = False

End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = True
Ubi3X = x
Ubi3Y = y

End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If SeMueve3 Then
    Picture5.Left = Picture5.Left - (Ubi3X - x)
    Picture5.top = Picture5.top - (Ubi3Y - y)
End If

End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

SeMueve3 = False

End Sub

Private Sub CargaAnos()
Dim A1 As Integer, Ano1 As Integer

A1 = 0
Ano1 = 2005

Do Until Ano1 > Year(Date)

        Combo1.AddItem Ano1
        Combo1.ItemData(A1) = Ano1

A1 = A1 + 1
Ano1 = Ano1 + 1

Loop

Combo1.ListIndex = A1 - 5

End Sub
