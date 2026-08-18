VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Listados 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Listados Varios"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8490
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
   ScaleHeight     =   6405
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "A Excel"
      Height          =   345
      Left            =   2130
      TabIndex        =   42
      Top             =   4380
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Caption         =   "a PDF"
      Height          =   375
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4350
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00F9FADC&
      Height          =   315
      ItemData        =   "Listados.frx":0000
      Left            =   270
      List            =   "Listados.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   4920
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00F9FADC&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Listados.frx":0052
      Left            =   2880
      List            =   "Listados.frx":007D
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   5910
      Visible         =   0   'False
      Width           =   2400
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
      ItemData        =   "Listados.frx":00E6
      Left            =   5340
      List            =   "Listados.frx":00E8
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   5910
      Visible         =   0   'False
      Width           =   2310
   End
   Begin MSAdodcLib.Adodc CLTE 
      Height          =   375
      Left            =   810
      Top             =   5400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=FCMENU"
      OLEDBString     =   "DSN=FCMENU"
      OLEDBFile       =   ""
      DataSourceName  =   "Alestel"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "CLTE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6300
      TabIndex        =   2
      Top             =   4350
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Emite"
      Height          =   375
      Left            =   4980
      TabIndex        =   1
      Top             =   4350
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc FCIVA 
      Height          =   375
      Left            =   2880
      Top             =   5370
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=FCMENU"
      OLEDBString     =   "DSN=FCMENU"
      OLEDBFile       =   ""
      DataSourceName  =   "Alestel1"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "FCIVA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc ARTIC 
      Height          =   375
      Left            =   4920
      Top             =   5370
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=FCMENU"
      OLEDBString     =   "DSN=FCMENU"
      OLEDBFile       =   ""
      DataSourceName  =   "Alestel1"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ARTIC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc CCTE 
      Height          =   375
      Left            =   6930
      Top             =   5370
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=FCMENU"
      OLEDBString     =   "DSN=FCMENU"
      OLEDBFile       =   ""
      DataSourceName  =   "Alestel1"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "CCTE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc TABLA 
      Height          =   375
      Left            =   870
      Top             =   5850
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=FCMENU"
      OLEDBString     =   "DSN=FCMENU"
      OLEDBFile       =   ""
      DataSourceName  =   "Alestel1"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "TABLA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid FG2 
      Height          =   3765
      Left            =   600
      TabIndex        =   43
      Top             =   5820
      Visible         =   0   'False
      Width           =   10485
      _cx             =   18494
      _cy             =   6641
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Listados.frx":00EA
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
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   660
      TabIndex        =   3
      Top             =   810
      Visible         =   0   'False
      Width           =   7275
      Begin VB.Frame Frame4 
         Caption         =   "  Opciones   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1410
         Left            =   1860
         TabIndex        =   23
         Top             =   900
         Visible         =   0   'False
         Width           =   5310
         Begin VB.ComboBox Combo6 
            BackColor       =   &H00F9FADC&
            Height          =   315
            ItemData        =   "Listados.frx":027D
            Left            =   1110
            List            =   "Listados.frx":027F
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   540
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.ComboBox Combo5 
            BackColor       =   &H00F9FADC&
            Height          =   315
            ItemData        =   "Listados.frx":0281
            Left            =   1125
            List            =   "Listados.frx":0283
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   540
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00F9FADC&
            Height          =   315
            Left            =   2700
            MaxLength       =   3
            TabIndex        =   35
            Text            =   "30"
            Top             =   960
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00F9FADC&
            Height          =   315
            ItemData        =   "Listados.frx":0285
            Left            =   1140
            List            =   "Listados.frx":0287
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   540
            Visible         =   0   'False
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker FECHA 
            Height          =   330
            Left            =   1140
            TabIndex        =   34
            Top             =   180
            Visible         =   0   'False
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
            Format          =   160104451
            CurrentDate     =   36877
         End
         Begin VB.Label Label4 
            Height          =   270
            Left            =   90
            TabIndex        =   38
            Top             =   225
            Visible         =   0   'False
            Width           =   4965
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Días a Vencer :"
            Height          =   270
            Left            =   990
            TabIndex        =   36
            Top             =   1020
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label2 
            Height          =   270
            Left            =   180
            TabIndex        =   27
            Top             =   570
            Visible         =   0   'False
            Width           =   4965
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "  Selección   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1410
         Left            =   90
         TabIndex        =   30
         Top             =   900
         Visible         =   0   'False
         Width           =   1680
         Begin VB.OptionButton Option3 
            Caption         =   "Un Día"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   450
            Width           =   1335
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Mes Completo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   780
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "  Selección   "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1410
         Left            =   105
         TabIndex        =   19
         Top             =   885
         Visible         =   0   'False
         Width           =   1680
         Begin VB.OptionButton Option2 
            Caption         =   "Una Zona"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   630
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Uno Sólo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Todos "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2880
         TabIndex        =   24
         Top             =   1350
         Width           =   3135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   1575
         TabIndex        =   5
         Top             =   330
         Width           =   4245
      End
      Begin VB.Label Label8 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   330
         Left            =   330
         TabIndex        =   4
         Top             =   3015
         Width           =   4245
      End
   End
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
      Height          =   3600
      Left            =   1050
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   6495
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Subdiario Vtas (Comis.)"
         Height          =   375
         Index           =   15
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1440
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comis.x Cobr.(nvo.)"
         Height          =   375
         Index           =   14
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2490
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldos de Cta. Cte."
         Height          =   375
         Index           =   12
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2100
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado de Cuenta"
         Height          =   375
         Index           =   11
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1710
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Clientes"
         Height          =   375
         Index           =   1
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   270
         Value           =   -1  'True
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Planilla de Cobranzas"
         Height          =   375
         Index           =   10
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1320
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Deuda Pendiente"
         Height          =   375
         Index           =   8
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   510
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Deuda Vencida"
         Height          =   375
         Index           =   9
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   915
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comisiones por Cobr."
         Height          =   375
         Index           =   7
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3000
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Percepciones ARBA"
         Height          =   375
         Index           =   13
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2880
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Comisiones por Ventas"
         Height          =   375
         Index           =   6
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2610
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ingresos Brutos"
         Height          =   375
         Index           =   5
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2220
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Subd.de Cobranzas AFIP"
         Height          =   375
         Index           =   4
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1830
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Subdiario de Ventas"
         Height          =   375
         Index           =   3
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1050
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Lista de Precios"
         Height          =   375
         Index           =   2
         Left            =   495
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   675
         Width           =   2500
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FG3 
      Height          =   1665
      Left            =   30
      TabIndex        =   46
      Top             =   6090
      Visible         =   0   'False
      Width           =   5235
      _cx             =   9234
      _cy             =   2937
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Listados.frx":0289
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
   Begin VSPDF8LibCtl.VSPDF8 VSPDF81 
      Left            =   300
      Top             =   270
      Author          =   ""
      Creator         =   ""
      Title           =   ""
      Subject         =   ""
      Keywords        =   ""
      Compress        =   3
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Listados Varios"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   2040
      TabIndex        =   18
      Top             =   300
      Width           =   4245
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000009&
      Caption         =   "(Impr.Matriz)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   225
      Left            =   6360
      TabIndex        =   17
      Top             =   390
      Width           =   1440
   End
End
Attribute VB_Name = "Listados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LaOpcion As Integer, TodosCLTE As Boolean
Dim Tit1 As String, Tit2 As String, Tit3 As String, Tit4 As String, Tit5 As String
Dim TitTit As String, TitDet As String, TitTot As String, NroHoja As Integer
Dim Asteriscos As String, Rayitas As String, lineas As Integer
Dim PorCompr As Integer, Ultcompr As Integer
Dim OptDos%, OptTres%, TotDebeP As Currency, TotDebeT As Currency, TotVenc As Currency
Dim STRImpre As String * 20, TotVencT As Currency
Dim TotItems As Integer, TotPed As Long, Rengln As Double

Dim s$, h$, d$, Cont As Integer

Dim dbCLTE As New ADODB.Connection
Dim RgCCTE As New ADODB.Recordset
Dim RgCLTE As New ADODB.Recordset
Dim RgFCIVA As New ADODB.Recordset
Dim RgART As New ADODB.Recordset

Private Sub Combo2_Click()

If Text1.Visible Then Text1.SetFocus

End Sub

'------------------------------------------------------------
'   1-Clientes                 8-Deuda Pendiente
'   2-Lista de Precios         9-Deuda Vencida
'   3-Subd. de Ventas         10-Planilla de Cobza.
'   4-Subd. de Cobrzas.       11-Estado de Cuenta
'   5-Ing. Brutos             12-Saldos de Cta.Cte.
'   6-Comisiones x Vtas.      13-Percepciones ARBA
'   7-Comisiones x Cobrzas.   14-Comisiones x Cobrzas. NUEVO
'------------------------------------------------------------

Private Sub Form_Load()

Me.Move 1000, 800, 8610, 5300

FCMENU.Toolbar2.Buttons(7).Value = tbrPressed

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

Option1(1) = False

Frame1.Visible = True
Frame2.Visible = False
Command2.Visible = False

CargaZona
CargaVend
CargaCVta
CargaPcia

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbCLTE.Close
FCMENU.Toolbar2.Buttons(7).Value = tbrUnpressed

End Sub

Private Sub Command1_Click()

If Frame2.Visible Then
    Frame2.Visible = False
    Frame1.Visible = True
    Command3.Visible = False

Else
 '   Unload Me

    Call SlideWindow(Me, 50)
End If

End Sub

Private Sub Command2_Click()
Dim Mensaje As String
                

Mensaje = "Verifique que esta todo en condiciones    " & vbCrLf & "                  para Imprimir ..." _
            & vbCrLf & vbCrLf & "                   ... continúa ?"
Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "Modificación de Precios")

If Respuesta = vbNo Then GoTo AlFinal
If Respuesta = vbCancel Then
         Exit Sub
         End If

Tit5 = ""

Asteriscos = String(80, "*")
Rayitas = String(80, "-")
NroHoja = 0

'  impresión
Select Case LaOpcion
    
     Case 1
            LisClientes
     Case 2
            LisPrecios
     Case 3
            LisSubVtas
     Case 4
            LisSubCobzasAFIP
     Case 5
            LisIngBrutos
     Case 6
            LisComisVtas
     Case 7
            LisComisCobra
     Case 8
            LisLisFact
     Case 9
            LisDeudaVenc
     Case 10
            LisPlaniCobra
     Case 11
            LisCCTE
     Case 12
            LisSaldos
     Case 13
            LisPerARBA
     Case 14
            LisComCobNVO
     Case 15
            LisSubVtasCom

End Select

AlFinal:
  
Frame1.Visible = True
Frame2.Visible = False
Command3.Visible = False
Command2.Visible = False
  
End Sub

Private Sub Command3_Click()

Select Case LaOpcion
    Case 10: LisPlaniCobraPDF
    Case 14: LisComCobPDF
End Select

Command2.Visible = False

End Sub

Private Sub Option1_Click(Index As Integer)

LaOpcion = Index

Option1(Index) = False

Frame1.Visible = False
Frame2.Visible = True

Check1.Visible = False

Command2.Visible = True

Select Case Index

     Case 1
            SelClientes
     Case 2
            SelPrecios
     Case 3
            SelSubVtas  '
     Case 4
            SelSubCobzasAFIP   '
     Case 5
            SelIngBrutos
     Case 6
            SelComisVtas
     Case 7
            SelComisCobra
     Case 8
            SelLisFact
     Case 9
            SelDeudaVenc
     Case 10
            SelPlaniCobra
     Case 11
            SelCCTE
     Case 12
            SelSaldos
     Case 13
            SelPerARBA    '
     Case 14
            SelComCobNVO
     Case 15
            SelSubVtasCom  '

End Select

End Sub

Private Sub Option2_Click(Index As Integer)

OptDos = Index
TodosCLTE = False

Frame3.Visible = True
Frame4.Visible = True

Combo2.Visible = False
Combo5.Visible = False
    
Select Case Index

    Case 0
        If LaOpcion <> 1 Then
            DeDonde = 4
            BusClte.Show
        Else
            Label2.Visible = True
            Label2.Caption = "Provincia :"
            Combo6.Visible = True
        End If
    Case 1
        Label2.Visible = True
        Label2.Caption = "Zona :"
        Combo2.Visible = True
    Case 2
        If LaOpcion <> 10 Then Frame4.Visible = False
        TodosCLTE = True
        Label6.Caption = "-- Todos --"
End Select

End Sub

Private Sub Option3_Click(Index As Integer)

OptTres = Index
TodosCLTE = False

Frame5.Visible = True
Frame4.Visible = True
Frame3.Visible = False

Combo1.Visible = False
Combo2.Visible = False
FECHA.Visible = True
Label4.Visible = True
        
        Label4.Caption = "Fecha :"
        FECHA.Value = Date

End Sub
' ---------------------------------------------------
'        Listados Varios
' ---------------------------------------------------

'        Clientes   - AC010L -   1
Private Sub SelClientes()

Label1.Caption = "De Clientes"

Option2(0).Caption = "Por Provincia"
Option2(1).Caption = "Una Zona"
Option2(2).Caption = "Todos"

Option2(2).Value = True
Label6.Visible = True

End Sub

Private Sub LisClientes()
Dim Cliente As String * 5, I1%, I4
  
IMPRE.Caption = " Impresión del Listado de Clientes "

IMPRE.VP1.Font.Name = "Arial"
IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprLegal
IMPRE.VP1.Orientation = orLandscape
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"
Rayitas = String(180, "=")

'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 500

TitCltes

Cont = 0

Select Case OptDos

    Case 0  '   Por Provincia
'        CLTE.RecordSource = "SELECT * FROM Clientes Order by PCIA ASC"
        Tit3 = "AC010L" & Space(26) & "Provincia : " & _
                Val(Mid(Combo6.List(Combo6.ListIndex), 1, 2)) & Space(25) & "Hoja Nro. "
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE pcia = '" & Mid(Combo6.List(Combo6.ListIndex), 1, 2) & "'"

    Case 1 '   Una Zona
        Tit3 = "AC010L" & Space(26) & "Zona : " & _
                Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & Space(25) & "Hoja Nro. "
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE zona = " & Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2))

    Case 2 ' Todos
        CLTE.RecordSource = "SELECT * FROM Clientes Order by NOMB ASC"
End Select
    
    CLTE.Refresh

If CLTE.Recordset.EOF Then Exit Sub

Label8.Caption = " Se listarán " & CLTE.Recordset.RecordCount & " renglones"
CLTE.Recordset.MoveFirst


TotPed = 0: TotItems = 0

Do Until CLTE.Recordset.EOF

RSet Cliente = CLTE.Recordset!CODIGO
I4 = Asc(CLTE.Recordset!PCIA)
If I4 = 32 Then I4 = 91
d$ = Cliente & "|" & CLTE.Recordset!Nomb & "|" & CLTE.Recordset!tel1 & "|" & CLTE.Recordset!Dir & "|" & _
           CLTE.Recordset!Loc & "|" & Mid(PCIA(I4 - 65), 2, 10)
        
        TotItems = TotItems + 1

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True

Cont = IMPRE.VP1.CurrentY

If Cont > 11500 Then

    IMPRE.VP1.NewPage
    TitCltes
    
End If
       
        CLTE.Recordset.MoveNext

Loop

'    TitTot = " Items : " & TotItems
'    ImprTot
   
    TotItems = 0

' total final

 IMPRE.VP1.EndDoc

'VSPDF81.Title = "Subdiario de Ventas - " & Format(FECHA.Value, "mmmm") & " de " & FECHA.Year
'VSPDF81.ConvertDocument IMPRE.VP1, "c:\Subdiario-" & Mid(FECHA.Value, 4, 2) & FECHA.Year & ".pdf"
 
IMPRE.Show 1


Option2(0).Caption = "Uno Sólo"
Option2(1).Caption = "Por Zona"
Option2(2).Caption = "Todos"

End Sub

Private Sub TitCltes()

Escribe "Arial", 9, 1, 0, 5, 10, NomEmpr
Escribe "Arial", 9, 1, 0, 5, 310, "Fecha " & Format(Date, "dd/mm/yyyy")
Escribe "Arial", 14, 1, 0, 10, 165, "C l i e n t e s"

Escribe "Arial", 8.7, 1, 0, 15, 5, Rayitas
Escribe "", 8.7, 1, 0, 20, 5, Space(21) & "Código" & Space(30) & "Razón Social" & _
                Space(75) & "Teléfono" & Space(15) & "Dirección" & Space(75) & "Localidad" & _
                Space(40) & "Provincia"
Escribe "", 8.7, 1, 0, 25, 5, Rayitas
    
 TotItems = 0

    IMPRE.VP1.TableBorder = tbNone

s$ = ">+30mm|<+100mm|<+35mm|<+75mm|<+60mm|<+40mm"
h$ = "Codigo|RSocial|Tel|Direccion|Localidad|Provincia"

With IMPRE.VP1
    .FontSize = 8.7
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Lucida Console"
    .CurrentY = "30mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "8mm"
    .MarginBottom = "3mm"
End With

End Sub

'        Listado de Facturas       - FC020L -      3
Private Sub SelLisFact()

Label1.Caption = "Listado de Facturas"

Frame4.Visible = True
Option3(0).Value = True

Label6.Visible = False

End Sub

Private Sub LisLisFact()
Dim CLTEDsd As Integer, CLTEHst As Integer, Clientes As Integer, I1%
Dim TotItems As Integer, TotPed As Integer, TipoCpbt As String * 4
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2
Dim PClte As String * 5, PNomb As String * 25, PCpbte As String * 6, PCUIT As String * 11
Dim PGrav As String * 12, PIVAIns As String * 11, PMonot As String * 10, PExento As String * 10
Dim PTotal As String * 13, TTotal As String * 13
Dim TPGrav As String * 12, TPIVAIns As String * 11, TPMonot As String * 11, TPExento As String * 12
Dim TPTotal As String * 14
Dim Tot1 As Currency, Tot2 As Currency, Tot3 As Currency

Asteriscos = String(132, "*")
Rayitas = String(132, "-")

'   Si es por Mes
If OptTres = 1 Then
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
    FECHADsd = "01/" & FECHA.Month & "/" & FECHA.Year
    FECHAHst = ElDia & "/" & FECHA.Month & "/" & FECHA.Year
End If

Tit1 = Chr(14) & NomEmpr & Chr(20) & Space(89) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(70) & Chr(14) & "Subdiario de Ventas" & Chr(20)
Tit3 = "FC020L" & Space(115) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-1234"
Tit4 = "*Cpbte.*Tipo*   Fecha  *               Clien" & _
       "te         *  C.U.I.T. * Vta.Grav.I.*  IVA I" & _
       "nsc. *  Exento  * Monotrib.*     Total     *"

laLOGIK1 = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year

Select Case OptTres
    
    Case 0
        FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE FECHA = " & Separa & laLOGIK1 & _
        Separa
        
    Case 1
        FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE FECHA >= " & Separa & FECHADsd & _
        Separa & " and FECHA <= " & Separa & FECHAHst & Separa

End Select

    FCIVA.Refresh

If FCIVA.Recordset.EOF Then Exit Sub

Label8.Caption = " Se listarán " & FCIVA.Recordset.RecordCount & " renglones"
FCIVA.Recordset.MoveFirst

'Open "\\Server\lx-300" For Output As #1 Len = 226
Open SaleImpr For Output As #1 Len = 226
    ImprMatri Chr(27) & "@" & Chr(27) & "CH" & Chr(18) & Chr(20), 0
    
    ImprTitulo

TotPed = 0: TotItems = 0

Do Until FCIVA.Recordset.EOF
   
Select Case FCIVA.Recordset!TIPO
    
    Case 0
        TipoCpbt = "Sdo.A."
    Case 1
        TipoCpbt = "Fact"
    Case 2
        TipoCpbt = "N/Cr"
    Case 3
        TipoCpbt = "N/Db"
    Case 4
        TipoCpbt = "Rec."
    Case 5
        TipoCpbt = "P/C."
    Case 6
        TipoCpbt = "Dcto"
    Case 7
        TipoCpbt = "RAnt"
    Case 8
        TipoCpbt = " -- "
    Case 9
        TipoCpbt = "Anul"

End Select

        RSet PCpbte = FCIVA.Recordset!Cpbte
        RSet PClte = FCIVA.Recordset!CLTE
        RSet PNomb = FCIVA.Recordset!Nomb
        RSet PCUIT = FCIVA.Recordset!Cuit
        
        RSet PGrav = Format(FCIVA.Recordset!GRINS, "###,##0.00")
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!GRINS + FCIVA.Recordset!IVAIns), "###,##0.00")
        
        TitDet = " " & PCpbte & " " & TipoCpbt & _
                " " & Format(FCIVA.Recordset!FECHA, "dd-mm-yyyy") & _
                " " & PClte & " " & PNomb & " " & PCUIT & _
                " " & PGrav & " " & PIVAIns & " " & PExento & _
                " " & PMonot & " " & PTotal
        ImprDet
        TotItems = TotItems + 1
        
        FCIVA.Recordset.MoveNext
        If FCIVA.Recordset.EOF Then Exit Do
        
        Tot1 = Tot1 + CCur(PGrav)
        Tot2 = Tot2 + CCur(PIVAIns)
        Tot3 = Tot3 + CCur(PTotal)
        
    Loop
    
        TitDet = ""
        ImprDet
        
        RSet PGrav = Format(Tot1, "###,##0.00")
        RSet PIVAIns = Format(Tot2, "###,##0.00")
        RSet PTotal = Format(Tot3, "###,##0.00")
  
    TitTot = " Comprobantes : " & TotItems & Space(51) & PGrav & " " & PIVAIns & Space(26) & PTotal
    ImprTot
   
    TotPed = 0: TotItems = 0: Tot1 = 0: Tot2 = 0: Tot3 = 0

' total final

Close #1

End Sub

'        Subdiario de Ventas - Para Comisiones  - xxxxxl -     4
Private Sub SelSubVtasCom()

Label1.Caption = "Subdiario Vtas. (Comisiones)"
Option3(1).Value = True
        
Label2.Visible = True
Label2.Caption = "Vend.:"
Combo2.Visible = False
Combo5.Visible = True

Label6.Visible = True
Check1.Visible = True

End Sub

Private Sub LisSubVtasCom()
Dim CLTEDsd As Integer, CLTEHst As Integer, Clientes As Integer, I1%
Dim TotItems As Integer, TotPed As Integer, VENDED As Integer
Dim PClte As String * 5, PNomb As String * 25, PCpbte As String * 8, PCUIT As String * 11
Dim PGrav As String * 12, PIVAIns As String * 12, PMonot As String * 10, PExento As String * 10
Dim TGrav As String * 12, TIVAIns As String * 11, TMonot As String * 10, TExento As String * 10
Dim PTotal As String * 13, PPtoVta As String * 4, TTotal As String * 13
Dim TPGrav As String * 12, TPIVAIns As String * 11, TPMonot As String * 11, TPExento As String * 12
Dim TPTotal As String * 14, SeniaEOF As Boolean
Dim PORCIB As String * 5, TotIB As String * 10, TTOTIB As String * 10
Dim Tot1 As Currency, Tot2 As Currency, Tot3 As Currency, Tot4 As Currency, Tot5 As Currency
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2
Dim CodIVA As String, TotIBsum As Currency, Qtipo As String * 1

FG2.Clear flexClearScrollable
FG2.Rows = 1

'   Si es por Mes
If OptTres = 1 Then
    FECHADsd = FECHA.Value
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
FECHADsd = FECHA.Month & "/" & "01/" & FECHA.Year
FECHAHst = FECHA.Month & "/" & ElDia & "/" & FECHA.Year
    
        If FECHA.Month = 12 Then
            FECHAHst = "01/01/" & FECHA.Year + 1
        Else
            FECHAHst = FECHA.Month + 1 & "/01/" & FECHA.Year
        End If

End If
  
IMPRE.Caption = " Impresión del Subdiario de Vtas. para Comisiones "

IMPRE.VP1.Font.Name = "Arial"
IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprLegal
IMPRE.VP1.Orientation = orLandscape
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"
Rayitas = String(180, "=")

'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 500

TituSV


' IMPRE.VP1.EndDoc
 
'IMPRE.Show 1

laLOGIK1 = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year
        
Select Case Combo5.ListIndex

    Case 0       '   Si es "TODOS" los vendedores
        Select Case OptTres
            
            Case 0    '   de una fecha (ver que compara con fecha mayor/igual y menor/igual
                FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE FECHA >= " & Separa & FECHA & _
                Separa & " and FECHA < " & Separa & FECHA + 1 & Separa & " order by zona, vend, fecha, tipo, ptovta, cpbte"
                
                CCTE.RecordSource = "SELECT * FROM CTASCTES WHERE FECHA >= " & Separa & FECHA & _
                Separa & " and FECHA < " & Separa & FECHA + 1 & Separa & " AND TIPO = 6 order by fecha, cpbte"
                
            Case 1     '   fechas desde hasta (ver que compara con fecdsd mayor/igual y fechst menor/igual
                FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE FECHA >= " & Separa & FECHADsd & _
                Separa & " and FECHA < " & Separa & FECHAHst & Separa & " order by zona, vend, fecha, tipo, ptovta, cpbte"
                
                CCTE.RecordSource = "SELECT * FROM CTASCTES WHERE FECHA >= " & Separa & FECHADsd & _
                Separa & " and FECHA < " & Separa & FECHAHst & Separa & " AND TIPO = 6  order by fecha, cpbte"
        
        End Select
    Case Else    '   Emite de un Vendedor
        Select Case OptTres
            
            Case 0    '   de una fecha (ver que compara con fecha mayor/igual y menor/igual
                FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE FECHA >= " & Separa & FECHA & _
                Separa & " and FECHA < " & Separa & FECHA + 1 & Separa & " and VEND = " & Val(Mid(Combo5.Text, 1, 2)) & _
                " order by zona, vend, fecha, tipo, ptovta, cpbte"
                
                CCTE.RecordSource = "SELECT * FROM CTASCTES WHERE FECHA >= " & Separa & FECHA & _
                Separa & " and FECHA < " & Separa & FECHA + 1 & Separa & " AND TIPO = 6 order by fecha, cpbte"
                
            Case 1     '   fechas desde hasta (ver que compara con fecdsd mayor/igual y fechst menor/igual
                FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE FECHA >= " & Separa & FECHADsd & _
                Separa & " and FECHA < " & Separa & FECHAHst & Separa & " and VEND = " & Val(Mid(Combo5.Text, 1, 2)) & _
                " order by zona, vend, fecha, tipo, ptovta, cpbte"
                
                CCTE.RecordSource = "SELECT * FROM CTASCTES WHERE FECHA >= " & Separa & FECHADsd & _
                Separa & " and FECHA < " & Separa & FECHAHst & Separa & " AND TIPO = 6  order by fecha, cpbte"
        
        End Select
    
End Select

    FCIVA.Refresh
    CCTE.Refresh

If FCIVA.Recordset.EOF Then Exit Sub

If CCTE.Recordset.EOF Then SeniaEOF = True

' la linea siguiente es para que no tome los recibos
SeniaEOF = True

Label8.Caption = " Se listarán " & FCIVA.Recordset.RecordCount & " renglones"

FCIVA.Recordset.MoveFirst
If Not SeniaEOF Then CCTE.Recordset.MoveFirst

Do Until FCIVA.Recordset.EOF

        RSet PPtoVta = Format(Val(FCIVA.Recordset!PtoVta), "0000")
        RSet PCpbte = Format(Val(FCIVA.Recordset!Cpbte), "00000000")
        RSet PClte = FCIVA.Recordset!CLTE
        RSet PNomb = FCIVA.Recordset!Nomb
        RSet PCUIT = FCIVA.Recordset!Cuit
        If Not FCIVA.Recordset!PORCIB Then
            RSet PORCIB = FCIVA.Recordset!PORCIB
        End If
        
Select Case FCIVA.Recordset!CIVA
    Case 1: CodIVA = "R.I."
    Case 2: CodIVA = "RNI."
    Case 3: CodIVA = "C.F."
    Case 4: CodIVA = "Exto"
    Case 5: CodIVA = "Mon."
End Select
        
     Qtipo = "x"
Select Case FCIVA.Recordset!TIPO
    Case 1: Qtipo = "F"
    Case 2: Qtipo = "C"
    Case 3: Qtipo = "D"
'    Case 4: CodIVA = "Exto"
'    Case 5: CodIVA = "Mon."
End Select

TotIB = ""
TotIBsum = 0
    
    '    si es N/C
    If FCIVA.Recordset!TIPO = 2 Then
        If FCIVA.Recordset!CIVA = 4 Then    '    si es exento - Ushuaia
                RSet PExento = Format(FCIVA.Recordset!GRINS * -1, "###,##0.00")
                RSet PGrav = ""
        Else
                RSet PExento = ""
                RSet PGrav = Format(FCIVA.Recordset!GRINS * -1, "###,##0.00")
        End If
        If FCIVA.Recordset!TotIB > 0 Then
                TotIBsum = FCIVA.Recordset!TotIB
                RSet TotIB = Format(FCIVA.Recordset!TotIB * -1, "###,##0.00")
            End If
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns * -1, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!GRINS + FCIVA.Recordset!IVAIns + TotIBsum) * -1, "###,##0.00")
    Else
        If FCIVA.Recordset!CIVA = 4 Then    '    si es exento - Ushuaia
                RSet PExento = Format(FCIVA.Recordset!GRINS, "###,##0.00")
                RSet PGrav = ""
        Else
                If FCIVA.Recordset!TotIB > 0 Then
                        TotIBsum = FCIVA.Recordset!TotIB
                        RSet TotIB = Format(FCIVA.Recordset!TotIB, "###,##0.00")
                End If
                RSet PExento = ""
                RSet PGrav = Format(FCIVA.Recordset!GRINS, "###,##0.00")
        End If
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!GRINS + FCIVA.Recordset!IVAIns + TotIBsum), "###,##0.00")
    End If
        
        
'  N/Debito por Cheque Rechazado
    If FCIVA.Recordset!TIPO = 3 And FCIVA.Recordset!moti = 2 Then
        RSet PGrav = Format(FCIVA.Recordset!IVAIns / IVAIns, "###,##0.00")
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns, "###,##0.00")
        RSet PExento = Format(FCIVA.Recordset!GRINS - PGrav, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!IVAIns + FCIVA.Recordset!GRINS + TotIBsum), "###,##0.00")
        End If
    If FCIVA.Recordset!TIPO = 2 And FCIVA.Recordset!moti = 2 Then
        RSet PGrav = Format(FCIVA.Recordset!IVAIns * -1 / IVAIns, "###,##0.00")
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns * -1, "###,##0.00")
        RSet PExento = Format((FCIVA.Recordset!GRINS + PGrav) * -1, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!IVAIns + FCIVA.Recordset!GRINS + TotIBsum) * -1, "###,##0.00")
        End If
                
        
d$ = Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & "|" & Qtipo & "|" & FCIVA.Recordset!Letra & "|" & _
           PPtoVta & "-" & PCpbte & "|" & PNomb & "|" & _
           Format(Val(PCUIT), "00-00000000-0") & "|" & CodIVA & "|" & _
           PGrav & "|" & PExento & "|" & TotIB & "|" & PIVAIns & "|" & PTotal
                
            If Check1.Value = 1 Then
                FG2.AddItem FCIVA.Recordset!VEND & vbTab & _
                            Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & vbTab & _
                            Qtipo & vbTab & _
                            FCIVA.Recordset!Letra & vbTab & _
                            PPtoVta & "-" & PCpbte & vbTab & _
                            PNomb & vbTab & _
                            Format(Val(PCUIT), "00-00000000-0") & vbTab & _
                            CodIVA & vbTab & _
                            PGrav & vbTab & _
                            PExento & vbTab & _
                            TotIB & vbTab & _
                            PIVAIns & vbTab & _
                            PTotal, FG2.Rows
            End If
        
'  Suma Totales
        If Val(PGrav) <> 0 Then Tot1 = Tot1 + CCur(PGrav)
        Tot2 = Tot2 + CCur(PIVAIns)
        Tot3 = Tot3 + CCur(PTotal)
        If Val(PExento) <> 0 Then Tot4 = Tot4 + CCur(PExento)
        If Val(TotIB) <> 0 Then Tot5 = Tot5 + CCur(TotIB)

        TotItems = TotItems + 1
        
        FCIVA.Recordset.MoveNext

sinfciva:

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 3
IMPRE.VP1.AddTable s$, h$, "|||||||", , , True
        
'   Salto de Hoja
Cont = Cont + 1
If Cont > 35 Then

        RSet TTOTIB = Format(Tot5, "###,##0.00")
        RSet TExento = Format(Tot4, "###,##0.00")
        RSet TGrav = Format(Tot1, "###,##0.00")
        RSet TIVAIns = Format(Tot2, "###,##0.00")
        RSet TTotal = Format(Tot3, "###,##0.00")

    TitTot = "TRANSPORTE" & Space(59) & TGrav & Space(11) & TExento & Space(1) & TTOTIB & Space(2) & TIVAIns & Space(10) & TTotal
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 201, 5, Rayitas

    IMPRE.VP1.NewPage
    TituSV
End If
        
    Loop

        RSet TTOTIB = Format(Tot5, "###,##0.00")
        RSet TGrav = Format(Tot1, "###,##0.00")
        RSet TIVAIns = Format(Tot2, "###,##0.00")
        RSet TTotal = Format(Tot3, "###,##0.00")
        RSet TExento = Format(Tot4, "###,##0.00")

    TitTot = "TOTAL" & Space(64) & TGrav & Space(11) & TExento & Space(1) & TTOTIB & Space(2) & TIVAIns & Space(10) & TTotal
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 201, 5, Rayitas
   
    TotPed = 0: TotItems = 0: Tot1 = 0: Tot2 = 0: Tot3 = 0: Tot4 = 0: Cont = 0
    TitTot = ""
' total final

Close #1

 IMPRE.VP1.EndDoc

'VSPDF81.Title = "Subdiario de Ventas - " & Format(FECHA.Value, "mmmm") & " de " & FECHA.Year
'VSPDF81.ConvertDocument IMPRE.VP1, "c:\Subdiario-" & Mid(FECHA.Value, 4, 2) & FECHA.Year & ".pdf"
 
IMPRE.Show 1

If Check1.Value = 1 Then

    FG2.TextMatrix(0, 0) = "Vend."
    FG2.TextMatrix(0, 1) = "Fecha"
    FG2.TextMatrix(0, 2) = "Tipo Cpbte."
    FG2.TextMatrix(0, 3) = "Letra"
    FG2.TextMatrix(0, 4) = "Nº Cpbte."
    FG2.TextMatrix(0, 5) = "Razón Social"
    FG2.TextMatrix(0, 6) = "CUIT"
    FG2.TextMatrix(0, 7) = "Cond.IVA"
    FG2.TextMatrix(0, 8) = "Gravado"
    FG2.TextMatrix(0, 9) = "Exento"
    FG2.TextMatrix(0, 10) = "Percep.IB BsAs"
    FG2.TextMatrix(0, 11) = "IVA 21%"
    FG2.TextMatrix(0, 12) = "Total"

    AExcel 5, 12
End If

End Sub

'        Subdiario de Ventas  - xxxxxl -     4
Private Sub SelSubVtas()

Label1.Caption = "Subdiario de Ventas"
Option3(1).Value = True

Label6.Visible = True
Check1.Visible = True

End Sub

Private Sub LisSubVtas()
Dim CLTEDsd As Integer, CLTEHst As Integer, Clientes As Integer, I1%
Dim TotItems As Integer, TotPed As Integer
Dim PClte As String * 5, PNomb As String * 25, PCpbte As String * 8, PCUIT As String * 11
Dim PGrav As String * 12, PIVAIns As String * 12, PMonot As String * 10, PExento As String * 10
Dim TGrav As String * 12, TIVAIns As String * 11, TMonot As String * 10, TExento As String * 10
Dim PTotal As String * 13, PPtoVta As String * 4, TTotal As String * 13
Dim TPGrav As String * 12, TPIVAIns As String * 11, TPMonot As String * 11, TPExento As String * 12
Dim TPTotal As String * 14, SeniaEOF As Boolean
Dim PORCIB As String * 5, TotIB As String * 10, TTOTIB As String * 10
Dim Tot1 As Currency, Tot2 As Currency, Tot3 As Currency, Tot4 As Currency, Tot5 As Currency
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2
Dim CodIVA As String, TotIBsum As Currency, Qtipo As String * 1


FG2.Clear flexClearScrollable
FG2.Rows = 1

'   Si es por Mes
If OptTres = 1 Then
    FECHADsd = FECHA.Value
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
FECHADsd = FECHA.Month & "/" & "01/" & FECHA.Year
FECHAHst = FECHA.Month & "/" & ElDia & "/" & FECHA.Year
    
        If FECHA.Month = 12 Then
            FECHAHst = "01/01/" & FECHA.Year + 1
        Else
            FECHAHst = FECHA.Month + 1 & "/01/" & FECHA.Year
        End If

End If
  
IMPRE.Caption = " Impresión del Subdiario de Ventas "

IMPRE.VP1.Font.Name = "Arial"
IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprLegal
IMPRE.VP1.Orientation = orLandscape
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"
Rayitas = String(180, "=")

'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 500

TituSV


' IMPRE.VP1.EndDoc
 
'IMPRE.Show 1

laLOGIK1 = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year

Select Case OptTres
    
    Case 0
'        FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE FECHA >= " & Separa & laLOGIK1 & _
'        Separa & " order by ptovta, cpbte"
        FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE FECHA >= " & Separa & FECHA & _
        Separa & " and FECHA < " & Separa & FECHA + 1 & Separa & " order by fecha, tipo, ptovta, cpbte"
        
        CCTE.RecordSource = "SELECT * FROM CTASCTES WHERE FECHA >= " & Separa & FECHA & _
        Separa & " and FECHA < " & Separa & FECHA + 1 & Separa & " AND TIPO = 6 order by fecha, cpbte"
        
    Case 1
'        laLOGIK1 = Month(FECHADsd) & "/" & Day(FECHADsd) & "/" & Year(FECHADsd)
'        laLOGIK2 = Month(FECHAHst) & "/01/" & Year(FECHAHst)
        FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE FECHA >= " & Separa & FECHADsd & _
        Separa & " and FECHA < " & Separa & FECHAHst & Separa & " order by fecha, tipo, ptovta, cpbte"
        
        CCTE.RecordSource = "SELECT * FROM CTASCTES WHERE FECHA >= " & Separa & FECHADsd & _
        Separa & " and FECHA < " & Separa & FECHAHst & Separa & " AND TIPO = 6  order by fecha, cpbte"

End Select

    FCIVA.Refresh
    CCTE.Refresh

If FCIVA.Recordset.EOF Then Exit Sub

If CCTE.Recordset.EOF Then SeniaEOF = True

' la linea siguiente es para que no tome los recibos
SeniaEOF = True

Label8.Caption = " Se listarán " & FCIVA.Recordset.RecordCount & " renglones"

FCIVA.Recordset.MoveFirst
If Not SeniaEOF Then CCTE.Recordset.MoveFirst

Do Until FCIVA.Recordset.EOF

'If Not SeniaEOF Then
'    If CCTE.Recordset!FECHA < FCIVA.Recordset!FECHA Then
'
' '       RSet PPtoVta = Format(Val(CCTE.Recordset!prefijo), "0000")
'        RSet PCpbte = Format(Val(CCTE.Recordset!Cpbte), "00000000")
'        RSet PClte = CCTE.Recordset!CLTE
'        RSet PNomb = ""
'        RSet PCUIT = ""
'
'        RSet PExento = ""
'        RSet PGrav = Format((CCTE.Recordset!Impte / 1.21) * -1, "###,##0.00")
'        RSet PIVAIns = Format((PGrav * 0.21), "###,##0.00")
'        RSet PTotal = Format(CCTE.Recordset!Impte * -1, "###,##0.00")
'
'
'd$ = Format(CCTE.Recordset!FECHA, "dd/mm/yyyy") & "|" & CCTE.Recordset!Letra & "|" & _
'           PPtoVta & "-" & PCpbte & "|" & PNomb & "|" & _
'           Format(Val(PCUIT), "00-00000000-0") & "|" & CodIVA & "|" & _
'           PGrav & "|" & PExento & "|" & PIVAIns & "|" & PTotal
        
'    CCTE.Recordset.MoveNext
'    If CCTE.Recordset.EOF Then SeniaEOF = True
    
'    GoTo sinfciva
'  End If
'End If

        RSet PPtoVta = Format(Val(FCIVA.Recordset!PtoVta), "0000")
        RSet PCpbte = Format(Val(FCIVA.Recordset!Cpbte), "00000000")
        RSet PClte = FCIVA.Recordset!CLTE
        RSet PNomb = FCIVA.Recordset!Nomb
        RSet PCUIT = FCIVA.Recordset!Cuit
        If Not FCIVA.Recordset!PORCIB Then
            RSet PORCIB = FCIVA.Recordset!PORCIB
        End If
        
Select Case FCIVA.Recordset!CIVA
    Case 1: CodIVA = "R.I."
    Case 2: CodIVA = "RNI."
    Case 3: CodIVA = "C.F."
    Case 4: CodIVA = "Exto"
    Case 5: CodIVA = "Mon."
End Select
        
     Qtipo = "x"
Select Case FCIVA.Recordset!TIPO
    Case 1: Qtipo = "F"
    Case 2: Qtipo = "C"
    Case 3: Qtipo = "D"
'    Case 4: CodIVA = "Exto"
'    Case 5: CodIVA = "Mon."
End Select

TotIB = ""
TotIBsum = 0

                RSet PExento = ""
        If FCIVA.Recordset!CIVA > 2 Then    '    si es CF, exento o MTbuto
        '        RSet PExento = Format(FCIVA.Recordset!GRINS * -1, "###,##0.00")
                RSet PGrav = Format(FCIVA.Recordset!GRINS / 1.21, "###,##0.00")
                RSet PIVAIns = Format(FCIVA.Recordset!GRINS - PGrav, "###,##0.00")
        Else
                RSet PGrav = Format(FCIVA.Recordset!GRINS, "###,##0.00")
                RSet PIVAIns = Format(FCIVA.Recordset!IVAIns, "###,##0.00")
        End If
    
    '    si es N/C
    If FCIVA.Recordset!TIPO = 2 Then
        If FCIVA.Recordset!TotIB > 0 Then
                RSet TotIB = Format(FCIVA.Recordset!TotIB * -1, "###,##0.00")
                TotIBsum = TotIB
            End If
        RSet PGrav = Format(PGrav * -1, "###,##0.00")
        RSet PIVAIns = Format(PIVAIns * -1, "###,##0.00")
 '       RSet PTotal = Format((PGrav + PIVAIns + TotIBsum) * -1, "###,##0.00")
    
    Else
        
                If FCIVA.Recordset!TotIB > 0 Then
                        TotIBsum = FCIVA.Recordset!TotIB
                        RSet TotIB = Format(FCIVA.Recordset!TotIB, "###,##0.00")
                End If
                
    
    End If
'        RSet PTotal = Format((FCIVA.Recordset!GRINS + FCIVA.Recordset!IVAIns + TotIBsum), "###,##0.00")
        RSet PTotal = Format((CCur(PGrav) + CCur(PIVAIns) + CCur(TotIBsum)), "###,##0.00")
        
        
'  N/Debito por Cheque Rechazado
    If FCIVA.Recordset!TIPO = 3 And FCIVA.Recordset!moti = 2 Then
        RSet PGrav = Format(FCIVA.Recordset!IVAIns / IVAIns, "###,##0.00")
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns, "###,##0.00")
        RSet PExento = Format(FCIVA.Recordset!GRINS - PGrav, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!IVAIns + FCIVA.Recordset!GRINS + TotIBsum), "###,##0.00")
        End If
    If FCIVA.Recordset!TIPO = 2 And FCIVA.Recordset!moti = 2 Then
        RSet PGrav = Format(FCIVA.Recordset!IVAIns * -1 / IVAIns, "###,##0.00")
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns * -1, "###,##0.00")
        RSet PExento = Format((FCIVA.Recordset!GRINS + PGrav) * -1, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!IVAIns + FCIVA.Recordset!GRINS + TotIBsum) * -1, "###,##0.00")
        End If
                
        
d$ = Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & "|" & Qtipo & "|" & FCIVA.Recordset!Letra & "|" & _
           PPtoVta & "-" & PCpbte & "|" & PNomb & "|" & _
           Format(Val(PCUIT), "00-00000000-0") & "|" & CodIVA & "|" & _
           PGrav & "|" & PExento & "|" & TotIB & "|" & PIVAIns & "|" & PTotal
                
            If Check1.Value = 1 Then
                FG2.AddItem Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & vbTab & _
                            Qtipo & vbTab & _
                            FCIVA.Recordset!Letra & vbTab & _
                            PPtoVta & "-" & PCpbte & vbTab & _
                            PNomb & vbTab & _
                            Format(Val(PCUIT), "00-00000000-0") & vbTab & _
                            CodIVA & vbTab & _
                            PGrav & vbTab & _
                            PExento & vbTab & _
                            TotIB & vbTab & _
                            PIVAIns & vbTab & _
                            PTotal, FG2.Rows
            End If
        
'  Suma Totales
        If Val(PGrav) <> 0 Then Tot1 = Tot1 + CCur(PGrav)
        Tot2 = Tot2 + CCur(PIVAIns)
        Tot3 = Tot3 + CCur(PTotal)
        If Val(PExento) <> 0 Then Tot4 = Tot4 + CCur(PExento)
        If Val(TotIB) <> 0 Then Tot5 = Tot5 + CCur(TotIB)

        TotItems = TotItems + 1
        
        FCIVA.Recordset.MoveNext

sinfciva:

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 3
IMPRE.VP1.AddTable s$, h$, "|||||||", , , True
        
'   Salto de Hoja
Cont = Cont + 1
If Cont > 35 Then

        RSet TTOTIB = Format(Tot5, "###,##0.00")
        RSet TExento = Format(Tot4, "###,##0.00")
        RSet TGrav = Format(Tot1, "###,##0.00")
        RSet TIVAIns = Format(Tot2, "###,##0.00")
        RSet TTotal = Format(Tot3, "###,##0.00")

    TitTot = "TRANSPORTE" & Space(59) & TGrav & Space(11) & TExento & Space(1) & TTOTIB & Space(2) & TIVAIns & Space(10) & TTotal
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 201, 5, Rayitas

    IMPRE.VP1.NewPage
    TituSV
End If
        
  'If FCIVA.Recordset!CIVA = 4 Then
  '  Tot4 = 0
  '  End If
    
        
    Loop

        RSet TTOTIB = Format(Tot5, "###,##0.00")
        RSet TGrav = Format(Tot1, "###,##0.00")
        RSet TIVAIns = Format(Tot2, "###,##0.00")
        RSet TTotal = Format(Tot3, "###,##0.00")
        RSet TExento = Format(Tot4, "###,##0.00")

    TitTot = "TOTAL" & Space(64) & TGrav & Space(11) & TExento & Space(1) & TTOTIB & Space(2) & TIVAIns & Space(10) & TTotal
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 201, 5, Rayitas
   
    TotPed = 0: TotItems = 0: Tot1 = 0: Tot2 = 0: Tot3 = 0: Tot4 = 0: Cont = 0
    TitTot = ""
' total final

Close #1

 IMPRE.VP1.EndDoc

'VSPDF81.Title = "Subdiario de Ventas - " & Format(FECHA.Value, "mmmm") & " de " & FECHA.Year
'VSPDF81.ConvertDocument IMPRE.VP1, "c:\Subdiario-" & Mid(FECHA.Value, 4, 2) & FECHA.Year & ".pdf"
 
IMPRE.Show 1

If Check1.Value = 1 Then

    FG2.TextMatrix(0, 0) = "Fecha"
    FG2.TextMatrix(0, 1) = "Tipo Cpbte."
    FG2.TextMatrix(0, 2) = "Letra"
    FG2.TextMatrix(0, 3) = "Nº Cpbte."
    FG2.TextMatrix(0, 4) = "Razón Social"
    FG2.TextMatrix(0, 5) = "CUIT"
    FG2.TextMatrix(0, 6) = "Cond.IVA"
    FG2.TextMatrix(0, 7) = "Gravado"
    FG2.TextMatrix(0, 8) = "Exento"
    FG2.TextMatrix(0, 9) = "Percep.IB BsAs"
    FG2.TextMatrix(0, 10) = "IVA 21%"
    FG2.TextMatrix(0, 11) = "Total"

    AExcel 1, 11
End If

End Sub

Private Sub TituSV()

Escribe "Arial", 12, 1, 0, 5, 100, NomEmpr
Escribe "Arial", 8.7, 0, 0, 10, 5, "CUIT Nº 33-70346790-9"
Escribe "Arial", 8.7, 0, 0, 10, 150, "MES : " & Format(FECHA.Value, "mmmm") & " de " & FECHA.Year
Escribe "Arial", 8.7, 1, 0, 10, 275, "IVA VENTAS"

Escribe "Arial", 8.7, 1, 0, 15, 5, Rayitas
Escribe "", 8.7, 0, 0, 20, 5, Space(11) & "FECHA" & Space(5) & "TIP" & Space(4) & "COMPROBANTE Nº" & _
                Space(50) & "C L I E N T E" & Space(60) & "N E T O   G R A V A D O" & _
                Space(7) & "CONC.y O." & Space(35) & "COMPRAS"
Escribe "", 8.7, 0, 0, 25, 90, "NOMBRE" & Space(34) & "CUIT" & Space(14) & "COND.IVA" & Space(20) & _
                "21%" & Space(20) & "27%" & Space(10) & "EXENTAS" & Space(2) & "Perc.IB BsAs" & Space(2) & _
                "IVA GENERAL" & Space(3) & "RESP.NO INS." & Space(3) & "PERC.IVA" & Space(2) & "TOT.FC."

Escribe "", 8.7, 1, 0, 30, 5, Rayitas
    
Escribe "Lucida Console", 8.7, 1, 0, 35, 50, TitTot

TotPed = 0: TotItems = 0

    IMPRE.VP1.TableBorder = tbNone

s$ = ">+22mm|^+5mm|^+5mm|^+36mm|<+52mm|<+28mm|^+14mm|>+28mm|>+42mm|>+24mm|>+26mm|>+52mm"
h$ = "fecha|letra|TCpbt|ptovta-cpbte|clte|cuit|civa|grav|exto|PIB|iva|total"

With IMPRE.VP1
    .FontSize = 8
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Lucida Console"
    .CurrentY = "40mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "8mm"
    .MarginBottom = "3mm"
End With

Cont = 0

End Sub
'        Percepciones ARBA  - xxxxxl -     4
Private Sub SelPerARBA()

Label1.Caption = "Percepciones ARBA"
Option3(1).Value = True

Label6.Visible = True
Check1.Visible = True

End Sub

Private Sub LisPerARBA()
Dim CLTEDsd As Integer, CLTEHst As Integer, Clientes As Integer, I1%
Dim TotItems As Integer, TotPed As Integer
Dim PClte As String * 5, PNomb As String * 25, PCpbte As String * 8, PCUIT As String * 11
Dim PGrav As String * 12, PIVAIns As String * 12, PMonot As String * 10, PExento As String * 10
Dim TGrav As String * 12, TIVAIns As String * 11, TMonot As String * 10, TExento As String * 10
Dim PTotal As String * 13, PPtoVta As String * 4, TTotal As String * 13
Dim TPGrav As String * 12, TPIVAIns As String * 11, TPMonot As String * 11, TPExento As String * 12
Dim TPTotal As String * 14, SeniaEOF As Boolean
Dim PORCIB As String * 5, TotIB As String * 10, TTOTIB As String * 10
Dim Tot1 As Currency, Tot2 As Currency, Tot3 As Currency, Tot4 As Currency, Tot5 As Currency
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2
Dim CodIVA As String, TCpbte As String

FG2.Clear flexClearScrollable
FG2.Rows = 1

'   Si es por Mes
If OptTres = 1 Then
    FECHADsd = FECHA.Value
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
FECHADsd = FECHA.Month & "/" & "01/" & FECHA.Year
FECHAHst = FECHA.Month & "/" & ElDia & "/" & FECHA.Year
    
        If FECHA.Month = 12 Then
            FECHAHst = "01/01/" & FECHA.Year + 1
        Else
            FECHAHst = FECHA.Month + 1 & "/01/" & FECHA.Year
        End If

End If
  
IMPRE.Caption = " Impresión del Subdiario de Ventas "

IMPRE.VP1.Font.Name = "Arial"
IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprLegal
IMPRE.VP1.Orientation = orLandscape
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"
Rayitas = String(180, "=")

'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 500

TituPer

laLOGIK1 = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year

Select Case OptTres
    
    Case 0
        FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE TOTIB > 0 AND FECHA >= " & Separa & FECHA & _
        Separa & " and FECHA < " & Separa & FECHA + 1 & Separa & " order by ptovta, cpbte"
        
    Case 1
        FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE TOTIB > 0 AND FECHA >= " & Separa & FECHADsd & _
        Separa & " and FECHA < " & Separa & FECHAHst & Separa & " order by ptovta, cpbte"

End Select

    FCIVA.Refresh

If FCIVA.Recordset.EOF Then Exit Sub

Dim NomARBA As String
NomARBA = "c:\AR-33703467909-"
NomARBA = NomARBA & Mid(FECHA, 7, 4)
NomARBA = NomARBA & Mid(FECHA, 4, 2)
NomARBA = NomARBA & "0-7-LOTE1.txt"

'Open ServidorSQL & "PercARBA.txt" For Output As #1 Len = 226
Open NomARBA For Output As #1 Len = 226

Label8.Caption = " Se listarán " & FCIVA.Recordset.RecordCount & " renglones"

FCIVA.Recordset.MoveFirst

Do Until FCIVA.Recordset.EOF

        RSet PPtoVta = Format(Val(FCIVA.Recordset!PtoVta), "0000")
        RSet PCpbte = Format(Val(FCIVA.Recordset!Cpbte), "00000000")
        RSet PClte = FCIVA.Recordset!CLTE
        RSet PNomb = FCIVA.Recordset!Nomb
        RSet PCUIT = FCIVA.Recordset!Cuit
        If Not FCIVA.Recordset!PORCIB Then
            RSet PORCIB = FCIVA.Recordset!PORCIB
        End If
        
Select Case FCIVA.Recordset!CIVA
    Case 1: CodIVA = "R.I."
    Case 2: CodIVA = "RNI."
    Case 3: CodIVA = "C.F."
    Case 4: CodIVA = "Exto"
    Case 5: CodIVA = "Mon."
End Select
        
Select Case FCIVA.Recordset!TIPO
    Case 1: TCpbte = "F"
    Case 2: TCpbte = "C"
    Case 3: TCpbte = "D"
    Case 4: TCpbte = "R"
End Select

TotIB = ""
    '    si es N/C
    If FCIVA.Recordset!TIPO = 2 Then
        If FCIVA.Recordset!CIVA = 4 Then    '    si es exento - Ushuaia
                RSet PExento = Format(FCIVA.Recordset!GRINS * -1, "###,##0.00")
                RSet PGrav = ""
        Else
                RSet PExento = ""
                RSet PGrav = Format(FCIVA.Recordset!GRINS * -1, "###,##0.00")
        End If
                If FCIVA.Recordset!TotIB > 0 Then
                        RSet TotIB = Format(FCIVA.Recordset!TotIB * -1, "###,##0.00")
                End If
                RSet PIVAIns = Format(FCIVA.Recordset!IVAIns * -1, "###,##0.00")
                RSet PTotal = Format((FCIVA.Recordset!GRINS + FCIVA.Recordset!IVAIns) * -1, "###,##0.00")
    Else
        If FCIVA.Recordset!CIVA = 4 Then    '    si es exento - Ushuaia
                RSet PExento = Format(FCIVA.Recordset!GRINS, "###,##0.00")
                RSet PGrav = ""
        Else
                If FCIVA.Recordset!TotIB > 0 Then
                        RSet TotIB = Format(FCIVA.Recordset!TotIB, "###,##0.00")
                End If
                RSet PExento = ""
                RSet PGrav = Format(FCIVA.Recordset!GRINS, "###,##0.00")
        End If
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!GRINS + FCIVA.Recordset!IVAIns), "###,##0.00")
    End If
        
    If FCIVA.Recordset!TIPO = 3 And FCIVA.Recordset!moti = 2 Then
        RSet PGrav = Format(FCIVA.Recordset!IVAIns / IVAIns, "###,##0.00")
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns, "###,##0.00")
        RSet PExento = Format(FCIVA.Recordset!GRINS - PGrav, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!IVAIns + FCIVA.Recordset!GRINS), "###,##0.00")
        End If
        
If TCpbte = "C" Then
    RSet PIVAIns = Format(CCur(PGrav), "00000000.00")
    RSet TIVAIns = Format(CCur(TotIB), "0000000.00")
Else
    RSet PIVAIns = Format(CCur(PGrav), "000000000.00")
    RSet TIVAIns = Format(CCur(TotIB), "00000000.00")
End If

d$ = Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & "|" & FCIVA.Recordset!Letra & "|" & _
           PPtoVta & "-" & PCpbte & "|" & PNomb & "|" & _
           Format(Val(PCUIT), "00-00000000-0") & "|" & CodIVA & "|" & _
           PGrav & "|" & TotIB
                
            If Check1.Value = 1 Then
                FG2.AddItem Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & vbTab & _
                            FCIVA.Recordset!Letra & vbTab & _
                            PPtoVta & "-" & PCpbte & vbTab & _
                            PNomb & vbTab & _
                            Format(Val(PCUIT), "00-00000000-0") & vbTab & _
                            CodIVA & vbTab & _
                            PGrav & vbTab & _
                            TotIB, FG2.Rows
            End If

'Print #1, Format(Val(PCUIT), "00-00000000-0"); Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & TCpbte & _
'           FCIVA.Recordset!Letra & PPtoVta & PCpbte & PIVAIns & TIVAIns
Print #1, Format(Val(PCUIT), "00-00000000-0"); Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & TCpbte & _
           FCIVA.Recordset!Letra & PPtoVta & PCpbte & PIVAIns & TIVAIns & "A"
        
'  Suma Totales
        If Val(PGrav) <> 0 Then Tot1 = Tot1 + CCur(PGrav)
        Tot2 = Tot2 + CCur(PIVAIns)
        Tot3 = Tot3 + CCur(PTotal)
        If Val(PExento) <> 0 Then Tot4 = Tot4 + CCur(PExento)
        If Val(TotIB) <> 0 Then Tot5 = Tot5 + CCur(TotIB)

        TotItems = TotItems + 1
        
        FCIVA.Recordset.MoveNext

sinfciva:

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 3
IMPRE.VP1.AddTable s$, h$, "|||||||", , , True
        
'   Salto de Hoja
Cont = Cont + 1
If Cont > 30 Then

        RSet TTOTIB = Format(Tot5, "###,##0.00")
        RSet TGrav = Format(Tot1, "###,##0.00")

    TitTot = "TRANSPORTE" & Space(53) & TGrav & Space(3) & TTOTIB
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 203, 5, Rayitas

    IMPRE.VP1.NewPage
    TituPer
End If

    Loop

        RSet TTOTIB = Format(Tot5, "###,##0.00")
        RSet TGrav = Format(Tot1, "###,##0.00")

    TitTot = "TOTAL" & Space(58) & TGrav & Space(3) & TTOTIB
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 201, 5, Rayitas
   
    TotPed = 0: TotItems = 0: Tot1 = 0: Tot2 = 0: Tot3 = 0: Tot4 = 0: Cont = 0
    TitTot = ""
' total final

Close #1

 IMPRE.VP1.EndDoc
 
IMPRE.Show 1

If Check1.Value = 1 Then

    FG2.TextMatrix(0, 0) = "Fecha"
    FG2.TextMatrix(0, 1) = "Letra"
    FG2.TextMatrix(0, 2) = "Nº Cpbte."
    FG2.TextMatrix(0, 3) = "Razón Social"
    FG2.TextMatrix(0, 4) = "CUIT"
    FG2.TextMatrix(0, 5) = "Cond.IVA"
    FG2.TextMatrix(0, 6) = "Imponible"
    FG2.TextMatrix(0, 7) = "Percepción"
    
    AExcel 3, 7
End If

End Sub

Private Sub LisPerARBAant()
Dim CLTEDsd As Integer, CLTEHst As Integer, Clientes As Integer, I1%
Dim TotItems As Integer, TotPed As Integer
Dim PClte As String * 5, PNomb As String * 25, PCpbte As String * 8, PCUIT As String * 11
Dim PGrav As String * 12, PIVAIns As String * 12, PMonot As String * 10, PExento As String * 10
Dim TGrav As String * 12, TIVAIns As String * 11, TMonot As String * 10, TExento As String * 10
Dim PTotal As String * 13, PPtoVta As String * 4, TTotal As String * 13
Dim TPGrav As String * 12, TPIVAIns As String * 11, TPMonot As String * 11, TPExento As String * 12
Dim TPTotal As String * 14, SeniaEOF As Boolean
Dim PORCIB As String * 5, TotIB As String * 10, TTOTIB As String * 10
Dim Tot1 As Currency, Tot2 As Currency, Tot3 As Currency, Tot4 As Currency, Tot5 As Currency
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2
Dim CodIVA As String, TCpbte As String

'   Si es por Mes
If OptTres = 1 Then
    FECHADsd = FECHA.Value
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
FECHADsd = FECHA.Month & "/" & "01/" & FECHA.Year
FECHAHst = FECHA.Month & "/" & ElDia & "/" & FECHA.Year
    
        If FECHA.Month = 12 Then
            FECHAHst = "01/01/" & FECHA.Year + 1
        Else
            FECHAHst = FECHA.Month + 1 & "/01/" & FECHA.Year
        End If

End If
  
IMPRE.Caption = " Impresión del Subdiario de Ventas "

IMPRE.VP1.Font.Name = "Arial"
IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprLegal
IMPRE.VP1.Orientation = orLandscape
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"
Rayitas = String(180, "=")

'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 500

TituPer

laLOGIK1 = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year

Select Case OptTres
    
    Case 0
        FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE TOTIB > 0 AND FECHA >= " & Separa & FECHA & _
        Separa & " and FECHA < " & Separa & FECHA + 1 & Separa & " order by ptovta, cpbte"
        
    Case 1
        FCIVA.RecordSource = "SELECT * FROM FCIVAVTA WHERE TOTIB > 0 AND FECHA >= " & Separa & FECHADsd & _
        Separa & " and FECHA < " & Separa & FECHAHst & Separa & " order by ptovta, cpbte"

End Select

    FCIVA.Refresh

If FCIVA.Recordset.EOF Then Exit Sub

Open ServidorSQL & "PercARBA.txt" For Output As #1 Len = 226

Label8.Caption = " Se listarán " & FCIVA.Recordset.RecordCount & " renglones"

FCIVA.Recordset.MoveFirst

Do Until FCIVA.Recordset.EOF

        RSet PPtoVta = Format(Val(FCIVA.Recordset!PtoVta), "0000")
        RSet PCpbte = Format(Val(FCIVA.Recordset!Cpbte), "00000000")
        RSet PClte = FCIVA.Recordset!CLTE
        RSet PNomb = FCIVA.Recordset!Nomb
        RSet PCUIT = FCIVA.Recordset!Cuit
        If Not FCIVA.Recordset!PORCIB Then
            RSet PORCIB = FCIVA.Recordset!PORCIB
        End If
        
Select Case FCIVA.Recordset!CIVA
    Case 1: CodIVA = "R.I."
    Case 2: CodIVA = "RNI."
    Case 3: CodIVA = "C.F."
    Case 4: CodIVA = "Exto"
    Case 5: CodIVA = "Mon."
End Select
        
Select Case FCIVA.Recordset!TIPO
    Case 1: TCpbte = "F"
    Case 2: TCpbte = "C"
    Case 3: TCpbte = "D"
    Case 4: TCpbte = "R"
End Select

TotIB = ""
    '    si es N/C
    If FCIVA.Recordset!TIPO = 2 Then
        If FCIVA.Recordset!CIVA = 4 Then    '    si es exento - Ushuaia
                RSet PExento = Format(FCIVA.Recordset!GRINS * -1, "###,##0.00")
                RSet PGrav = ""
        Else
                RSet PExento = ""
                RSet PGrav = Format(FCIVA.Recordset!GRINS * -1, "###,##0.00")
        End If
                If FCIVA.Recordset!TotIB > 0 Then
                        RSet TotIB = Format(FCIVA.Recordset!TotIB * -1, "###,##0.00")
                End If
                RSet PIVAIns = Format(FCIVA.Recordset!IVAIns * -1, "###,##0.00")
                RSet PTotal = Format((FCIVA.Recordset!GRINS + FCIVA.Recordset!IVAIns) * -1, "###,##0.00")
    Else
        If FCIVA.Recordset!CIVA = 4 Then    '    si es exento - Ushuaia
                RSet PExento = Format(FCIVA.Recordset!GRINS, "###,##0.00")
                RSet PGrav = ""
        Else
                If FCIVA.Recordset!TotIB > 0 Then
                        RSet TotIB = Format(FCIVA.Recordset!TotIB, "###,##0.00")
                End If
                RSet PExento = ""
                RSet PGrav = Format(FCIVA.Recordset!GRINS, "###,##0.00")
        End If
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!GRINS + FCIVA.Recordset!IVAIns), "###,##0.00")
    End If
        
    If FCIVA.Recordset!TIPO = 3 And FCIVA.Recordset!moti = 2 Then
        RSet PGrav = Format(FCIVA.Recordset!IVAIns / IVAIns, "###,##0.00")
        RSet PIVAIns = Format(FCIVA.Recordset!IVAIns, "###,##0.00")
        RSet PExento = Format(FCIVA.Recordset!GRINS - PGrav, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!IVAIns + FCIVA.Recordset!GRINS), "###,##0.00")
        End If
        
If TCpbte = "C" Then
    RSet PIVAIns = Format(CCur(PGrav), "00000000.00")
    RSet TIVAIns = Format(CCur(TotIB), "0000000.00")
Else
    RSet PIVAIns = Format(CCur(PGrav), "000000000.00")
    RSet TIVAIns = Format(CCur(TotIB), "00000000.00")
End If

d$ = Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & "|" & FCIVA.Recordset!Letra & "|" & _
           PPtoVta & "-" & PCpbte & "|" & PNomb & "|" & _
           Format(Val(PCUIT), "00-00000000-0") & "|" & CodIVA & "|" & _
           PGrav & "|" & TotIB

'Print #1, Format(Val(PCUIT), "00-00000000-0"); Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & TCpbte & _
'           FCIVA.Recordset!Letra & PPtoVta & PCpbte & PIVAIns & TIVAIns
Print #1, Format(Val(PCUIT), "00-00000000-0"); Format(FCIVA.Recordset!FECHA, "dd/mm/yyyy") & TCpbte & _
           FCIVA.Recordset!Letra & PPtoVta & PCpbte & PIVAIns & TIVAIns & "A"
        
'  Suma Totales
        If Val(PGrav) <> 0 Then Tot1 = Tot1 + CCur(PGrav)
        Tot2 = Tot2 + CCur(PIVAIns)
        Tot3 = Tot3 + CCur(PTotal)
        If Val(PExento) <> 0 Then Tot4 = Tot4 + CCur(PExento)
        If Val(TotIB) <> 0 Then Tot5 = Tot5 + CCur(TotIB)

        TotItems = TotItems + 1
        
        FCIVA.Recordset.MoveNext

sinfciva:

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 3
IMPRE.VP1.AddTable s$, h$, "|||||||", , , True
        
'   Salto de Hoja
Cont = Cont + 1
If Cont > 30 Then

        RSet TTOTIB = Format(Tot5, "###,##0.00")
        RSet TGrav = Format(Tot1, "###,##0.00")

    TitTot = "TRANSPORTE" & Space(53) & TGrav & Space(3) & TTOTIB
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 203, 5, Rayitas

    IMPRE.VP1.NewPage
    TituPer
End If

    Loop

        RSet TTOTIB = Format(Tot5, "###,##0.00")
        RSet TGrav = Format(Tot1, "###,##0.00")

    TitTot = "TOTAL" & Space(58) & TGrav & Space(3) & TTOTIB
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 201, 5, Rayitas
   
    TotPed = 0: TotItems = 0: Tot1 = 0: Tot2 = 0: Tot3 = 0: Tot4 = 0: Cont = 0
    TitTot = ""
' total final

Close #1

 IMPRE.VP1.EndDoc
 
IMPRE.Show 1

End Sub

Private Sub TituPer()

Escribe "Arial", 12, 1, 0, 5, 100, NomEmpr
Escribe "Arial", 8.7, 0, 0, 10, 5, "CUIT Nº 33-70346790-9"
Escribe "Arial", 8.7, 0, 0, 10, 150, "MES : " & Format(FECHA.Value, "mmmm") & " de " & FECHA.Year
Escribe "Arial", 8.7, 1, 0, 10, 275, "Percepciones ARBA"

Escribe "Arial", 8.7, 1, 0, 15, 5, Rayitas
Escribe "", 8.7, 0, 0, 20, 5, Space(11) & "FECHA" & Space(5) & "TIP" & Space(4) & "COMPROBANTE Nº" & _
                Space(50) & "C L I E N T E" & Space(48) & "MONTO IMPONIBLE" & _
                Space(4) & "PERCEPCION"
Escribe "", 8.7, 0, 0, 25, 90, "NOMBRE" & Space(34) & "CUIT" & Space(14) & "COND.IVA"

Escribe "", 8.7, 1, 0, 30, 5, Rayitas
    
Escribe "Lucida Console", 8.7, 1, 0, 35, 50, TitTot

TotPed = 0: TotItems = 0

    IMPRE.VP1.TableBorder = tbNone

s$ = ">+22mm|^+5mm|^+37mm|<+55mm|<+28mm|^+15mm|>+28mm|>+24mm"
h$ = "fecha|letra|ptovta-cpbte|clte|cuit|civa|grav|PIB"

With IMPRE.VP1
    .FontSize = 8
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Lucida Console"
    .CurrentY = "40mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "8mm"
    .MarginBottom = "3mm"
End With

Cont = 0

End Sub

'        Lista de Precios     - LP010L -    2
Private Sub SelPrecios()

Label1.Caption = "Lista de Precios"
Option2(0).Value = True

Label6.Visible = False

End Sub

Private Sub LisPrecios()
Dim CLTEDsd As Integer, CLTEHst As Integer, Clientes As Integer, I1%
Dim TotItems As Integer, TotPed As Integer

Tit1 = Chr(14) & NomEmpr & Chr(20) & Space(34) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(10) & Chr(14) & "Faltas de Alta Rotación" & Chr(20)
Tit3 = "ST090L" & Space(63) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-"
Tit4 = "*             Número FORD          * Cód" & _
       ".Int. *CCTE *C.Ped*  Pedir *Promesa*CA*"

If TodosCLTE Then
    CLTEDsd = 0: CLTEHst = 9999
Else
    CLTEDsd = CodCLTE: CLTEHst = CodCLTE
End If

    CCTE.RecordSource = "SELECT * FROM ArtiFalt WHERE CODIGO >= " & CLTEDsd & _
    " AND CODIGO <= " & CLTEHst & " ORDER BY CODIGO" & _
    " WHERE (ART.CALIF <> 1) AND (ART.CALIF < 10 OR" & _
    " ART.CALIF > 14) AND STREAL < STOK.STREP"
    
    CCTE.Refresh
Debug.Print CCTE.RecordSource

If CCTE.Recordset.EOF Then Exit Sub

Label8.Caption = " Se listarán " & CCTE.Recordset.RecordCount & " renglones"
CCTE.Recordset.MoveFirst

'Open "LPT1:" For Output As #1 Len = 226
Open SaleImpr For Output As #1 Len = 226
    ImprMatri Chr(27) & "@" & Chr(27) & "CH" & Chr(18) & Chr(20), 0

    ImprTitulo

TotPed = 0: TotItems = 0

Do Until CCTE.Recordset.EOF
        
        TitDet = " " & CCTE.Recordset!NROFORD & "   " & CCTE.Recordset!pNroInt & _
                "   " & CCTE.Recordset!CABEC & _
                "   " & Format(CCTE.Recordset!CoefV, "##,##0") & _
                "   " & Format(CCTE.Recordset!FUCPRA, "dd-mm-yyyy")
                
        ImprDet
        TotItems = TotItems + 1
        
        CCTE.Recordset.MoveNext
        If CCTE.Recordset.EOF Then Exit Do
        
    Loop
    
    TitTot = " Items : " & TotItems & "      Cantidad : " & TotPed
    ImprTot
   
    TotPed = 0: TotItems = 0

' total final

Close #1

End Sub
'        Subdiario de Ventas  - xxxxxl -     4
Private Sub SelIngBrutos()

Label1.Caption = "Ingresos Brutos"
Option3(1).Value = True

Label6.Visible = True

End Sub

Private Sub LisIngBrutos()

Dim TotItems As Integer, TotPed As Integer
Dim PPcia As String
Dim PGrav As String * 12, PIVAIns As String * 12, PMonot As String * 10, PExento As String * 10
Dim TGrav As String * 12, TIVAIns As String * 12, TMonot As String * 10, TExento As String * 10
Dim PTotal As String * 13, PPtoVta As String * 4, TTotal As String * 13
Dim TPGrav As String * 12, TPIVAIns As String * 12, TPMonot As String * 11, TPExento As String * 12
Dim TPTotal As String * 14, SeniaEOF As Boolean, DbCr As String
Dim Tot1 As Currency, Tot2 As Currency, Tot3 As Currency
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2
Dim I1%, I4%

'   Calcula Fecha desde-hasta (mes completo)
If OptTres = 1 Then
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
    FECHADsd = "01/" & FECHA.Month & "/" & FECHA.Year
    FECHAHst = ElDia & "/" & FECHA.Month & "/" & FECHA.Year
End If
  
IMPRE.Caption = " Impresión de Ingresos Brutos"

IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprLegal
IMPRE.VP1.Orientation = orPortrait
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"
Rayitas = String(110, "=")

'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 200

TituIB


' IMPRE.VP1.EndDoc
 
'IMPRE.Show 1

laLOGIK1 = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year
        
        laLOGIK1 = Month(FECHADsd) & "/" & Day(FECHADsd) & "/" & Year(FECHADsd)
        laLOGIK2 = Month(FECHAHst) & "/" & Day(FECHAHst) & "/" & Year(FECHAHst)
        FCIVA.RecordSource = "SELECT TIPO, PCIA, sum(GRINS) as SumGRAV, sum(IVAINS) as SumIVA  FROM FCIVAVTA WHERE FECHA >= " & Separa & laLOGIK1 & _
        Separa & " and FECHA <= " & Separa & laLOGIK2 & " 23:59:59" & Separa & "  group by PCIA, TIPO ORDER by PCIA, TIPO"

    FCIVA.Refresh

If FCIVA.Recordset.EOF Then Exit Sub

Label8.Caption = " Se listarán " & FCIVA.Recordset.RecordCount & " renglones"

FCIVA.Recordset.MoveFirst

Do Until FCIVA.Recordset.EOF

        RSet PPcia = FCIVA.Recordset!PCIA
        
    If FCIVA.Recordset!TIPO = 2 Then
        DbCr = "Cr.-->"
        RSet PGrav = Format(FCIVA.Recordset!SUMGRAV * -1, "###,##0.00")
        RSet PIVAIns = Format(FCIVA.Recordset!SUMIVA * -1, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!SUMGRAV + FCIVA.Recordset!SUMIVA) * -1, "###,##0.00")
    Else
        RSet PGrav = Format(FCIVA.Recordset!SUMGRAV, "###,##0.00")
        RSet PIVAIns = Format(FCIVA.Recordset!SUMIVA, "###,##0.00")
        RSet PTotal = Format((FCIVA.Recordset!SUMGRAV + FCIVA.Recordset!SUMIVA), "###,##0.00")
        DbCr = ""
       End If
       
If FCIVA.Recordset!PCIA = "  " Then FCIVA.Recordset!PCIA = "B "
I4 = Asc(FCIVA.Recordset!PCIA)
d$ = DbCr & "|" & Mid(PCIA(I4 - 65), 2, 10) & "|" & PGrav & "|" & PIVAIns & "|" & PTotal
        
'        FCIVA.Recordset.MoveNext


sinfciva:

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 3
IMPRE.VP1.AddTable s$, h$, "|||||", , , True
        
'   Salto de Hoja
Cont = Cont + 1
If Cont > 50 Then

        RSet TGrav = Format(Tot1, "###,##0.00")
        RSet TIVAIns = Format(Tot2, "###,##0.00")
        RSet TTotal = Format(Tot3, "###,##0.00")

    TitTot = "TRANSPORTE" & Space(40) & TGrav & Space(22) & TIVAIns & Space(17) & TTotal
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "ARIAL", 8.7, 1, 0, 280, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 283, 50, TitTot
    Escribe "ARIAL", 8.7, 1, 0, 286, 5, Rayitas

    IMPRE.VP1.NewPage
    TituIB
End If

        TotItems = TotItems + 1
        
        Tot1 = Tot1 + CCur(PGrav)
        Tot2 = Tot2 + CCur(PIVAIns)
        Tot3 = Tot3 + CCur(PTotal)
         
        FCIVA.Recordset.MoveNext
       
    Loop

        RSet TGrav = Format(Tot1, "###,##0.00")
        RSet TIVAIns = Format(Tot2, "###,##0.00")
        RSet TTotal = Format(Tot3, "###,##0.00")

    TitTot = "TOTAL" & Space(50) & TGrav & Space(22) & TIVAIns & Space(17) & TTotal
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "ARIAL", 8.7, 1, 0, 280, 5, Rayitas
    Escribe "", 8.7, 1, 0, 283, 50, TitTot
    Escribe "ARIAL", 8.7, 1, 0, 286, 5, Rayitas
   
   TitTot = ""
   TotPed = 0: TotItems = 0: Tot1 = 0: Tot2 = 0: Tot3 = 0

' total final

Close #1

 IMPRE.VP1.EndDoc
 
IMPRE.Show 1

End Sub

Private Sub TituIB()

Escribe "ARIAL", 12, 1, 0, 5, 5, NomEmpr
Escribe "ARIAL", 8.7, 0, 0, 10, 5, "CUIT Nº 33-70346790-9"
Escribe "ARIAL", 8.7, 0, 0, 10, 150, "MES : " & Format(FECHA.Value, "mmmm") & " de " & FECHA.Year
Escribe "ARIAL", 10, 1, 0, 10, 95, "Ingresos Brutos"

Escribe "ARIAL", 8.7, 1, 0, 15, 5, Rayitas
Escribe "", 8.7, 0, 0, 20, 1, Space(8) & "cr/db" & Space(37) & "P R O V I N C I A" & Space(35) & _
                 "G R A V A D O" & Space(24) & "I . V . A ." & Space(22) & "T O T A L"

Escribe "", 8.7, 1, 0, 25, 5, Rayitas
    
Escribe "", 8.7, 1, 0, 35, 50, TitTot

TotPed = 0: TotItems = 0

    IMPRE.VP1.TableBorder = tbNone

s$ = "<+22mm|<+60mm|>+35mm|>+35mm|>+35mm"
h$ = "db|Pcia|Grav|IVA|total"

With IMPRE.VP1
    .FontSize = 8.7
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Arial"
    .CurrentY = "40mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "8mm"
    .MarginBottom = "3mm"
End With

Cont = 0

End Sub


'        Comisiones por Cobranzas             - PR250L -     7
Private Sub SelComisCobra()

Label1.Caption = "Comisiones por Cobranzas"
Option2(0).Visible = False
Option2(2).Visible = False
Option2(1).Value = True

FECHA.Visible = True
FECHA.Value = Date
Label4.Visible = True
Label4.Caption = "Fecha :"
        
Label2.Caption = "Vend.:"
Combo2.Visible = False
Combo5.Visible = True

Frame3.Visible = False: Frame5.Visible = False

Label6.Visible = False

End Sub

Private Sub LisComisCobra()
Dim CLTEDsd As Integer, CLTEHst As Integer, Clientes As Integer, I1%
Dim TotItems As Integer, TotImpte As Currency
Dim TotIt As Integer, TotGral As Currency, NomVend As String * 20
Dim PGrav As String * 12, TGrav As String * 12, PItem As String * 3, PCpbte As String * 6
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2

NomVend = Combo5.Text

Tit1 = Chr(14) & NomEmpr & Chr(20) & Space(34) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(26) & Chr(14) & "Comisiones por Cobranzas de " & MonthName(FECHA.Month) & Chr(20)
Tit3 = "FC070L" & Space(14) & "Vendedor : " & NomVend & Space(15) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-123456789-"
Tit4 = "  Cod.Clte.        Razón Social         " & _
       "  Nro.Cpbte.    Fecha       Importe     "

'   Si es por Mes
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
       ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
       End If
       
    FECHADsd = "01/" & FECHA.Month & "/" & FECHA.Year
    FECHAHst = ElDia & "/" & FECHA.Month & "/" & FECHA.Year

CLTE.RecordSource = "SELECT * FROM Clientes WHERE VEND = " & Val(Mid(Combo5.Text, 1, 2)) & _
    " ORDER BY CODIGO"
    
CLTE.Refresh

If CLTE.Recordset.EOF Then Exit Sub

CLTE.Recordset.MoveFirst
    
MsgBox "1"
    'Open "LPT1:" For Output As #1 Len = 226
    Open SaleImpr For Output As #1 Len = 226
    
MsgBox "11"
ImprMatri Chr(27) & "@" & Chr(27) & "CH" & Chr(18) & Chr(20), 0
     
MsgBox "111"

        laLOGIK1 = Month(FECHADsd) & "/" & Day(FECHADsd) & "/" & Year(FECHADsd)
        laLOGIK2 = Month(FECHAHst) & "/" & Day(FECHAHst) & "/" & Year(FECHAHst)
        
        ImprTitulo

    Label8.Caption = " Se listarán " & CLTE.Recordset.RecordCount & " renglones"

Do Until CLTE.Recordset.EOF

        ARTIC.RecordSource = "SELECT * FROM CtasCtes WHERE CLTE = " & CLTE.Recordset!CODIGO & _
        " and FECHA >= " & Separa & laLOGIK1 & Separa & " and FECHA <= " & Separa & laLOGIK2 & Separa & _
        " and TIPO = 4 or TIPO = 5 ORDER BY CLTE, FECHA"
      
        
        ARTIC.Refresh
    
    If ARTIC.Recordset.EOF Then GoTo AlLoop
    
    Label8.Caption = " Se listarán " & ARTIC.Recordset.RecordCount & " renglones"
        
        TitTit = Chr(27) & "E" & CLTE.Recordset!CODIGO & "  " & CLTE.Recordset!Nomb & Chr(27) & "F"
        ImprTit
    
    ARTIC.Recordset.MoveFirst
    
    TotImpte = 0: TotItems = 0
    
    Do Until ARTIC.Recordset.EOF
         
         RSet PCpbte = Format(ARTIC.Recordset!Cpbte, "###,##0")
         RSet PGrav = Format(ARTIC.Recordset!Impte, "###,##0.00")
           
            TitDet = Space(46) & PCpbte & _
                    "   " & ARTIC.Recordset!FECHA & _
                    "   " & PGrav
            ImprDet
            
            TotItems = TotItems + 1
            TotImpte = TotImpte + ARTIC.Recordset!Impte
            
            ARTIC.Recordset.MoveNext
            
            If ARTIC.Recordset.EOF Then Exit Do
    
    Loop
        
        RSet PItem = Format(TotItems, "##0")
        RSet TGrav = Format(TotImpte, "###,##0.00")

        TitTot = Space(40) & " Pagos : " & PItem & "      Importe : " & TGrav
        ImprTot
        
        TotGral = TotImpte + TotGral
        TotIt = TotIt + TotItems
        TotImpte = 0: TotItems = 0
            
AlLoop:
            
            CLTE.Recordset.MoveNext
            
            If CLTE.Recordset.EOF Then Exit Do
            
Loop

' total final
        
        RSet PItem = Format(TotIt, "##0")
        RSet TGrav = Format(TotGral, "###,##0.00")

        TitTot = Space(36) & " Tot.Pagos : " & PItem & "   Impte.Tot. : " & TGrav
        ImprTot
        
        TotGral = TotGral / 1.21
        RSet TGrav = Format(TotGral, "###,##0.00")

        TitTot = Space(61) & "Neto : " & TGrav
        ImprTot

Close #1

Option2(0).Visible = True
Option2(2).Visible = True

End Sub


'        Comisiones por Cobranzas     NUEVO        - PR250L -     7
Private Sub SelComCobNVO()

Label1.Caption = "Comisiones por Cobranzas"
Option2(0).Visible = False
Option2(2).Visible = False
Option2(1).Value = True

FECHA.Visible = True
FECHA.Value = Date
Label4.Visible = True
Label4.Caption = "Fecha :"
        
Label2.Caption = "Vend.:"
Combo2.Visible = False
Combo5.Visible = True

Frame3.Visible = False: Frame5.Visible = False

Label6.Visible = False
Command3.Visible = True

End Sub

Private Sub LisComCobNVO()
Dim CLTEDsd As Integer, CLTEHst As Integer, Clientes As Integer, I1%
Dim TotItems As Integer, TotImpte As Currency, NotDeb As String, TotNDeb As Currency
Dim TotIt As Integer, TotGral As Currency, NomVend As String * 20, CanNDeb As Integer
Dim PGrav As String * 12, TGrav As String * 12, PItem As String * 3, PCpbte As String * 6
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2

FG2.Clear flexClearScrollable
FG2.Rows = 1

NomVend = Combo5.Text

Tit1 = Chr(14) & NomEmpr & Chr(20) & Space(34) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(26) & Chr(14) & "Comisiones por Cobranzas de " & MonthName(FECHA.Month) & Chr(20)
Tit3 = "FC070L" & Space(14) & "Vendedor : " & NomVend & Space(15) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-123456789-"
Tit4 = "  Cod.Clte.        Razón Social         " & _
       "  Nro.Cpbte.    Fecha       Importe     "

'   Si es por Mes
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
       ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
       End If
       
    FECHADsd = "01/" & FECHA.Month & "/" & FECHA.Year
    FECHAHst = ElDia & "/" & FECHA.Month & "/" & FECHA.Year

CLTE.RecordSource = "SELECT * FROM Clientes WHERE VEND = " & Val(Mid(Combo5.Text, 1, 2)) & _
    " ORDER BY CODIGO"
    
CLTE.Refresh

If CLTE.Recordset.EOF Then Exit Sub

CLTE.Recordset.MoveFirst
     
    'Open "LPT1:" For Output As #1 Len = 226
    Open SaleImpr For Output As #1 Len = 226
    
        ImprMatri Chr(27) & "@" & Chr(27) & "CH" & Chr(18) & Chr(20), 0

'        laLOGIK1 = Month(FECHADsd) & "/" & Day(FECHADsd) & "/" & Year(FECHADsd)
'        laLOGIK2 = Month(FECHAHst) & "/" & Day(FECHAHst) & "/" & Year(FECHAHst)
        laLOGIK1 = Format(FECHADsd, "mm/dd/yyyy")
        laLOGIK2 = Format(FECHAHst, "mm/dd/yyyy")
       
        ImprTitulo

    Label8.Caption = " Se listarán " & CLTE.Recordset.RecordCount & " renglones"

Do Until CLTE.Recordset.EOF

        ARTIC.RecordSource = "SELECT * FROM CtasCtes WHERE CLTE = " & CLTE.Recordset!CODIGO & _
        " and FECHA >= " & Separa & laLOGIK1 & Separa & " and FECHA <= " & Separa & laLOGIK2 & Separa & _
        " and ( (TIPO = 4 or TIPO = 5) OR (TIPO = 3  AND MOTI = ' 2'  )  ) ORDER BY CLTE, FECHA"
      
        
        ARTIC.Refresh
    
    If ARTIC.Recordset.EOF Then GoTo AlLoop
    
    Label8.Caption = " Se listarán " & ARTIC.Recordset.RecordCount & " renglones"
        
        TitTit = Chr(27) & "E" & CLTE.Recordset!CODIGO & "  " & CLTE.Recordset!Nomb & Chr(27) & "F"
        ImprTit
    
    ARTIC.Recordset.MoveFirst
    
    TotImpte = 0: TotItems = 0
    
    Do Until ARTIC.Recordset.EOF
         
         RSet PCpbte = Format(ARTIC.Recordset!Cpbte, "###,##0")
         
          NotDeb = "        "
    If ARTIC.Recordset!TIPO = 3 Then
          NotDeb = "(NDeb.) "
            CanNDeb = CanNDeb + 1
            TotNDeb = TotNDeb + ARTIC.Recordset!Impte
'          ARTIC.Recordset!Impte = ARTIC.Recordset!Impte * -1
    End If
            
          RSet PGrav = Format(ARTIC.Recordset!Impte, "###,##0.00")
           TitDet = Space(38) & NotDeb & PCpbte & _
                    "   " & ARTIC.Recordset!FECHA & _
                    "   " & PGrav
                    
    If ARTIC.Recordset!TIPO = 3 Then GoTo NoSuma

            TotItems = TotItems + 1
            TotImpte = TotImpte + ARTIC.Recordset!Impte
            
NoSuma:
    
            ImprDet
            
            ARTIC.Recordset.MoveNext
            If ARTIC.Recordset.EOF Then Exit Do
    
    Loop
        
        RSet PItem = Format(TotItems, "##0")
        RSet TGrav = Format(TotImpte, "###,##0.00")

        TitTot = Space(40) & " Pagos : " & PItem & "      Importe : " & TGrav
        ImprTot
        
        TotGral = TotImpte + TotGral
        TotIt = TotIt + TotItems
        TotImpte = 0: TotItems = 0
            
AlLoop:
            
            CLTE.Recordset.MoveNext
            
            If CLTE.Recordset.EOF Then Exit Do
            
Loop

' total final
        
        RSet PItem = Format(TotIt, "##0")
        RSet TGrav = Format(TotGral, "###,##0.00")

        TitTot = Space(36) & " Tot.Pagos : " & PItem & "   Impte.Tot. : " & TGrav
        ImprTot
        
        RSet PItem = Format(CanNDeb, "##0")
        RSet TGrav = Format(TotNDeb, "###,##0.00")
        TitTot = Space(36) & " Tot.NDeb. : " & PItem & "   Impte.Tot. : " & TGrav
        ImprTot
        
        TotGral = (TotGral - TotNDeb) / 1.21
        RSet TGrav = Format(TotGral, "###,##0.00")

        TitTot = Space(61) & "Neto : " & TGrav
        ImprTot

Close #1

Option2(0).Visible = True
Option2(2).Visible = True

End Sub

'        Comisiones por Cobranzas     NUEVO  a PDF para enviar       - PR250L -     7

Private Sub LisComCobPDF()
Dim CLTEDsd As Integer, CLTEHst As Integer, Clientes As Integer, I1%
Dim TotItems As Integer, TotImpte As Currency, NotDeb As String, TotNDeb As Currency
Dim TotIt As Integer, TotGral As Currency, NomVend As String * 20, CanNDeb As Integer
Dim PGrav As String * 12, TGrav As String * 12, PItem As String * 3, PCpbte As String * 6
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2
Rayitas = String(95, "-")

NomVend = Combo5.Text
NroHoja = 0

Tit1 = NomEmpr & Space(68) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(26) & "Comisiones por Cobranzas de " & MonthName(FECHA.Month)
Tit3 = "FC070L" & Space(24) & "Vendedor : " & NomVend & Space(23) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-123456789-"
Tit4 = "  Cod.Clte.        Razón Social         " & _
       "  Nro.Cpbte.    Fecha       Importe     "

'   Si es por Mes
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
       ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
       End If
       
    FECHADsd = "01/" & FECHA.Month & "/" & FECHA.Year
    FECHAHst = ElDia & "/" & FECHA.Month & "/" & FECHA.Year

If Combo5.ListIndex = 0 Then
    CLTE.RecordSource = "SELECT * FROM Clientes ORDER BY vend,codigo"
Else
    CLTE.RecordSource = "SELECT * FROM Clientes WHERE VEND = " & Val(Mid(Combo5.Text, 1, 2)) & _
                        " ORDER BY CODIGO"
End If

CLTE.Refresh

If CLTE.Recordset.EOF Then Exit Sub

CLTE.Recordset.MoveFirst
  
IMPRE.Caption = " Documento PDF de Comisiones "

IMPRE.VP1.Font.Name = "Arial"
IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprA4
IMPRE.VP1.Orientation = orPortrait
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"

'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 500
        
        laLOGIK1 = Format(FECHADsd, "mm/dd/yyyy")
        laLOGIK2 = Format(FECHAHst, "mm/dd/yyyy")

TituComCob Rengln

    Label8.Caption = " Se listarán " & CLTE.Recordset.RecordCount & " renglones"

Do Until CLTE.Recordset.EOF

        ARTIC.RecordSource = "SELECT * FROM CtasCtes WHERE CLTE = " & CLTE.Recordset!CODIGO & _
        " and FECHA >= " & Separa & laLOGIK1 & Separa & " and FECHA <= " & Separa & laLOGIK2 & Separa & _
        " and ( (TIPO = 4 or TIPO = 5) OR (TIPO = 3  AND MOTI = ' 2'  )  ) ORDER BY CLTE, FECHA"
        
        ARTIC.Refresh
    
    If ARTIC.Recordset.EOF Then GoTo AlLoop
    
    Label8.Caption = " Se listarán " & ARTIC.Recordset.RecordCount & " renglones"
        
'    Imprime nombre del cliente
        TitTit = CLTE.Recordset!CODIGO & "  " & CLTE.Recordset!Nomb
        Escribe "Lucida Console", 10, 1, 0, (Rengln * 5), 20, TitTit
        Rengln = Rengln + 1
        
    ARTIC.Recordset.MoveFirst
    
    TotImpte = 0: TotItems = 0
    
    Do Until ARTIC.Recordset.EOF
         
         RSet PCpbte = Format(ARTIC.Recordset!Cpbte, "###,##0")
         
          NotDeb = "        "
    If ARTIC.Recordset!TIPO = 3 Then
          NotDeb = "(NDeb.) "
            CanNDeb = CanNDeb + 1
            TotNDeb = TotNDeb + ARTIC.Recordset!Impte
'          ARTIC.Recordset!Impte = ARTIC.Recordset!Impte * -1
    End If
            
          RSet PGrav = Format(ARTIC.Recordset!Impte, "###,##0.00")
           TitDet = Space(38) & NotDeb & PCpbte & _
                    "   " & ARTIC.Recordset!FECHA & _
                    "   " & PGrav
                    
    If ARTIC.Recordset!TIPO = 3 Then GoTo NoSuma

            TotItems = TotItems + 1
            TotImpte = TotImpte + ARTIC.Recordset!Impte
            
NoSuma:
    
'    Imprime el detalle
        Escribe "Lucida Console", 9, 0, 0, (Rengln * 5), 5, TitDet
        Rengln = Rengln + 1
        If Rengln > 52 Then
            IMPRE.VP1.MarginBottom = "1mm"
            IMPRE.VP1.NewPage
            TituComCob Rengln
        End If
        
 '---------------------------------
                FG2.AddItem CLTE.Recordset!VEND & vbTab & _
                            CLTE.Recordset!CODIGO & vbTab & CLTE.Recordset!Nomb & vbTab & _
                            ARTIC.Recordset!Cpbte & vbTab & _
                            Format(ARTIC.Recordset!FECHA, "dd/mm/yyyy") & vbTab & _
                            Format(ARTIC.Recordset!Impte, "######0.00"), FG2.Rows
'----------------------------------

            ARTIC.Recordset.MoveNext
            If ARTIC.Recordset.EOF Then Exit Do
    
    Loop
        
        RSet PItem = Format(TotItems, "##0")
        RSet TGrav = Format(TotImpte, "###,##0.00")

        TitTot = Space(40) & " Pagos : " & PItem & "      Importe : " & TGrav
        Escribe "Lucida Console", 9, 1, 0, (Rengln * 5), 5, TitTot
        Rengln = Rengln + 1
        If Rengln > 50 Then
            IMPRE.VP1.MarginBottom = "1mm"
            IMPRE.VP1.NewPage
            TituComCob Rengln
        End If

        TotGral = TotImpte + TotGral
        TotIt = TotIt + TotItems
        TotImpte = 0: TotItems = 0
            
AlLoop:
            
            CLTE.Recordset.MoveNext
            
            If CLTE.Recordset.EOF Then Exit Do
            
Loop

' total final
        
        RSet PItem = Format(TotIt, "##0")
        RSet TGrav = Format(TotGral, "###,##0.00")
        
        Escribe "Lucida Console", 9, 1, 0, (Rengln * 5), 5, Rayitas
        Rengln = Rengln + 2

        TitTot = Space(36) & " Tot.Pagos : " & PItem & "   Impte.Tot. : " & TGrav
        
                 FG2.AddItem "" & vbTab & _
                            "" & vbTab & "" & vbTab & _
                            "Tot.Pagos :" & vbTab & _
                            TotIt & vbTab & _
                            TotGral, FG2.Rows
                   
'    Imprime los totales
        Escribe "Lucida Console", 9, 1, 0, (Rengln * 5), 5, TitTot
        Rengln = Rengln + 1
        
        RSet PItem = Format(CanNDeb, "##0")
        RSet TGrav = Format(TotNDeb, "###,##0.00")
        TitTot = Space(36) & " Tot.NDeb. : " & PItem & "   Impte.Tot. : " & TGrav
        
                 FG2.AddItem "" & vbTab & _
                            "" & vbTab & "" & vbTab & _
                            "Tot.NDeb. :" & vbTab & _
                            CanNDeb & vbTab & _
                            TGrav, FG2.Rows
        
        Escribe "Lucida Console", 9, 1, 0, (Rengln * 5), 5, TitTot
        Rengln = Rengln + 1
        
        TotGral = (TotGral - TotNDeb) / 1.21
        RSet TGrav = Format(TotGral, "###,##0.00")
        
                 FG2.AddItem "" & vbTab & _
                            "" & vbTab & "" & vbTab & _
                            "Neto :" & vbTab & _
                            0 & vbTab & _
                            Format(TotGral, "########0.00"), FG2.Rows

        TitTot = Space(61) & "Neto : " & TGrav
        Escribe "Lucida Console", 9, 1, 0, (Rengln * 5), 5, TitTot


 IMPRE.VP1.EndDoc

VSPDF81.Title = "Comisiones - " & NomVend & "-" & Format(FECHA.Value, "mmmm") & " de " & FECHA.Year
VSPDF81.ConvertDocument IMPRE.VP1, "c:\Comisiones-" & NomVend & Mid(FECHA.Value, 4, 2) & FECHA.Year & ".pdf"
 
'IMPRE.Show 1
Tit4 = "  Cod.Clte.        Razón Social         " & _
       "  Nro.Cpbte.    Fecha       Importe     "

    FG2.TextMatrix(0, 0) = "Vend."
    FG2.TextMatrix(0, 1) = "Cód.Clte."
    FG2.TextMatrix(0, 2) = "Razón Social"
    FG2.TextMatrix(0, 3) = "Nro.Cpbte."
    FG2.TextMatrix(0, 4) = "Fecha"
    FG2.TextMatrix(0, 5) = "Importe"

    AExcel 4, 5

Option2(0).Visible = True
Option2(2).Visible = True
Command1_Click

End Sub

'        Titulos del PDF
Private Sub TituComCob(LaLinea As Double)

NroHoja = NroHoja + 1
        Escribe "Lucida Console", 9, 1, 0, 10, 5, Tit1
        Escribe "Lucida Console", 9, 1, 0, 15, 5, Tit2
        Escribe "Lucida Console", 9, 1, 0, 20, 5, Tit3 & NroHoja
        Escribe "Lucida Console", 9, 1, 0, 25, 5, Rayitas
        Escribe "Lucida Console", 9, 1, 0, 30, 5, Tit4
        Escribe "Lucida Console", 9, 1, 0, 35, 5, Rayitas
        Rengln = 8

End Sub

'        Subdiario de Cobranzas            - CC030L -     8
Private Sub SelSubCobzas()

Label1.Caption = "Subdiario de Cobranzas"
Option2(1).Value = True

Label2.Visible = True
Label6.Visible = False
Combo2.Visible = True

Combo2.SetFocus

End Sub

Private Sub LisSubCobzas()
Dim I1%, LosDias As String * 6
Dim Valor1 As String * 14, Valor2 As String * 14
Dim Dias As Integer, ElTipo As String * 4, Estado As String * 5, ElCpbte As String * 6

TotDebeP = 0: TotDebeT = 0: TotVenc = 0: TotVencT = 0

Asteriscos = String(80, "*")
Rayitas = String(80, "-")

Tit1 = Chr(14) & NomEmpr & Chr(20) & Space(34) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(10) & Chr(14) & "Subdiario de Cobranzas" & Chr(20)

Tit3 = "CC030L" & Space(63) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-"
Tit4 = " Tipo  Nro.    Fecha     Fecha        Im" & _
       "porte         Saldo      " & Chr(15) & "Estado    Dias     U$S" & Chr(18)
Tit5 = "      Cpbte.  Emision  Vencimien.      T" & _
       "otal         a Pagar     " & Chr(15) & "          Venc." & Chr(18)

Select Case OptDos

    Case 0  '   Uno Solo
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE

    Case 1 '   Una Zona
        Tit3 = "CC030L" & Space(26) & "Zona : " & _
                Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & Space(25) & "Hoja Nro. "
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE zona = " & Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2))

    Case 2 ' Todos
        CLTE.RecordSource = "SELECT * FROM Clientes"
End Select
    
CLTE.Refresh

If CLTE.Recordset.EOF Then Exit Sub

CLTE.Recordset.MoveFirst
    
    Label8.Caption = " Se listarán " & CLTE.Recordset.RecordCount & " Clientes"
    
    'Open "LPT1:" For Output As #1 Len = 226
    Open SaleImpr For Output As #1 Len = 226
    ImprMatri Chr(27) & "@" & Chr(27) & "CH" & Chr(18) & Chr(20), 0
        
        ImprTitulo
    
    TotPed = 0: TotItems = 0

Do Until CLTE.Recordset.EOF

    '----------------------------------------------------
    CCTE.RecordSource = "SELECT * FROM CtasCtes WHERE CLTE = " & CLTE.Recordset!CODIGO & _
        " AND ( tipo = 1 or TIPO = 3 ) AND DEBE > 1"
        
        CCTE.Refresh
    'Debug.Print CCTE.RecordSource
    
    If CCTE.Recordset.EOF Then GoTo FinDo1

    '----------------------------------------------------
    '   Imprime datos del Clientes
    
    TitDet = " Cliente : " & CLTE.Recordset!CODIGO & _
                    " - " & CLTE.Recordset!Nomb
    ImprDet
    
    TitDet = "           " & CLTE.Recordset!Dir & _
                    "   Tel.:" & CLTE.Recordset!tel1
    ImprDet
    
    TitDet = "           (" & Mid(CLTE.Recordset!CP, 1, 4) & ") " & _
                    " - " & CLTE.Recordset!Loc
    ImprDet
     
     TitDet = ""
     ImprDet
            
            TotItems = TotItems + 1

    '----------------------------------------------------
    
    CCTE.Recordset.MoveFirst
    
    Do Until CCTE.Recordset.EOF
    
    Select Case CCTE.Recordset!TIPO
        Case 1:      ElTipo = "Fact"
        Case 3:      ElTipo = "N/D."
    End Select
    
    RSet ElCpbte = Format(CCTE.Recordset!Cpbte, "#####0")
    RSet Valor1 = Format(CCTE.Recordset!Impte, "####,###,##0.00")
    RSet Valor2 = Format(CCTE.Recordset!DEBE, "####,###,##0.00")
    
    TitDet = " " & ElTipo & " " & ElCpbte & _
                    " " & Format(CCTE.Recordset!FECHA, "dd/mm/yyyy") & _
                    " " & Format(CCTE.Recordset!FECVTO, "dd/mm/yyyy") & _
                    " " & Valor1 & _
                    " " & Valor2
                    
    Estado = ""
    If CCTE.Recordset!FECVTO < Date Then
        Estado = "Venc."
        Dias = Date - CCTE.Recordset!FECVTO
        RSet LosDias = Format(Dias, "##,##0")
        TitDet = TitDet & " " & Estado & " " & LosDias
        TotVenc = TotVenc + CCTE.Recordset!DEBE
    End If
            
            TotPed = TotPed + 1
    
            ImprDet
            TotDebeP = TotDebeP + CCTE.Recordset!DEBE
    
            CCTE.Recordset.MoveNext
            
        Loop
'-----------------------------------------------------------

     TitDet = ""
     ImprDet
     RSet STRImpre = Format(TotDebeP, "$ ###,###,##0.00")
     TitDet = Space(28) & "Deuda Total   : " & STRImpre
     ImprDet
     RSet STRImpre = Format(TotVenc, "$ ###,###,##0.00")
     TitDet = Space(28) & "Deuda Vencida : " & STRImpre
     ImprDet
     TitDet = "--------------------------------------------------------------------------"
     ImprDet
     TitDet = ""
     ImprDet
     
     TotDebeT = TotDebeT + TotDebeP
     TotDebeP = 0: TotVenc = 0

FinDo1:
            
    CLTE.Recordset.MoveNext

Loop

    TitTot = " Clientes : " & TotItems & "      Comprobantes : " & TotPed
    ImprTot
   
    TotPed = 0: TotItems = 0

' total final

Close #1

End Sub

Private Sub SelSubCobzasAFIP()

Label1.Caption = "Subdiario de Cobranzas"

Option3(1).Value = True

Label6.Visible = True
Check1.Visible = True

End Sub

Private Sub LisSubCobzasAFIP()
Dim I1%
Dim TotItems As Integer, TotPed As Integer
Dim PClte As String * 5, PNomb As String * 25, PCpbte As String * 8, PCUIT As String * 11
Dim PGrav As String * 12, PIVAIns As String * 12, PMonot As String * 10, PExento As String * 10
Dim TGrav As String * 12, TIVAIns As String * 12, TMonot As String * 10, TExento As String * 10
Dim PTotal As String * 13, PPtoVta As String * 4
Dim TPGrav As String * 12, TPIVAIns As String * 12, TPMonot As String * 11, TPExento As String * 12
Dim TPTotal As String * 14, SeniaEOF As Boolean
Dim TotGan As Currency, TotIB As Currency, TOTIVA As Currency, TotSUS As Currency
Dim PTotGan  As String * 10, PTotIB  As String * 10, PTotIVA  As String * 10, PTotSUS  As String * 10
Dim TPTotGan  As String * 10, TPTotIB  As String * 10, TPTotIVA  As String * 10, TPTotSUS  As String * 10
Dim TotGrav As Currency

Dim Tot1 As Currency, Tot2 As Currency, Tot3 As Currency, Tot4 As Currency, Tot5 As Currency, TTotal As Currency
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2
Dim CodIVA As String

FG2.Clear flexClearScrollable
FG2.Rows = 1

'-- FG3: Grilla para el resumen de incobrables
FG3.Clear flexClearScrollable
FG3.Rows = 1


'   Si es por Mes
If OptTres = 1 Then
    FECHADsd = FECHA.Value
    ElDia = 31
    If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
FECHADsd = FECHA.Month & "/" & "01/" & FECHA.Year
FECHAHst = FECHA.Month & "/" & ElDia & "/" & FECHA.Year
    
        If FECHA.Month = 12 Then
            FECHAHst = "01/01/" & FECHA.Year + 1
        Else
            FECHAHst = FECHA.Month + 1 & "/01/" & FECHA.Year
        End If

End If
  
IMPRE.Caption = " Impresión del Subdiario de Cobranzas "

IMPRE.VP1.Font.Name = "Arial"
IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprLegal
IMPRE.VP1.Orientation = orLandscape
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"
Rayitas = String(180, "=")

'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 500

TituSC


' IMPRE.VP1.EndDoc
 
'IMPRE.Show 1

laLOGIK1 = Format(FECHA, "mm/dd/yyyy")
laLOGIK2 = Format(FECHA + 1, "mm/dd/yyyy")

Select Case OptTres
    
    Case 0
        
        CCTE.RecordSource = "SELECT * FROM CTASCTES WHERE TIPO = 4 AND FECHA >= " & Separa & laLOGIK1 & _
        Separa & " and FECHA < " & Separa & laLOGIK2 & Separa & " order by fecha, cpbte"
        
    Case 1
        
        CCTE.RecordSource = "SELECT * FROM CTASCTES WHERE TIPO = 4 AND  FECHA >= " & Separa & FECHADsd & _
        Separa & " and FECHA < " & Separa & FECHAHst & Separa & " order by fecha, cpbte"

End Select

    CCTE.Refresh

If CCTE.Recordset.EOF Then Exit Sub

Label8.Caption = " Se listarán " & CCTE.Recordset.RecordCount & " renglones"

CCTE.Recordset.MoveFirst

Do Until CCTE.Recordset.EOF
        
    CLTE.RecordSource = "SELECT * FROM CLIENTES WHERE CODIGO = " & CCTE.Recordset!CLTE
    CLTE.Refresh
If CLTE.Recordset.EOF Then
        PNomb = "** Desconocido **"
        PCUIT = "N/I"
Else
        RSet PNomb = CLTE.Recordset!Nomb
        RSet PCUIT = CLTE.Recordset!Cuit
        RSet CodIVA = "N/I"
        
    Select Case CLTE.Recordset!CIVA
        Case 1: CodIVA = "R.I."
        Case 2: CodIVA = "RNI."
        Case 3: CodIVA = "C.F."
        Case 4: CodIVA = "Exto"
        Case 5: CodIVA = "Mon."
    End Select
    
End If

        RSet PPtoVta = Format(Val(CCTE.Recordset!prefijo), "0000")
        RSet PCpbte = Format(Val(CCTE.Recordset!Cpbte), "00000000")
        RSet PClte = CCTE.Recordset!CLTE

PIVAIns = "          "
PGrav = "          "
PExento = "          "
'        RSet PGrav = Format((Round((CCTE.Recordset!Impte / (1 + IVAIns)), 2)), "###,##0.00")
'        RSet PIVAIns = Format(CCur(PGrav) * IVAIns, "###,##0.00")
        RSet PTotal = Format((CCTE.Recordset!Impte), "###,##0.00")
        
        TotGrav = CCTE.Recordset!Impte
    
 PTotGan = "": PTotIB = "": PTotIVA = "": PTotSUS = ""
        
        FCIVA.RecordSource = "SELECT * FROM MOVIMVS WHERE CLTE =  " & CCTE.Recordset!CLTE & _
         " and CPBTE = " & CCTE.Recordset!Cpbte & " order by TIPREG"
    FCIVA.Refresh

If FCIVA.Recordset.EOF Then GoTo SaleLoop
    
FCIVA.Recordset.MoveFirst
  Do Until FCIVA.Recordset.EOF
     Select Case FCIVA.Recordset!TIPREG
        Case 1
             RSet PTotGan = Format(FCIVA.Recordset!Impte, "###,##0.00")
        Case 2
             RSet PTotIB = Format(FCIVA.Recordset!Impte, "###,##0.00")
        Case 3
             RSet PTotIVA = Format(FCIVA.Recordset!Impte, "###,##0.00")
        Case 4
             RSet PTotSUS = Format(FCIVA.Recordset!Impte, "###,##0.00")
     End Select

If FCIVA.Recordset!TIPREG < 5 Then TotGrav = TotGrav - FCIVA.Recordset!Impte
    
'  Si es tipo=9 (Incobrable) lo carga en FG3
            If FCIVA.Recordset!TIPREG = 9 Then
                FG3.AddItem Format(CCTE.Recordset!FECHA, "dd/mm/yyyy") & vbTab & _
                            CCTE.Recordset!Letra & vbTab & _
                            PPtoVta & "-" & PCpbte & vbTab & _
                            PNomb & vbTab & _
                            Format(Val(PCUIT), "00-00000000-0") & vbTab & _
                            CodIVA & vbTab & _
                            PTotal, FG3.Rows
            End If
    
FCIVA.Recordset.MoveNext
  
  Loop

SaleLoop:

RSet TGrav = Format(TotGrav, "###,##0.00")

d$ = Format(CCTE.Recordset!FECHA, "dd/mm/yyyy") & "|" & CCTE.Recordset!Letra & "|" & _
           PPtoVta & "-" & PCpbte & "|" & PNomb & "|" & _
           Format(Val(PCUIT), "00-00000000-0") & "|" & CodIVA & "|" & TGrav & "|" & _
           PTotGan & "|" & PTotIB & "|" & PTotIVA & "|" & PTotSUS & "|" & PTotal
                
            If Check1.Value = 1 Then
                FG2.AddItem Format(CCTE.Recordset!FECHA, "dd/mm/yyyy") & vbTab & _
                            CCTE.Recordset!Letra & vbTab & _
                            PPtoVta & "-" & PCpbte & vbTab & _
                            PNomb & vbTab & _
                            Format(Val(PCUIT), "00-00000000-0") & vbTab & _
                            CodIVA & vbTab & _
                            TGrav & vbTab & _
                            PTotGan & vbTab & _
                            PTotIB & vbTab & _
                            PTotIVA & vbTab & _
                            PTotSUS & vbTab & _
                            PTotal, FG2.Rows
            End If
        
'  Suma Totales
        If Val(TGrav) <> 0 Then Tot1 = Tot1 + CCur(TGrav)
        If Val(PTotGan) <> 0 Then Tot2 = Tot2 + CCur(PTotGan)
        If Val(PTotIB) <> 0 Then Tot3 = Tot3 + CCur(PTotIB)
        If Val(PTotIVA) <> 0 Then Tot4 = Tot4 + CCur(PTotIVA)
        If Val(PTotSUS) <> 0 Then Tot5 = Tot5 + CCur(PTotSUS)
        TTotal = TTotal + CCur(PTotal)

        TotItems = TotItems + 1
        
        CCTE.Recordset.MoveNext

sinCCTE:

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 3
IMPRE.VP1.AddTable s$, h$, "|||||||", , , True
        
'   Salto de Hoja
Cont = Cont + 1
If Cont > 30 Then

TPTotGan = "          "
TGrav = "          "
TPTotIB = "          "
TPTotIVA = "          "
TPTotSUS = "          "

        RSet TGrav = Format(Tot1, "###,##0.00")
        RSet TPTotGan = Format(Tot2, "###,##0.00")
        RSet TPTotIB = Format(Tot3, "###,##0.00")
        RSet TPTotIVA = Format(Tot4, "###,##0.00")
        RSet TPTotSUS = Format(Tot5, "###,##0.00")
        RSet TPTotal = Format(TTotal, "###,##0.00")

    TitTot = "TRANSPORTE" & Space(56) & TGrav & Space(3) & TPTotGan & Space(3) & TPTotIB & Space(3) & _
                                        TPTotIVA & Space(3) & TPTotSUS & Space(3) & TPTotal
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 203, 5, Rayitas

    IMPRE.VP1.NewPage
    TituSC
End If
        
  'If CCTE.Recordset!CIVA = 4 Then
  '  Tot4 = 0
  '  End If
    
        
    Loop

TIVAIns = "          "
TGrav = "          "
TExento = "          "

        RSet TGrav = Format(Tot1, "###,##0.00")
        RSet TPTotGan = Format(Tot2, "###,##0.00")
        RSet TPTotIB = Format(Tot3, "###,##0.00")
        RSet TPTotIVA = Format(Tot4, "###,##0.00")
        RSet TPTotSUS = Format(Tot5, "###,##0.00")
        RSet TPTotal = Format(TTotal, "###,##0.00")


    TitTot = "Total Gral" & Space(56) & TGrav & Space(3) & TPTotGan & Space(3) & TPTotIB & Space(3) & _
                                        TPTotIVA & Space(3) & TPTotSUS & Space(3) & TPTotal
    
    IMPRE.VP1.MarginBottom = "1mm"
        
    Escribe "Arial", 8.7, 1, 0, 195, 5, Rayitas
    Escribe "Lucida Console", 8.7, 1, 0, 198, 50, TitTot
    Escribe "Arial", 8.7, 1, 0, 201, 5, Rayitas
   
    TotPed = 0: TotItems = 0: Tot1 = 0: Tot2 = 0: Tot3 = 0: Tot4 = 0:: Tot5 = 0: TTotal = 0: Cont = 0
    TitTot = ""
' total final

Close #1

'-------------------------------------------------------
'Cuadro Resumen
'-----------------
Dim A1 As Integer, A2 As Single, Incobr1 As String * 200, Incobr2 As String * 200, Incobr3 As String * 30
Dim TotIncobr As Currency

If FG3.Rows = 1 Then GoTo SinResumen

    IMPRE.VP1.NewPage
    TituSC
    
    Escribe "Lucida Console", 14, 1, 0, 40, 150, "Resumen de Incobrables"
    Escribe "Lucida Console", 12, 1, 1, 55, 90, "Fecha   Comprobante       Razón Social           CUIT    IVA              Total"


A2 = 55

    For A1 = 1 To FG3.Rows - 1
    
        Incobr1 = Format(FG3.TextMatrix(A1, 0), "dd/mm/yyyy") & Space(3) & FG3.TextMatrix(A1, 1) & Space(3) & _
                   FG3.TextMatrix(A1, 2) & Space(3) & FG3.TextMatrix(A1, 3)
        Incobr2 = Format(FG3.TextMatrix(A1, 4), "00-00000000-0") & Space(5) & FG3.TextMatrix(A1, 5)
              
              A2 = A2 + 7
              
              RSet Incobr3 = Format(FG3.TextMatrix(A1, 6), "###,###,##0.00")
              
              TotIncobr = TotIncobr + FG3.TextMatrix(A1, 6)
              
    Escribe "Arial", 10, 0, 0, A2, 90, Incobr1
    Escribe "Arial", 10, 0, 0, A2, 210, Incobr2
    Escribe "Arial", 10, 0, 0, A2, 260, Incobr3

    Next A1
    
              A2 = A2 + 8
              RSet Incobr3 = Format(TotIncobr, "###,###,##0.00")
     Escribe "Arial", 10, 1, 1, A2, 260, Incobr3
   
SinResumen:


'-------------------------------------------------------

 IMPRE.VP1.EndDoc

'VSPDF81.Title = "Subdiario de Ventas - " & Format(FECHA.Value, "mmmm") & " de " & FECHA.Year
'VSPDF81.ConvertDocument IMPRE.VP1, "c:\Subdiario-" & Mid(FECHA.Value, 4, 2) & FECHA.Year & ".pdf"
 
IMPRE.Show 1

If Check1.Value = 1 Then
    
    FG2.TextMatrix(0, 0) = "Fecha"
    FG2.TextMatrix(0, 1) = "Letra"
    FG2.TextMatrix(0, 2) = "Cpbte."
    FG2.TextMatrix(0, 3) = "Razón Social"
    FG2.TextMatrix(0, 4) = "CUIT"
    FG2.TextMatrix(0, 5) = "Cond.IVA"
    FG2.TextMatrix(0, 6) = "Gravado"
    FG2.TextMatrix(0, 7) = "Ganancias"
    FG2.TextMatrix(0, 8) = "I.B."
    FG2.TextMatrix(0, 9) = "IVA"
    FG2.TextMatrix(0, 10) = "SUS"
    FG2.TextMatrix(0, 11) = "Total"
    
    AExcel 2, 11
End If

End Sub

Private Sub TituSC()

Escribe "Arial", 12, 1, 0, 5, 100, NomEmpr
Escribe "Arial", 8.7, 0, 0, 10, 5, "CUIT Nº 33-70346790-9"
Escribe "Arial", 8.7, 0, 0, 10, 150, "MES : " & Format(FECHA.Value, "mmmm") & " de " & FECHA.Year
Escribe "Arial", 8.7, 1, 0, 10, 275, "SUBDIARIO COBRANZAS"

Escribe "Arial", 8.7, 1, 0, 15, 5, Rayitas
Escribe "", 8.7, 0, 0, 20, 5, Space(11) & "FECHA" & Space(5) & "TIP" & Space(4) & "COMPROBANTE Nº" & _
                Space(50) & "C L I E N T E" & Space(60) & "P A G O S" & Space(50) & "R E T E N C I O N E S" & _
                 Space(54) & "TOTAL"
Escribe "", 8.7, 0, 0, 25, 90, "NOMBRE" & Space(34) & "CUIT" & Space(14) & "COND.IVA" & Space(15) & _
                "EFECT./CHEQUE" & Space(15) & "GANANC." & Space(10) & "I.BRUTOS" & Space(10) & "IVA" & Space(10) & _
                "SUS" & Space(38) & "RECIBO"

Escribe "", 8.7, 1, 0, 30, 5, Rayitas
    
Escribe "Lucida Console", 8.7, 1, 0, 35, 50, TitTot

TotPed = 0: TotItems = 0

    IMPRE.VP1.TableBorder = tbNone

s$ = ">+22mm|^+5mm|^+37mm|<+55mm|<+28mm|^+15mm|>+35mm|>+25mm|>+25mm|>+25mm|>+25mm|>+35mm"
h$ = "fecha|letra|ptovta-cpbte|clte|cuit|civa|pago|gan|ib|iva|sus|total"

With IMPRE.VP1
    .FontSize = 8.7
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Lucida Console"
    .CurrentY = "40mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "8mm"
    .MarginBottom = "3mm"
End With

Cont = 0

End Sub
'        Comisiones Ventas                  - PR900L -     6
Private Sub SelComisVtas()

End Sub

Private Sub LisComisVtas()

End Sub

'        Deuda Vencida       - PR120L -      9
Private Sub SelDeudaVenc()

End Sub

Private Sub LisDeudaVenc()

End Sub

'        Planilla de Cobranzas (Facturas Vencidas y a Vencer) - CC070L -     10
Private Sub SelPlaniCobra()

Label1.Caption = "Planilla de Cobranzas"
Option2(1).Value = True

Label2.Visible = True
Label6.Visible = False
Combo2.Visible = True

Label3.Visible = True: Text1.Visible = True
Command3.Visible = True

Combo2.SetFocus
Check1.Visible = True

End Sub

Private Sub LisPlaniCobra()
Dim I1%, LosDias As String * 6
Dim Valor1 As String * 14, Valor2 As String * 14
Dim Dias As Integer, ElTipo As String * 4, Estado As String * 5, ElCpbte As String * 6

FG2.Clear flexClearScrollable
FG2.Rows = 1

TotDebeP = 0: TotDebeT = 0: TotVenc = 0: TotVencT = 0

Asteriscos = String(132, "*")
Rayitas = String(132, "-")

Tit1 = Chr(15) & Chr(14) & NomEmpr & Chr(20) & Space(86) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(36) & Chr(14) & "Planilla de Cobranzas" & Chr(20)

Tit3 = "CC070L" & Space(112) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-"
Tit4 = "* Clte.*         Razon Social         *     " & _
       "  Direcci¢n         *   Localidad   *C.P.* I" & _
       ".V.A. *    CUIT   *     Telefono / Fax     *"
       
Tit5 = "*                                           " & _
       "  * Tipo * Cpbte. *    Fecha   *        *   " & _
       "   Importe     * Dias *                    *"

Ordenar = " ORDER BY NOMB"

Select Case OptDos

    Case 0  '   Uno Solo
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE & Ordenar

    Case 1 '   Una Zona
        Tit3 = "CC070L" & Space(52) & "Zona : " & _
                Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & Space(51) & "Hoja Nro. "
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE zona = " & _
                            Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & Ordenar

    Case 2 ' Todos
        CLTE.RecordSource = "SELECT * FROM Clientes" & Ordenar
End Select
    
CLTE.Refresh

If CLTE.Recordset.EOF Then Exit Sub

CLTE.Recordset.MoveFirst
    
    Label8.Caption = " Se listarán " & CLTE.Recordset.RecordCount & " Clientes"
    
    'Open "LPT1:" For Output As #1 Len = 226
    Open SaleImpr For Output As #1 Len = 226
    ImprMatri Chr(27) & "@" & Chr(27) & "CH" & Chr(18) & Chr(20), 0
        
        ImprTitulo
    
    TotPed = 0: TotItems = 0

Do Until CLTE.Recordset.EOF

    '----------------------------------------------------
    CCTE.RecordSource = "SELECT * FROM CtasCtes WHERE CLTE = " & CLTE.Recordset!CODIGO & _
        " AND ( tipo = 1 or TIPO = 3 ) AND DEBE > 1 ORDER BY FECHA, CPBTE, TIPO"
        
        CCTE.Refresh
    'Debug.Print CCTE.RecordSource
    
    If CCTE.Recordset.EOF Then GoTo FinDo1

    '----------------------------------------------------
    '   Imprime datos del Clientes
    
    TitDet = " Cliente : " & CLTE.Recordset!CODIGO & _
             " - " & Mid(CLTE.Recordset!Nomb, 1, 30) & _
             "    Tel.:" & Mid(CLTE.Recordset!tel1, 1, 15)
    ImprDet
    
    TitDet = "           " & Mid(CLTE.Recordset!Dir, 1, 25) & _
             "  (" & Mid(CLTE.Recordset!CP, 1, 4) & ") " & _
             " - " & CLTE.Recordset!Loc
    ImprDet
    
    TitDet = "Zona: " & CLTE.Recordset!ZONA & "   IVA " & Combo4.List(CLTE.Recordset!CIVA - 1)
    TitDet = TitDet & "   CUIT : " & CLTE.Recordset!Cuit
    ImprDet
     
            TotItems = TotItems + 1

    '----------------------------------------------------
    
    CCTE.Recordset.MoveFirst
    
    Do Until CCTE.Recordset.EOF
    
    Select Case CCTE.Recordset!TIPO
        Case 1:      ElTipo = "Fact"
        Case 3:      ElTipo = "N/D."
    End Select
    
    RSet ElCpbte = Format(CCTE.Recordset!Cpbte, "#####0")
    RSet Valor1 = Format(CCTE.Recordset!Impte, "####,###,##0.00")
    RSet Valor2 = Format(CCTE.Recordset!DEBE, "###,###,##0.00")
    
    TitDet = Space(48) & ElTipo & "   " & ElCpbte & _
                    "   " & Format(CCTE.Recordset!FECHA, "dd/mm/yyyy") & _
                    Space(15) & Valor2
  '                  "   " & Format(CCTE.Recordset!FECVTO, "dd/mm/yyyy") & _

    If CCTE.Recordset!FECVTO >= Date Then
        Dias = CCTE.Recordset!FECVTO - Date
        If Dias > Text1.Text Then GoTo AOtro
    End If
    
    If CCTE.Recordset!FECVTO < Date Then
        Dias = Date - CCTE.Recordset!FECVTO
        RSet LosDias = Format(Dias, "##,##0")
        TitDet = TitDet & "   " & LosDias
        TotVenc = TotVenc + CCTE.Recordset!DEBE
    End If
                 
            If Check1.Value = 1 Then
                FG2.AddItem CLTE.Recordset!CODIGO & vbTab & _
                            Mid(CLTE.Recordset!Nomb, 1, 30) & vbTab & _
                            Mid(CLTE.Recordset!tel1, 1, 15) & vbTab & _
                            CLTE.Recordset!ZONA & vbTab & _
                            Combo4.List(CLTE.Recordset!CIVA - 1) & vbTab & _
                            CLTE.Recordset!Cuit & vbTab & _
                            ElTipo & vbTab & _
                            ElCpbte & vbTab & _
                            Format(CCTE.Recordset!FECHA, "dd/mm/yyyy") & vbTab & _
                            Valor2 & vbTab & _
                            LosDias, FG2.Rows
            End If
           
            TotPed = TotPed + 1
    
            ImprDet
            TotDebeP = TotDebeP + CCTE.Recordset!DEBE
AOtro:
    
    CCTE.Recordset.MoveNext
            
        Loop
'-----------------------------------------------------------

     TitDet = ""
     ImprDet
     RSet STRImpre = Format(TotDebeP, "$ ###,###,##0.00")
     TitDet = Space(67) & "Deuda Total   : " & STRImpre
     
'     RSet STRImpre = Format(TotVenc, "$ ###,###,##0.00")
'     TitDet = TitDet & Space(5) & "Deuda Vencida : " & STRImpre
     ImprDet
     TitDet = Rayitas
     ImprDet
     TitDet = ""
     ImprDet
     
     TotDebeT = TotDebeT + TotDebeP
     TotVencT = TotVencT + TotVenc
     TotDebeP = 0: TotVenc = 0

FinDo1:
            
    CLTE.Recordset.MoveNext

Loop

    TitTot = " Clientes : " & TotItems & "      Comprobantes : " & TotPed
    ImprTot

     TitDet = ""
     ImprDet
     RSet STRImpre = Format(TotDebeT, "$ ###,###,##0.00")
     TitDet = Space(13) & "Deuda Total   : " & STRImpre
     ImprDet
     RSet STRImpre = Format(TotVencT, "$ ###,###,##0.00")
     TitDet = Space(13) & "Deuda Vencida : " & STRImpre
     ImprDet
     TitDet = Rayitas
     ImprDet
     TitDet = "" & Chr(18) & Chr(12)
     ImprDet
   
    TotPed = 0: TotItems = 0

' total final

Close #1

If Check1.Value = 1 Then

    FG2.TextMatrix(0, 0) = "Código"
    FG2.TextMatrix(0, 1) = "Razón Social"
    FG2.TextMatrix(0, 2) = "Teléfono"
    FG2.TextMatrix(0, 3) = "Zona"
    FG2.TextMatrix(0, 4) = "IVA"
    FG2.TextMatrix(0, 5) = "CUIT"
    FG2.TextMatrix(0, 6) = "Tipo Cpbte."
    FG2.TextMatrix(0, 7) = "Nº Cpbte."
    FG2.TextMatrix(0, 8) = "Fecha"
    FG2.TextMatrix(0, 9) = "Importe"
    FG2.TextMatrix(0, 10) = "Dias V."

    AExcel 7, 10
End If

Label3.Visible = False: Text1.Visible = False

End Sub

Private Sub LisPlaniCobraPDF()
Dim I1%, LosDias As String * 6
Dim Valor1 As String * 14, Valor2 As String * 14
Dim Dias As Integer, ElTipo As String * 4, Estado As String * 5, ElCpbte As String * 6

TotDebeP = 0: TotDebeT = 0: TotVenc = 0: TotVencT = 0
NroHoja = 0

Asteriscos = String(107, "*")
Rayitas = String(107, "-")

Tit1 = NomEmpr & Space(79) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(26) & "Planilla de Cobranzas"

Tit3 = "CC070L" & Space(45) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-"
Tit4 = "* Clte.*         Razon Social         *     " & _
       "  Direcci¢n         *   Localidad   *C.P.* I" & _
       ".V.A. *    CUIT   *     Telefono / Fax     *"
       
Tit5 = "*                   " & _
       "  * Tipo * Cpbte. *    Fecha   *        *   " & _
       "   Importe       *   Dias   *             *"

Ordenar = " ORDER BY NOMB"

Select Case OptDos
    Case 0  '   Uno Solo
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE & Ordenar

    Case 1 '   Una Zona
        Tit3 = "CC070L" & Space(40) & "Zona : " & _
                Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & Space(41) & "Hoja Nro. "
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE zona = " & _
                            Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & Ordenar
    Case 2 ' Todos
        CLTE.RecordSource = "SELECT * FROM Clientes" & Ordenar
End Select
    
CLTE.Refresh

If CLTE.Recordset.EOF Then Exit Sub

CLTE.Recordset.MoveFirst
  
IMPRE.Caption = " Documento PDF de Cobranzas "

IMPRE.VP1.Font.Name = "Arial"
IMPRE.VP1.Copies = 1

IMPRE.VP1.PaperSize = pprA4
IMPRE.VP1.Orientation = orPortrait
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
'IMPRE.VP1.MarginLeft = "1cm"

'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = 500

TituPlanCobr Rengln
    
    Label8.Caption = " Se listarán " & CLTE.Recordset.RecordCount & " Clientes"
    
    TotPed = 0: TotItems = 0

Do Until CLTE.Recordset.EOF

    '----------------------------------------------------
    CCTE.RecordSource = "SELECT * FROM CtasCtes WHERE CLTE = " & CLTE.Recordset!CODIGO & _
        " AND ( tipo = 1 or TIPO = 3 ) AND DEBE > 1 ORDER BY FECHA, CPBTE, TIPO"
        
        CCTE.Refresh
    'Debug.Print CCTE.RecordSource
    
    If CCTE.Recordset.EOF Then GoTo FinDo1

    '----------------------------------------------------
    '   Imprime datos del Clientes
        
'    Imprime nombre del cliente
        TitTit = " Cliente : " & CLTE.Recordset!CODIGO & _
             " - " & Mid(CLTE.Recordset!Nomb, 1, 30) & _
             "    Tel.:" & Mid(CLTE.Recordset!tel1, 1, 15)
        Escribe "Lucida Console", 10, 1, 0, (Rengln * 5), 20, TitTit
        Rengln = Rengln + 1: SaltoPlCobr (Rengln)
    
    
    TitTit = "           " & Mid(CLTE.Recordset!Dir, 1, 25) & _
             "  (" & Mid(CLTE.Recordset!CP, 1, 4) & ") " & _
             " - " & CLTE.Recordset!Loc
        Escribe "Lucida Console", 10, 1, 0, (Rengln * 5), 20, TitTit
        Rengln = Rengln + 1: SaltoPlCobr (Rengln)
    
    TitTit = "Zona: " & CLTE.Recordset!ZONA & "   IVA " & Combo4.List(CLTE.Recordset!CIVA - 1)
    TitTit = TitTit & "   CUIT : " & CLTE.Recordset!Cuit
        Escribe "Lucida Console", 10, 1, 0, (Rengln * 5), 20, TitTit
        Rengln = Rengln + 2: SaltoPlCobr (Rengln)
     
            TotItems = TotItems + 1

    '----------------------------------------------------
    
    CCTE.Recordset.MoveFirst
    
    Do Until CCTE.Recordset.EOF
    
    Select Case CCTE.Recordset!TIPO
        Case 1:      ElTipo = "Fact"
        Case 3:      ElTipo = "N/D."
    End Select
    
    RSet ElCpbte = Format(CCTE.Recordset!Cpbte, "#####0")
    RSet Valor1 = Format(CCTE.Recordset!Impte, "####,###,##0.00")
    RSet Valor2 = Format(CCTE.Recordset!DEBE, "###,###,##0.00")
    
    TitDet = Space(25) & ElTipo & "   " & ElCpbte & _
                    "   " & Format(CCTE.Recordset!FECHA, "dd/mm/yyyy") & _
                    Space(15) & Valor2
  '                  "   " & Format(CCTE.Recordset!FECVTO, "dd/mm/yyyy") & _

    If CCTE.Recordset!FECVTO >= Date Then
        Dias = CCTE.Recordset!FECVTO - Date
        If Dias > Text1.Text Then GoTo AOtro
    End If
    
    If CCTE.Recordset!FECVTO < Date Then
        Dias = Date - CCTE.Recordset!FECVTO
        RSet LosDias = Format(Dias, "##,##0")
        TitDet = TitDet & "   " & LosDias
        TotVenc = TotVenc + CCTE.Recordset!DEBE
    End If
            
            TotPed = TotPed + 1
    
        Escribe "Lucida Console", 8, 0, 0, (Rengln * 5), 5, TitDet
        Rengln = Rengln + 1: SaltoPlCobr (Rengln)
        
            TotDebeP = TotDebeP + CCTE.Recordset!DEBE
AOtro:
    
    CCTE.Recordset.MoveNext
            
        Loop
'-----------------------------------------------------------

     TitDet = ""
        Escribe "Lucida Console", 8, 0, 0, (Rengln * 5), 5, TitDet
        Rengln = Rengln + 1: SaltoPlCobr (Rengln)
     
     RSet STRImpre = Format(TotDebeP, "$ ###,###,##0.00")
     TitDet = Space(44) & "Deuda Total   : " & STRImpre
     
'     RSet STRImpre = Format(TotVenc, "$ ###,###,##0.00")
'     TitDet = TitDet & Space(5) & "Deuda Vencida : " & STRImpre
        Escribe "Lucida Console", 8, 0, 0, (Rengln * 5), 5, TitDet
        Rengln = Rengln + 1: SaltoPlCobr (Rengln)
     
     TitDet = Rayitas
        Escribe "Lucida Console", 8, 0, 0, (Rengln * 5), 5, TitDet
        Rengln = Rengln + 1: SaltoPlCobr (Rengln)
     
     TotDebeT = TotDebeT + TotDebeP
     TotVencT = TotVencT + TotVenc
     TotDebeP = 0: TotVenc = 0

FinDo1:
            
    CLTE.Recordset.MoveNext

Loop

    TitTot = " Clientes : " & TotItems & "      Comprobantes : " & TotPed
        Escribe "Lucida Console", 8, 0, 0, (Rengln * 5), 5, TitTot
        Rengln = Rengln + 2: SaltoPlCobr (Rengln)

     RSet STRImpre = Format(TotDebeT, "$ ###,###,##0.00")
     TitDet = Space(13) & "Deuda Total   : " & STRImpre
        Escribe "Lucida Console", 8, 0, 0, (Rengln * 5), 5, TitDet
        Rengln = Rengln + 1: SaltoPlCobr (Rengln)
        
     RSet STRImpre = Format(TotVencT, "$ ###,###,##0.00")
     TitDet = Space(13) & "Deuda Vencida : " & STRImpre
        Escribe "Lucida Console", 8, 0, 0, (Rengln * 5), 5, TitDet
        Rengln = Rengln + 1: SaltoPlCobr (Rengln)
        
     TitDet = Rayitas
        Escribe "Lucida Console", 8, 0, 0, (Rengln * 5), 5, TitDet
        Rengln = Rengln + 1: SaltoPlCobr (Rengln)
   
    TotPed = 0: TotItems = 0

' total final

 IMPRE.VP1.EndDoc

VSPDF81.Title = "Planilla Cobranza - Zona " & Mid(Combo2.List(Combo2.ListIndex), 1, 2)
VSPDF81.ConvertDocument IMPRE.VP1, "c:\Planilla Cobr-" & Mid(Combo2.List(Combo2.ListIndex), 1, 2) & ".pdf"

Label3.Visible = False: Text1.Visible = False

End Sub

Private Function SaltoPlCobr(Renglon As Double)

        If Renglon > 52 Then
            IMPRE.VP1.MarginBottom = "1mm"
            IMPRE.VP1.NewPage
            TituPlanCobr Renglon
            
        End If

End Function
'        Titulos del PDF
Private Sub TituPlanCobr(LaLinea As Double)

NroHoja = NroHoja + 1
        Escribe "Lucida Console", 14, 1, 0, 8, 5, NomEmpr
        Escribe "Lucida Console", 8, 1, 0, 10, 165, "Fecha " & Format(Date, "dd/mm/yyyy")
        Escribe "Lucida Console", 12, 1, 0, 15, 5, Tit2
        Escribe "Lucida Console", 8, 1, 0, 20, 5, Tit3 & NroHoja
        Escribe "Lucida Console", 8, 1, 0, 25, 5, Rayitas
        Escribe "Lucida Console", 8, 1, 0, 30, 5, Tit5
        Escribe "Lucida Console", 8, 1, 0, 35, 5, Rayitas
        Rengln = 9

End Sub

'        Estado de Cuenta             - CC020L -      11
Private Sub SelCCTE()

Label1.Caption = "Estado de Cuenta"
Option2(1).Value = True

Label2.Visible = True
Label6.Visible = False
Combo2.Visible = True

Combo2.SetFocus

End Sub

Private Sub LisCCTE()
Dim I1%, LosDias As String * 6
Dim Valor1 As String * 14, Valor2 As String * 14, ElDebe As String * 13, ElHaber As String * 13
Dim ElSaldo As String * 13, ElResto As String * 12, Saldo As Currency
Dim Dias As Integer, TipoCpbt As String * 4, Estado As String * 5, ElCpbte As String * 6
Dim Impu1 As String * 2, Impu2 As String * 2, FecDsd As Date, SaldAnt As Currency

TotDebeP = 0: TotDebeT = 0: TotVenc = 0: TotVencT = 0
FecDsd = Date - 1095     '  Tres años para atrás

Asteriscos = String(80, "*")
Rayitas = String(80, "-")

Tit1 = Chr(14) & NomEmpr & Chr(20) & Space(34) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(10) & Chr(14) & "Estado de Cuenta" & Chr(20)

Tit3 = "CC020L" & Space(63) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-"
Tit4 = "*  Fecha *Tipo*Cpbte.*I1*I2*    Resto   " & _
       "*    Debe    *    Haber   *    Saldo   *"
Tit5 = ""
Select Case OptDos

    Case 0  '   Uno Solo
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE

    Case 1 '   Una Zona
        Tit3 = "CC020L" & Space(26) & "Zona : " & _
                Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & Space(25) & "Hoja Nro. "
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE zona = " & Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2))

    Case 2 ' Todos
        CLTE.RecordSource = "SELECT * FROM Clientes"
End Select
    
CLTE.Refresh

If CLTE.Recordset.EOF Then Exit Sub

CLTE.Recordset.MoveFirst
    
    Label8.Caption = " Se listarán " & CLTE.Recordset.RecordCount & " Clientes"
    
    'Open "LPT1:" For Output As #1 Len = 226
    Open SaleImpr For Output As #1 Len = 226
    ImprMatri Chr(27) & "@" & Chr(27) & "CH" & Chr(18) & Chr(20), 0
        
        ImprTitulo
    
    TotPed = 0: TotItems = 0

Do Until CLTE.Recordset.EOF

    '----------------------------------------------------
    CCTE.RecordSource = "SELECT * FROM CtasCtes WHERE CLTE = " & CLTE.Recordset!CODIGO

        CCTE.Refresh
    'Debug.Print CCTE.RecordSource
    
    If CCTE.Recordset.EOF Then GoTo FinDo1

    '----------------------------------------------------
    '   Imprime datos del Clientes
    
    TitDet = " Cliente : " & CLTE.Recordset!CODIGO & _
                    " - " & CLTE.Recordset!Nomb
    ImprDet
    
    TitDet = "           " & CLTE.Recordset!Dir & _
                    "   Tel.:" & CLTE.Recordset!tel1
    ImprDet
    
    TitDet = "           (" & Mid(CLTE.Recordset!CP, 1, 4) & ") " & _
                    " - " & CLTE.Recordset!Loc
    ImprDet
     
     TitDet = ""
     ImprDet
            
            TotItems = TotItems + 1

    '----------------------------------------------------
    
    CCTE.Recordset.MoveFirst
    SaldAnt = 0
    
    Do Until CCTE.Recordset.EOF
    
    Select Case CCTE.Recordset!TIPO
        Case 0:        TipoCpbt = "Sdo.A."
        Case 1:        TipoCpbt = "Fact"
        Case 2:        TipoCpbt = "N/C."
        Case 3:        TipoCpbt = "N/D."
        Case 4:        TipoCpbt = "Rec."
        Case 5:        TipoCpbt = "P/C."
        Case 6:        TipoCpbt = "Desc"
        Case 7:        TipoCpbt = "Ctdo"
        Case 8:        TipoCpbt = " -- "
        Case 9:        TipoCpbt = "Anul."
    End Select
    
    RSet ElDebe = Format(0, "#####,##0.00"): RSet ElHaber = Format(0, "#####,##0.00"): RSet ElResto = Format(0, "#####,##0.00")
        
    
    If CCTE.Recordset!FECHA >= FecDsd And SaldAnt <> 0 Then
        RSet ElSaldo = Format(Saldo, "###,##0.00")
        TitDet = "  ++ Saldo Anterior ++" & Space$(44) & ElSaldo
            ImprDet
        SaldAnt = 0
    End If
    
    If CCTE.Recordset!TIPO = 1 Or CCTE.Recordset!TIPO = 3 Or CCTE.Recordset!TIPO = 7 Then
        RSet ElResto = Format(CCTE.Recordset!DEBE, "#####,##0.00")
        RSet ElDebe = Format(CCTE.Recordset!Impte, "#####,##0.00")
        Saldo = Saldo + CCTE.Recordset!Impte
    Else
        RSet ElResto = Format(0, "#####,##0.00")
        RSet ElHaber = Format(CCTE.Recordset!Impte, "#####,##0.00")
        Saldo = Saldo - CCTE.Recordset!Impte
    End If
        
    If CCTE.Recordset!FECHA < FecDsd Then
        SaldAnt = SaldAnt + Saldo
        GoTo SaltoTo
    End If
    
    RSet Impu1 = Format(CCTE.Recordset!imput1, "##")
    RSet Impu2 = Format(CCTE.Recordset!imput2, "##")
    
    RSet ElCpbte = Format(CCTE.Recordset!Cpbte, "#####0")
    RSet ElSaldo = Format(Saldo, "###,##0.00")
    
    If ElResto = "        0,00" Then ElResto = "            "
    If ElDebe = "         0,00" Then ElDebe = "            "
    If ElHaber = "         0,00" Then ElHaber = "            "
    'SALDO = Len(ElHaber)
    
    TitDet = " " & Format(CCTE.Recordset!FECHA, "dd/mm/yy") & " " & TipoCpbt & " " & ElCpbte & _
                    " " & Impu1 & " " & Impu2 & _
                    " " & ElResto & ElDebe & ElHaber & ElSaldo
            
            TotPed = TotPed + 1
    
            ImprDet
            If Val(ElResto) > 0 Then
                TotDebeP = TotDebeP + ElResto
            End If
            
SaltoTo:
            
            CCTE.Recordset.MoveNext
            
        Loop
'-----------------------------------------------------------

     TitDet = ""
     ImprDet
     RSet STRImpre = Format(TotDebeP, "$ ###,###,##0.00")
     TitDet = Space(8) & "Sum.Restos :" & STRImpre & Space(19) & "Saldo :" & ElSaldo
     ImprDet
     TitDet = Rayitas
     ImprDet
     TitDet = ""
     ImprDet
     
     TotDebeT = TotDebeT + TotDebeP
     TotDebeP = 0: TotVenc = 0

FinDo1:
            
    CLTE.Recordset.MoveNext

Loop

    TitTot = " Clientes : " & TotItems & "      Comprobantes : " & TotPed
    ImprTot
   
    TotPed = 0: TotItems = 0

' total final

Close #1
End Sub
'        Saldos                      - CC010L -      12
Private Sub SelSaldos()

Label1.Caption = "Saldos de Cta. Cte."
Option2(1).Value = True

Label2.Visible = True
Label6.Visible = False
Combo2.Visible = True

Combo2.SetFocus
Check1.Visible = True

End Sub

Private Sub LisSaldos()
Dim I1%, LosDias As String * 6
Dim Valor1 As String * 14, Valor2 As String * 14
Dim Dias As Integer, ElTipo As String * 4, Estado As String * 5, ElCpbte As String * 6
Dim TotImp As Currency, TotGral As Currency, PClte As String * 5, PNomb As String * 25

Asteriscos = String(80, "*")
Rayitas = String(80, "-")

FG2.Clear flexClearScrollable
FG2.Rows = 1

TotDebeP = 0: TotDebeT = 0: TotVenc = 0: TotVencT = 0

Tit1 = Chr(14) & NomEmpr & Chr(20) & Space(34) & "Fecha " & Format(Date, "dd/mm/yyyy")
Tit2 = Space(10) & Chr(14) & "Saldos de Cta. Cte." & Chr(20)

Tit3 = "CC010L" & Space(60) & "Hoja Nro. "

'      "123456789-123456789-123456789-123456789-"
Tit4 = "*            C l i e n t e            * " & _
       "   Total    *       Observaciones      *"

Select Case OptDos

    Case 0  '   Uno Solo
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE

    Case 1 '   Una Zona
        Tit3 = "CC010L" & Space(26) & "Zona : " & _
                Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & Space(25) & "Hoja Nro. "
        CLTE.RecordSource = "SELECT * FROM Clientes WHERE zona = " & Val(Mid(Combo2.List(Combo2.ListIndex), 1, 2)) & " order by NOMB"

    Case 2 ' Todos
        CLTE.RecordSource = "SELECT * FROM Clientes order by NOMB"
End Select
    
CLTE.Refresh

If CLTE.Recordset.EOF Then Exit Sub

CLTE.Recordset.MoveFirst

    Label8.Visible = True
    Label8.Caption = " Se listarán " & CLTE.Recordset.RecordCount & " Clientes"
    
    'Open "LPT1:" For Output As #1 Len = 226
    Open SaleImpr For Output As #1 Len = 226
    ImprMatri Chr(27) & "@" & Chr(27) & "CH" & Chr(18) & Chr(20), 0
        
        ImprTitulo
    
    TotPed = 0: TotItems = 0

Do Until CLTE.Recordset.EOF

    '----------------------------------------------------
    CCTE.RecordSource = "SELECT * FROM CtasCtes WHERE CLTE = " & CLTE.Recordset!CODIGO
        
        CCTE.Refresh
    'Debug.Print CCTE.RecordSource
    
    If CCTE.Recordset.EOF Then GoTo FinDo1

    '----------------------------------------------------
    
    CCTE.Recordset.MoveFirst
    
    Do Until CCTE.Recordset.EOF
    
    Select Case CCTE.Recordset!TIPO
        Case 0, 1, 3, 7
            TotImp = TotImp + CCTE.Recordset!Impte
        Case Else
            TotImp = TotImp - CCTE.Recordset!Impte
    End Select
    
            CCTE.Recordset.MoveNext
            
        Loop
'-----------------------------------------------------------

    '----------------------------------------------------
    '   Imprime datos
    
    If TotImp < 0.51 And TotImp > -0.51 Then GoTo FinDo1
    
    RSet PClte = CLTE.Recordset!CODIGO
    LSet PNomb = CLTE.Recordset!Nomb
    RSet Valor1 = Format(TotImp, "###,##0.00")
    
    TitDet = PClte & " " & PNomb & " " & Valor1
    ImprDet
               
            If Check1.Value = 1 Then
                FG2.AddItem PClte & vbTab & _
                            PNomb & vbTab & _
                            Valor1, FG2.Rows
            End If
      
'     TitDet = ""
'     ImprDet
            
            TotItems = TotItems + 1
            TotGral = TotGral + TotImp
            
     TotImp = 0

FinDo1:
            
    CLTE.Recordset.MoveNext

Loop
     
     TitDet = Rayitas
     ImprDet
    
    RSet Valor1 = Format(TotGral, "####,###,##0.00")

    TitTot = " Clientes : " & TotItems & "          Deuda Total : " & Valor1 & Chr(18)
    ImprTot
   
    TotPed = 0: TotItems = 0

' total final

Close #1

If Check1.Value = 1 Then

    FG2.TextMatrix(0, 0) = "Código"
    FG2.TextMatrix(0, 1) = "Razón Social"
    FG2.TextMatrix(0, 2) = "Total"

    AExcel 6, 2
End If

End Sub
'        STOCK Valorizado            - ST010L -      13
Private Sub SelStkVal()

End Sub

Private Sub LisStkVal()

End Sub

'---------------------------------------------------------

Private Sub ImprTitulo()
Dim LaHoja As String * 4

lineas = 7

NroHoja = NroHoja + 1
RSet LaHoja = NroHoja

ImprMatri Tit1, 1
ImprMatri Tit2, 1
ImprMatri Tit3 & LaHoja, 1
ImprMatri Asteriscos, 1
ImprMatri Tit4, 1

If Tit5 <> "" Then
    ImprMatri Tit5, 1
    lineas = lineas + 1
End If

ImprMatri Asteriscos, 2

End Sub

Private Sub ImprTit()

If lineas > 65 Then
    ImprMatri "", (72 - lineas)
    ImprTitulo
End If

ImprMatri TitTit, 2

lineas = lineas + 2

End Sub

Private Sub ImprDet()

If lineas > 66 Then
    ImprMatri "", (72 - lineas)
    ImprTitulo
End If

ImprMatri TitDet, 1

lineas = lineas + 1

End Sub

Private Sub ImprTot()

If lineas > 65 Then
    ImprMatri "", (72 - lineas)
    ImprTitulo
End If

ImprMatri "", 1
ImprMatri TitTot, 1
ImprMatri Rayitas, 1

lineas = lineas + 3

End Sub

Private Sub CargaZona()
Dim I1

    ' Tabla de Zonas
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'ZN   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

Do Until RgTABL.EOF
Combo2.AddItem Mid(RgTABL!COD, 1, 3) & "-" & RgTABL!Descri
RgTABL.MoveNext
Loop

RgTABL.Close

End Sub

Private Sub CargaPcia()
Dim I1, I4

    ' Tabla de Zonas
miSQL2 = "SELECT pcia from clientes group by pcia"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

Do Until RgTABL.EOF

If RgTABL!PCIA <> "  " Then
    'Cod.Post., Loc.,Pcia.
    I4 = Asc(RgTABL!PCIA)
    Combo6.AddItem Mid(RgTABL!PCIA, 1, 3) & "-" & Mid(PCIA(I4 - 65), 2, 10)
    End If
    
RgTABL.MoveNext
Loop

RgTABL.Close

End Sub

Private Sub CargaVend()
Dim I1

Combo5.Clear

Combo5.AddItem "000- TODOS "

    ' Tabla de Vendedores
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'VD   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

Do Until RgTABL.EOF
Combo5.AddItem Mid(RgTABL!COD, 1, 3) & "-" & RgTABL!Descri
RgTABL.MoveNext
Loop

RgTABL.Close

End Sub

Private Sub CargaCVta()
Dim I1

    ' Tabla de Condiciones de Venta
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'CV   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

Do Until RgTABL.EOF
Combo3.AddItem Mid(RgTABL!COD, 1, 3) & "-" & RgTABL!Descri

RgTABL.MoveNext
Loop

RgTABL.Close

End Sub

'
'  Genera excell
Private Sub AExcel(Cual As Integer, CanCpos As Integer)

Dim I1 As Integer
Dim PUsuario As String * 6, PNomb As String * 25, PCpbte As String * 13, PCUIT As String * 11
Dim PZona As String * 15, PCateg As String * 15, PSubCat As String * 10
Dim PMedidor As String * 15, PConsumo As String * 7
Dim PConsumo2 As String * 7, PFecha As String * 16
Dim PImpte1 As String * 12, PImpte2 As String * 12
Dim PImpte3 As String * 12, PImpte4 As String * 12
Dim PImpte5 As String * 12, PImpte6 As String * 12, PImpte7 As String * 12
Dim SepMil As String, SepDec As String
SepMil = ",": SepDec = "."

'localidad  titular  cuenta  cod.categoria  des.categ. cod.subcat  des.subcat
'cuit-cuil   medidor   consumo consumo/2  comprobante fecha
'impt.electr.  Res.652/  iva  ley1098   Disp.439/  Subs. Dis.332/

'If FG2.Rows < 2 Then Exit Sub

Dim i As Double, J As Double

Dim IMPTE1         As Double, IMPTE2        As Double, IMPTE3        As Double
Dim IMPTE4         As Double, IMPTE5        As Double, IMPTE6        As Double
Dim IMPTE7         As Double

Dim Obj_Excel   As Object
Dim Obj_Libro   As Object
Dim Obj_Hoja    As Object
Dim Obj_Celda   As Object
Dim NomArch     As String

'-----------------------------

        Set Obj_Excel = CreateObject("Excel.Application")
        Set Obj_Libro = Obj_Excel.Workbooks.Add        ' -- Agregar nuevo libro
'        Set Obj_Libro = Obj_Excel.Workbooks.Open("C:\Coso.xls")
        Set Obj_Hoja = Obj_Excel.ActiveSheet
'        Set Obj_Celda = Excel.range

FG2.AutoSize 0, CanCpos
   
For i = 0 To FG2.Rows - 1                ' -- Recorrer las filas de la Grilla

    For J = 0 To CanCpos                      ' --  Recorrer Las columnas
' -- Asignar el valor a la celda del Excel
        If i = 0 Then
            Obj_Hoja.Cells(i + 2, J + 1) = FG2.TextMatrix(i, J)
        Else
'        If i < FG2.Rows - 1 Then
          Select Case Cual
          
            Case Is < 4
                          Select Case J
                              Case 0
                                  If IsDate(FG2.TextMatrix(i, J)) Then
                                      Obj_Hoja.Cells(i + 2, J + 1) = CDate(FG2.TextMatrix(i, J))
                                  End If
                              Case Is < 6
                                  Obj_Hoja.Cells(i + 2, J + 1) = FG2.TextMatrix(i, J)
                              Case 6, 7, 8, 9, 10, 11, 12
                                  If IsNumeric(FG2.TextMatrix(i, J)) Then
                                      Obj_Hoja.Cells(i + 2, J + 1) = CDec(FG2.TextMatrix(i, J))
                                  End If
                          End Select
            Case 4
                          Select Case J
                              Case 0
                                  Obj_Hoja.Cells(i + 2, J + 1) = FG2.TextMatrix(i, J)
                              Case Is < 4
                                  Obj_Hoja.Cells(i + 2, J + 1) = FG2.TextMatrix(i, J)
                              Case 4
                                      If IsDate(FG2.TextMatrix(i, J)) Then
                                          Obj_Hoja.Cells(i + 2, J + 1) = CDate(FG2.TextMatrix(i, J))
                                      Else
                                          Obj_Hoja.Cells(i + 2, J + 1) = CDec(FG2.TextMatrix(i, J))
                                      End If
                              Case 5
'                                      Obj_Hoja.Cells(i + 2, J + 1) = CDec(FG2.TextMatrix(i, J))
                                      Obj_Hoja.Cells(i + 2, J + 1) = CDec(FG2.TextMatrix(i, J))
                          End Select
            Case 5
                          Select Case J
                              Case 0
                                  Obj_Hoja.Cells(i + 2, J + 1) = FG2.TextMatrix(i, J)
                              Case 1
                                      If IsDate(FG2.TextMatrix(i, J)) Then
                                          Obj_Hoja.Cells(i + 2, J + 1) = CDate(FG2.TextMatrix(i, J))
                                      End If
                              Case Is < 8
                                  Obj_Hoja.Cells(i + 2, J + 1) = FG2.TextMatrix(i, J)
                              Case 8, 9, 10, 11, 12, 13
                                  If IsNumeric(FG2.TextMatrix(i, J)) Then
                                      Obj_Hoja.Cells(i + 2, J + 1) = CDec(FG2.TextMatrix(i, J))
                                  End If
                          End Select
            Case 6
                          Select Case J
                              Case Is < 2
                                  Obj_Hoja.Cells(i + 2, J + 1) = FG2.TextMatrix(i, J)
                              Case 2
                                  If IsNumeric(FG2.TextMatrix(i, J)) Then
                                      Obj_Hoja.Cells(i + 2, J + 1) = CDec(FG2.TextMatrix(i, J))
                                  End If
                          End Select
            Case 7
                          Select Case J
                              Case Is < 8
                                  Obj_Hoja.Cells(i + 2, J + 1) = "'" & FG2.TextMatrix(i, J)
                              Case 8
                                      If IsDate(FG2.TextMatrix(i, J)) Then
                                          Obj_Hoja.Cells(i + 2, J + 1) = CDate(FG2.TextMatrix(i, J))
                                      End If
                              Case 9
                                  If IsNumeric(FG2.TextMatrix(i, J)) Then
                                      Obj_Hoja.Cells(i + 2, J + 1) = CDec(FG2.TextMatrix(i, J))
                                  End If
                               Case 10
                                  Obj_Hoja.Cells(i + 2, J + 1) = FG2.TextMatrix(i, J)
                         End Select
          End Select
        End If
        
    Next
Next
SepMil = ",": SepDec = "."
        
'-----------   Final
'-------------------------------------------------------------
        With Obj_Hoja
            .Rows(1).Font.Size = 15        ' -- Opcional : colocar en negrita y de color AZUL los enbezados en la hoja
            .Rows(1).Font.Bold = True        ' -- Opcional : colocar en negrita y de color AZUL los enbezados en la hoja
            .Rows(1).Font.Color = vbBlue
            .Rows(2).Font.Bold = True        ' -- Opcional : colocar en negrita y de color AZUL los enbezados en la hoja
            .Rows(2).Font.Color = vbBlue
            .Columns("A:Z").AutoFit          ' -- Autoajustar las cabeceras
            .Rows(FG2.Rows + 2).Font.Bold = True      '   Pone en negrita los totales
'            .Range("e1:i" & FG2.Rows + 1).Select        '  selecciona los importes
'            .NumberFormat = "##########" + "0" + SepDec + "00"
'            .Columns ("E:I")
        End With
        
'   Autosuma todos los importes
Dim NLet As String
NLet = Asc("A")
Dim w1 As Integer, w2 As Integer

Select Case Cual
    Case 1
        w1 = 8: w2 = 12
        Obj_Hoja.Cells(1, 2) = "SUBDIARIO DE IVA"
    Case 2
        w1 = 7: w2 = 12
        Obj_Hoja.Cells(1, 2) = "SUBDIARIO DE COBRANZA"
    Case 3
        w1 = 7: w2 = 8
        Obj_Hoja.Cells(1, 2) = "PERCEPCIONES ARBA"
    Case 4
        w1 = 6: w2 = 6
        Obj_Hoja.Cells(1, 2) = "COMISIONES POR COBRANZAS"
    Case 5
        w1 = 9: w2 = 13
        Obj_Hoja.Cells(1, 2) = "Ventas por Vendedor - Comisiones"
    Case 6
        w1 = 3: w2 = 3
        Obj_Hoja.Cells(1, 2) = "Saldos de Cuenta Corriente"
    Case 7
        w1 = 9: w2 = 9
        Obj_Hoja.Cells(1, 2) = "Planilla de Cobranzas"
End Select


'Obj_Excel.Rows(FG2.Rows + 1).Select
Select Case Cual
    Case 4
        If FG2.Rows > 1 Then
            For i = w1 To w2
                NLet = Chr(64 + i)
                Obj_Excel.Cells(FG2.Rows + 2, i).Formula = "=SUM(" & NLet & "3:" & NLet & FG2.Rows + 1 & ")"
            Next i
        End If
        Obj_Excel.Range("f3:f" & FG2.Rows + 2).Select            ' Reformatea las celdas de importes
    Case 6
        If FG2.Rows > 1 Then
            For i = w1 To w2
                NLet = Chr(64 + i)
                Obj_Excel.Cells(FG2.Rows + 2, i).Formula = "=SUM(" & NLet & "3:" & NLet & FG2.Rows + 1 & ")"
            Next i
        End If
        Obj_Excel.Range("c3:f" & FG2.Rows + 2).Select            ' Reformatea las celdas de importes
    Case 7
        If FG2.Rows > 1 Then
            For i = w1 To w2
                NLet = Chr(64 + i)
                Obj_Excel.Cells(FG2.Rows + 2, i).Formula = "=SUM(" & NLet & "3:" & NLet & FG2.Rows + 1 & ")"
            Next i
        End If
        Obj_Excel.Range("c3:f" & FG2.Rows + 2).Select            ' Reformatea las celdas de importes
    Case Else
        If FG2.Rows > 1 Then
            For i = w1 To w2
                NLet = Chr(64 + i)
                Obj_Excel.Cells(FG2.Rows + 2, i).Formula = "=SUM(" & NLet & "3:" & NLet & FG2.Rows + 1 & ")"
            Next i
        End If
        Obj_Excel.Range("e3:l" & FG2.Rows + 2).Select           ' Reformatea las celdas de importes
End Select

    Obj_Excel.Selection.NumberFormat = "###,###,##0.00"
    Obj_Excel.Rows(FG2.Rows + 2).Select                     '  Selecciona la linea de totales
'             Obj_Excel.SendKeys ("%=~")
       
'Obj_Excel.Columns(1).Select
'Obj_Excel.selection.NumberFormat = "dd/MM/yyyy"

'            Obj_Excel.selection.NumberFormat = "#" + sepMil + "0" + sepDec + "00"

' Obj_Excel.Visible = True               ' -- Hacer excel visible
Select Case Cual
    Case 1
        NomArch = "C:\Subdiario de Ventas - " & Format(FECHA.Value, "mmm") & "-" & FECHA.Year
    Case 2
        NomArch = "C:\Subdiario de Cobranza - " & Format(FECHA.Value, "mmm") & "-" & FECHA.Year
    Case 3
        NomArch = "C:\Percepciones ARBA - " & Format(FECHA.Value, "mmm") & "-" & FECHA.Year
    Case 4
        NomArch = "C:\Com.x Cobr.-" & Format(FECHA.Value, "mmm") & "-" & FECHA.Year & "-" & Combo5.Text
    Case 5
        NomArch = "C:\Com. x Vtas.- " & Format(FECHA.Value, "mmm") & "-" & FECHA.Year & "-" & Combo5.Text
    Case 6
        NomArch = "C:\Saldos Cta.Cte.- " & Format(Date, "mmm") & "-" & Format(Date, "yyyy") & "-" & Combo5.Text
    Case 7
        NomArch = "C:\Planilla de Cobranzas - " & Format(Date, "mmm") & "-" & Format(Date, "yyyy") & "-" & Combo5.Text
End Select

Obj_Libro.SaveAs filename:=NomArch & ".xls", FileFormat:=1, CreateBackup:=False
Obj_Libro.Close SaveChanges:=False

    Obj_Excel.Quit
   
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing


End Sub


