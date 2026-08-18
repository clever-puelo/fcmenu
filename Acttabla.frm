VERSION 5.00
Begin VB.Form ActTabla 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Cambios"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Default         =   -1  'True
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
      Left            =   3960
      TabIndex        =   41
      Top             =   4470
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
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
      Left            =   2670
      TabIndex        =   40
      Top             =   4470
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      Caption         =   "  Descripción  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   16
      Top             =   960
      Width           =   4935
      Begin VB.CheckBox Check2 
         Caption         =   "Precio en Dólares"
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
         Left            =   270
         TabIndex        =   60
         ToolTipText     =   "Para el Facturador - Usa esta descripción y no la del Artículo"
         Top             =   600
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Usar para Facturar"
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
         Left            =   2580
         TabIndex        =   59
         ToolTipText     =   "Para el Facturador - Usa esta descripción y no la del Artículo"
         Top             =   600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox Text2 
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
         Left            =   270
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Descripción del Elemento"
         Top             =   300
         Width           =   4485
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Código  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   240
      TabIndex        =   15
      Top             =   210
      Width           =   4935
      Begin VB.TextBox CODNUM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   300
         MaxLength       =   5
         TabIndex        =   0
         ToolTipText     =   "Código del Elemento"
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame UMED 
      Caption         =   "  Valores  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   28
      Top             =   1695
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00F9FADC&
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
         ItemData        =   "Acttabla.frx":0000
         Left            =   1800
         List            =   "Acttabla.frx":0013
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Posición de Carga en el Facturador"
         Top             =   600
         Width           =   1830
      End
      Begin VB.TextBox Text30 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   1800
         LinkTimeout     =   5
         MaxLength       =   1
         TabIndex        =   45
         ToolTipText     =   "Valor Numérico"
         Top             =   600
         Width           =   315
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   3510
         LinkTimeout     =   5
         MaxLength       =   1
         TabIndex        =   4
         ToolTipText     =   "Cant. de Posiciones decimales en el ingreso"
         Top             =   255
         Width           =   315
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   1800
         LinkTimeout     =   5
         MaxLength       =   1
         TabIndex        =   3
         ToolTipText     =   "Cant. de Posiciones enteras en el ingreso"
         Top             =   255
         Width           =   315
      End
      Begin VB.Label Label20 
         Caption         =   "Posición :"
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
         Left            =   1005
         TabIndex        =   46
         ToolTipText     =   "Columna dentro del Facturador"
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label7 
         Caption         =   "Enteros :"
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
         Left            =   1005
         TabIndex        =   30
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label5 
         Caption         =   "Decimales :"
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
         Left            =   2475
         TabIndex        =   29
         Top             =   300
         Width           =   1050
      End
   End
   Begin VB.Frame VARIOS 
      Caption         =   "Valores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1410
      Left            =   240
      TabIndex        =   35
      Top             =   1695
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text14 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   5
         EndProperty
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
         Left            =   3540
         LinkTimeout     =   5
         MaxLength       =   6
         TabIndex        =   9
         ToolTipText     =   "Valor Numérico"
         Top             =   705
         Width           =   960
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   5
         EndProperty
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
         Left            =   3540
         LinkTimeout     =   5
         MaxLength       =   6
         TabIndex        =   8
         ToolTipText     =   "Valor Numérico"
         Top             =   390
         Width           =   960
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   1230
         LinkTimeout     =   5
         MaxLength       =   10
         TabIndex        =   6
         ToolTipText     =   "Valor Numérico"
         Top             =   390
         Width           =   960
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   1230
         LinkTimeout     =   5
         MaxLength       =   10
         TabIndex        =   7
         ToolTipText     =   "Valor Numérico"
         Top             =   705
         Width           =   960
      End
      Begin VB.Label Label19 
         Caption         =   "Porcent. 2  :"
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
         Left            =   2430
         TabIndex        =   39
         Top             =   735
         Width           =   1050
      End
      Begin VB.Label Label18 
         Caption         =   "Porcent. 1  :"
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
         Left            =   2445
         TabIndex        =   38
         Top             =   435
         Width           =   1050
      End
      Begin VB.Label Label17 
         Caption         =   "Numero 2  :"
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
         Left            =   135
         TabIndex        =   37
         Top             =   735
         Width           =   1050
      End
      Begin VB.Label Label16 
         Caption         =   "Numero 1  :"
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
         Left            =   135
         TabIndex        =   36
         Top             =   435
         Width           =   1050
      End
   End
   Begin VB.Frame PCIA 
      Caption         =   "  Valores  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   33
      Top             =   1695
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text10 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
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
         Left            =   2760
         LinkTimeout     =   5
         MaxLength       =   1
         TabIndex        =   13
         ToolTipText     =   "Valor Numérico"
         Top             =   435
         Width           =   315
      End
      Begin VB.Label Label10 
         Caption         =   "Letra  :"
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
         Left            =   2085
         TabIndex        =   34
         Top             =   465
         Width           =   885
      End
   End
   Begin VB.Frame VEND 
      Caption         =   "  Valores  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1425
      Left            =   240
      TabIndex        =   17
      Top             =   1695
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   5
         EndProperty
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
         Left            =   3630
         LinkTimeout     =   5
         MaxLength       =   15
         TabIndex        =   12
         ToolTipText     =   "Valor Numérico"
         Top             =   1005
         Width           =   945
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   5
         EndProperty
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
         Left            =   1050
         LinkTimeout     =   5
         MaxLength       =   15
         TabIndex        =   11
         ToolTipText     =   "Valor Numérico"
         Top             =   1005
         Width           =   945
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   2550
         LinkTimeout     =   5
         MaxLength       =   3
         TabIndex        =   10
         ToolTipText     =   "Valor Numérico"
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label4 
         Caption         =   "x Cobranza :"
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
         Left            =   2505
         TabIndex        =   27
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Comisiones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1545
         TabIndex        =   26
         Top             =   630
         Width           =   1500
      End
      Begin VB.Label Label9 
         Caption         =   "x Venta :"
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
         Left            =   210
         TabIndex        =   22
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "Zona      :"
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
         Left            =   1725
         TabIndex        =   21
         Top             =   255
         Width           =   975
      End
   End
   Begin VB.Frame CVTA 
      Caption         =   "  Valores  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   31
      Top             =   1695
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text9 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
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
         Left            =   2400
         LinkTimeout     =   5
         MaxLength       =   3
         TabIndex        =   14
         ToolTipText     =   "Valor Numérico"
         Top             =   465
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Cant. Días  :"
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
         Left            =   1320
         TabIndex        =   32
         Top             =   480
         Width           =   1065
      End
   End
   Begin VB.Frame SECC1 
      Caption         =   "  Datos  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2355
      Left            =   240
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ComboBox LIST1 
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
         Left            =   1245
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Unid. de Facturación :"
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
         Left            =   780
         TabIndex        =   58
         Top             =   2010
         Width           =   2130
      End
      Begin VB.Label Text27 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2940
         TabIndex        =   57
         ToolTipText     =   "Le indica a los Facturadores la Unidad por la que será vendida el Artículo"
         Top             =   1950
         Width           =   1005
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   2520
         X2              =   2520
         Y1              =   300
         Y2              =   1860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   2490
         X2              =   2490
         Y1              =   300
         Y2              =   1860
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "Cálculo de Precio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2730
         TabIndex        =   56
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Especificación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   270
         TabIndex        =   55
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Text24 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3690
         TabIndex        =   54
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "1 -"
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
         Left            =   2700
         TabIndex        =   53
         Top             =   930
         Width           =   420
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "2 -"
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
         Left            =   2670
         TabIndex        =   52
         Top             =   1290
         Width           =   450
      End
      Begin VB.Label Text25 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3690
         TabIndex        =   51
         Top             =   1215
         Width           =   1005
      End
      Begin VB.Label Text26 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3690
         TabIndex        =   50
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "Unid. Medida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3510
         TabIndex        =   49
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Posición"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2670
         TabIndex        =   48
         Top             =   510
         Width           =   765
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "3 -"
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
         Left            =   2670
         TabIndex        =   47
         Top             =   1590
         Width           =   450
      End
      Begin VB.Label Text23 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1245
         TabIndex        =   44
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label Text22 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1245
         TabIndex        =   43
         Top             =   1215
         Width           =   1005
      End
      Begin VB.Label Text21 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1260
         TabIndex        =   42
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label Label13 
         Caption         =   "3 - "
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
         Left            =   405
         TabIndex        =   25
         Top             =   1575
         Width           =   405
      End
      Begin VB.Label Label12 
         Caption         =   "2 - "
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
         Left            =   405
         TabIndex        =   24
         Top             =   1275
         Width           =   405
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Posición"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   510
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Unid. Medida"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1170
         TabIndex        =   20
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label Label11 
         Caption         =   "1 - "
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
         Left            =   405
         TabIndex        =   19
         Top             =   945
         Width           =   405
      End
   End
End
Attribute VB_Name = "ActTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UnaVez As Boolean, ALF1 As String, PROGRES As String, CanMin As Single
Dim NumSD1 As Integer, NumSD2 As Integer, NumSD3 As Integer
Dim NumCD1 As Currency, NumCD2 As Currency, NumCD3 As Currency

'------------------------------------------------------------
'   1-Secciones             5-Unidades de medida
'   2-Vendedores            6-Cond. de Venta
'   3-Zonas                 7-Motivos de NC y ND
'   4-Provincias            8-Datos Varios
'------------------------------------------------------------

Private Sub Form_Load()

Me.Move 5700, 0, 5595, 3405
Dim I1 As Single

       VEND.Visible = False
       SECC1.Visible = False
       PCIA.Visible = False
       UMED.Visible = False
       CVTA.Visible = False
       LIST1.Visible = False
       VARIOS.Visible = False
       Check1.Visible = False
        
        CODNUM.Visible = True
       
Command2.Enabled = False

If LaOpcion = 1 Then
        miSQL2 = "SELECT * FROM FCTabla1 WHERE CTAB = 'UM   ' ORDER BY COD"
        RgTABL.Open miSQL2, dbTABL, adOpenDynamic, adLockReadOnly
        
        LIST1.AddItem "     "
        Do Until RgTABL.EOF
        LIST1.AddItem RgTABL!COD
        RgTABL.MoveNext
        Loop

RgTABL.Close

End If


Command2.Caption = "Grabar"
    
    
    miSQL2 = "SELECT * FROM FCTabla1 WHERE CTAB = '" & CodTabla & "' AND COD = '" _
                                    & CodItem & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenDynamic, adLockReadOnly

        
        If Not RgTABL.EOF Then
                Text2.Text = RgTABL!Descri
                Select Case LaOpcion
                        Case 1
                          Text21.Caption = RgTABL!ALF1
                          Text22.Caption = RgTABL!ALF2
                          Text23.Caption = RgTABL!ALF3
                          Text24.Caption = RgTABL!ALF4
                          Text27.Caption = RgTABL!ALF7
                          If RgTABL!ALF5 <> "" Then Text25.Caption = RgTABL!ALF5
                          If RgTABL!ALF6 <> "" Then Text26.Caption = RgTABL!ALF6
                          Text21.Visible = True
                          Text22.Visible = True
                          Text23.Visible = True
                          Text25.Visible = True
                          Text26.Visible = True
                          Text27.Visible = True
                          If RgTABL!mca1 Then Check1.Value = 1
                          Check1.Visible = True
                          If RgTABL!mca2 Then Check2.Value = 1
                          Check2.Visible = True
                        Case 2
                          Text3.Text = Format(RgTABL!NumSD1, "##0")
                          Text4.Text = Format(RgTABL!NumCD1, "##0.00")
                          Text6.Text = Format(RgTABL!NumCD2, "##0.00")
                        Case 3
                        Case 4
                          Text10.Text = RgTABL!ALF1
                        Case 5
                          Text7.Text = Format(RgTABL!NumSD1, "##0")
                          Text8.Text = Format(RgTABL!NumSD2, "##0")
                          BuscaUno RgTABL!NumSD3

                        Case 6
                          Text9.Text = Format(RgTABL!NumSD3, "##0")
                        Case 7
                End Select
                
        End If
        
Select Case LaOpcion

    Case 1
        CODNUM.Text = CodItem
        Me.Move 5650, 0, 5595, 5300
        Command1.top = 4470
        Command2.top = 4470
        SECC1.Visible = True
            
    Case 2
        CODNUM.Text = CodItem
        Me.Move 5650, 1200, 5595, 4000
        Command1.top = 3200
        Command2.top = 3200
        VEND.Visible = True
            
    Case 3
        CODNUM.Text = CodItem
        Me.Move 5650, 1200, 5595, 3550
        Command1.top = 2750
        Command2.top = 2750
'        PCIA.Visible = True
            
    Case 4
        CODNUM.Text = CodItem
        Me.Move 5650, 1200, 5595, 3550
        Command1.top = 2750
        Command2.top = 2750
        PCIA.Visible = True

    Case 5
        CODNUM.Text = CodItem
        Me.Move 5650, 1200, 5595, 3550
        Command1.top = 2750
        Command2.top = 2750
        UMED.Visible = True
        
    Case 6
        CODNUM.Text = CodItem
        Me.Move 5650, 1200, 5595, 3550
        Command1.top = 2750
        Command2.top = 2750
        CVTA.Visible = True
        
    Case 7
        CODNUM.Text = CodItem
        Me.Move 5650, 1200, 5595, 3200
        Command1.top = 2400
        Command2.top = 2400
        
    Case 8
        CODNUM.Text = CodItem
        Me.Move 5650, 1200, 5595, 4000
        Command1.top = 3200
        Command2.top = 3200
        VARIOS.Visible = True
    
End Select


Select Case TIPOMov

Case 1
    
    ActTabla.Caption = "Alta de un Elemento"
    CODNUM.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Combo1.ListIndex = 0
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    Text14.Text = ""
    If LaOpcion = 1 Then Check1.Visible = True

Case 2, 3

    
    CODNUM.Enabled = False
    Label10.Enabled = False
    
    'Text2.Enabled = False
    'Text3.Enabled = False
    'Text4.Enabled = False
    'Text6.Enabled = False
    'Text7.Enabled = False
    'Text8.Enabled = False
    'Text9.Enabled = False
    'Text10.Enabled = False
    'Text11.Enabled = False
    'Text12.Enabled = False
    'Text13.Enabled = False
    'Text14.Enabled = False
    
    Command2.Enabled = True
    
    If RgTABL.EOF Then Exit Sub
    
    Descrip = RgTABL!Descri
    Text2.Text = RgTABL!Descri
        
    NumSD1 = RgTABL!NumSD1
    NumSD2 = RgTABL!NumSD2
    NumSD3 = RgTABL!NumSD3
    NumCD1 = RgTABL!NumCD1
    NumCD2 = RgTABL!NumCD2
    NumCD3 = RgTABL!NumCD3

        Select Case LaOpcion
            Case 1
            
            Case 2
                Text3.Text = Format(RgTABL!NumSD1, "##0")
                Text4.Text = Format(RgTABL!NumCD1, "##0.00")
                Text6.Text = Format(RgTABL!NumCD2, "##0.00")
            
            Case 3
                Text10.Text = RgTABL!ALF1
            
            Case 4
                Text10.Text = RgTABL!ALF1
            
            Case 5
                Text10.Text = RgTABL!ALF1
             
            Case 6
                Text7.Text = Format(RgTABL!NumSD1, "#0")
                Text8.Text = Format(RgTABL!NumSD2, "#0")
            
            Case 7
                Text9.Text = Format(RgTABL!NumSD3, "##0")
           
            Case 8
                Text11.Text = RgTABL!NumSD1
                Text12.Text = RgTABL!NumSD2
                Text13.Text = RgTABL!NumCD1
                Text14.Text = RgTABL!NumCD2
        
        End Select

    If RgTABL.EOF = True Then
        RgTABL.Close
        Exit Sub
    End If

End Select

Select Case TIPOMov
    
    Case 2
    ActTabla.Caption = "Baja de un Elemento"
    Command2.Caption = "Eliminar"
    
        
    Case 3
    ActTabla.Caption = "Cambios en un Elemento"
       
End Select

       RgTABL.Close

End Sub

Private Sub Option1_Click(Index As Integer)

    Label10.Enabled = True: Text6.Enabled = True

End Sub

Private Sub Form_Deactivate()

Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF2 Then
    BusClte.Show
    Unload Me
        End If


'If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 39 Then SendKeys "{TAB}"
'If KeyCode = 38 Or KeyCode = 37 Then SendKeys "+{TAB}"
If KeyCode = 27 Then End

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Command1_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Command2_Click()

Grabacion

Unload Me

ABMTablas.CargaGrilla

End Sub

Private Sub CODNUM_LostFocus()
    
    
    miSQL2 = "SELECT * FROM FCTabla1 WHERE CTAB = '" & CodTabla & "' AND COD = '" _
                                    & CODNUM & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenDynamic, adLockReadOnly

If (RgTABL.EOF And (TIPOMov = 2 Or TIPOMov = 3)) Or (Not RgTABL.EOF And TIPOMov = 1) Then
    MsgBox "Error en el Código", vbCritical + vbOKOnly, "Error"
    CODNUM.SetFocus
End If

 RgTABL.Close

End Sub


Private Sub LIST1_Click()

If UnaVez Then
'   UnaVez = False
   Exit Sub
   End If
   
Select Case LIST1.top + LIST1.left
  
  Case Text21.top + Text21.left
    Text21.Caption = LIST1.Text
    
  Case Text22.top + Text22.left
    Text22.Caption = LIST1.Text
  
  Case Text23.top + Text23.left
    Text23.Caption = LIST1.Text
  
  Case Text24.top + Text24.left
    If LIST1.left = Text24.left Then
        Text24.Caption = LIST1.Text
    End If
  
  Case Text25.top + Text25.left
        Text25.Caption = LIST1.Text
  
  Case Text26.top + Text26.left
        Text26.Caption = LIST1.Text
  
  Case Text27.top + Text27.left
        Text27.Caption = LIST1.Text
    
End Select


End Sub


Private Sub Text10_Change()

If TIPOMov = 3 Then
        If Text10.Text <> ALF1 Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Private Sub Text11_Change()

If TIPOMov = 3 Then
        If Text11.Text <> NumSD1 Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Private Sub Text12_Change()

If TIPOMov = 3 Then
        If Text12.Text <> NumCD1 Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Private Sub Text2_Change()

If TIPOMov = 3 Then
        If Text2.Text <> Descrip Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Private Sub Text21_Click()

LIST1.Visible = True

BuscaUM Text21.Caption
LIST1.top = Text21.top
LIST1.left = Text21.left

End Sub

Private Sub Text22_Click()

LIST1.Visible = True

BuscaUM Text22.Caption
LIST1.top = Text22.top
LIST1.left = Text22.left

End Sub

Private Sub Text23_Click()

LIST1.Visible = True

BuscaUM Text23.Caption
LIST1.top = Text23.top
LIST1.left = Text23.left

End Sub

Private Sub Text24_Click()

LIST1.Visible = True

BuscaUM Text24.Caption
LIST1.top = Text24.top
LIST1.left = Text24.left

End Sub

Private Sub Text25_Click()

LIST1.Visible = True

BuscaUM Text25.Caption
LIST1.top = Text25.top
LIST1.left = Text25.left

End Sub

Private Sub Text26_Click()

LIST1.Visible = True

BuscaUM Text26.Caption
LIST1.top = Text26.top
LIST1.left = Text26.left

End Sub

Private Sub Text27_Click()

LIST1.Visible = True

BuscaUM Text27.Caption
LIST1.top = Text27.top
LIST1.left = Text27.left

End Sub

Private Sub Text3_Change()

If TIPOMov = 3 Then
        If Text3.Text <> NumSD1 Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Private Sub text4_Change()

If TIPOMov = 3 Then
        If Text4.Text <> NumCD1 Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Private Sub Text6_Change()

If TIPOMov = 3 Then
        If Text6.Text <> NumCD2 Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Private Sub Grabacion()
Dim Mensaje As String

Select Case TIPOMov
        Case 1
                Mensaje = "Desea Grabar ?"
        Case 2
                Mensaje = "Desea Eliminar ?"
        Case 3
                Mensaje = "Desea Cambiar ? "
End Select

Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "Grabación en Tablas")

If Respuesta = vbNo Then Exit Sub
If Respuesta = vbCancel Then
         Exit Sub
         Load Me
         End If
        

    miSQL2 = "SELECT * FROM FCTabla1 WHERE CTAB = '" & CodTabla & "' AND COD = '" _
                                    & CODNUM & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenDynamic, adLockPessimistic

If TIPOMov = 2 Then
        RgTABL.Delete
        RgTABL.Close
        Exit Sub
        End If

If TIPOMov = 1 Then
    RgTABL.AddNew
        If LaOpcion = 1 Or LaOpcion = 5 Then
            RgTABL!COD = CODNUM.Text
            Else
            RgTABL!COD = Val(CODNUM.Text)
            End If
        RgTABL!CTAB = CodTabla
        End If

    RgTABL!Descri = Text2.Text
    RgTABL!USUARIO = Mid$(Red_Usuario, 1, 6)
    RgTABL!FECACT = Date
    
    RgTABL!ALF1 = ""
    RgTABL!ALF2 = ""
    RgTABL!ALF3 = ""
    RgTABL!ALF4 = ""
    RgTABL!ALF5 = ""
    RgTABL!ALF6 = ""
    RgTABL!ALF7 = ""
    RgTABL!ALF8 = ""
    RgTABL!mca1 = False
    RgTABL!mca2 = False
    RgTABL!MCA3 = False
    RgTABL!FECHA1 = "01/01/90"
    RgTABL!FECHA2 = "01/01/90"
    RgTABL!NumCD1 = 0
    RgTABL!NumCD2 = 0
    RgTABL!NumCD3 = 0
    RgTABL!NumSD1 = 0
    RgTABL!NumSD2 = 0
    RgTABL!NumSD3 = 0


Select Case LaOpcion
   Case 1
    RgTABL!ALF1 = Mid(Text21.Caption, 1, 5)
    RgTABL!ALF2 = Mid(Text22.Caption, 1, 5)
    RgTABL!ALF3 = Mid(Text23.Caption, 1, 5)
    RgTABL!ALF4 = Mid(Text24.Caption, 1, 5)
    RgTABL!ALF5 = Mid(Text25.Caption, 1, 5)
    RgTABL!ALF6 = Mid(Text26.Caption, 1, 5)
    RgTABL!ALF7 = Mid(Text27.Caption, 1, 5)
    RgTABL!mca1 = Check1.Value
    RgTABL!mca2 = Check2.Value
   
   Case 2
    RgTABL!NumSD1 = Val(Replace(Text3.Text, ",", "."))
    RgTABL!NumCD1 = Val(Replace(Text4.Text, ",", "."))
    RgTABL!NumCD2 = Val(Replace(Text6.Text, ",", "."))
   
   Case 3
   
   Case 4
    RgTABL!ALF1 = Text10.Text
   
   Case 5
    RgTABL!NumSD1 = Val(Text7.Text)
    RgTABL!NumSD2 = Val(Text8.Text)
    RgTABL!NumSD3 = Mid(Combo1.Text, 1, 1)
    
   Case 6
    RgTABL!NumSD1 = Val(Text7.Text)
    RgTABL!NumSD2 = Val(Text8.Text)
    RgTABL!NumSD3 = Val(Text9.Text)
   
   Case 7
    RgTABL!NumSD1 = Val(Text7.Text)
    RgTABL!NumSD2 = Val(Text8.Text)
    RgTABL!NumSD3 = Val(Text9.Text)
   
   Case 8
    RgTABL!NumSD1 = Val(Text7.Text)
    RgTABL!NumSD2 = Val(Text8.Text)
    RgTABL!NumSD3 = Val(Text9.Text)

End Select

RgTABL.Update
RgTABL.Close

End Sub


Private Sub text1_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'Private Sub Text2_KeyPress(KeyAscii As Integer)

'If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

'If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
'    KeyAscii = KeyAscii - 32
'    End If

'End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)


If KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub text4_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub CODNUM_GotFocus()

CODNUM.SelStart = 0
CODNUM.SelLength = Len(CODNUM.Text)

End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text3_GotFocus()

Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub text4_GotFocus()

Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub Text7_GotFocus()

Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)

End Sub

Private Sub Text8_GotFocus()

Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)

End Sub

Private Sub Text9_GotFocus()

Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)

End Sub


Private Sub Text10_GotFocus()

Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)

End Sub

Private Sub Text11_GotFocus()

Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)

End Sub

Private Sub Text7_Change()

If TIPOMov = 3 Then
        If Text7.Text <> NumSD1 Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Private Sub Text8_Change()

If TIPOMov = 3 Then
        If Text8.Text <> NumSD2 Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Private Sub Text9_Change()

If TIPOMov = 3 Then
        If Text9.Text <> NumSD3 Then
                Command2.Enabled = True
            Else
                Command2.Enabled = False
        End If
    Else
        Command2.Enabled = True
End If

End Sub

Function BuscaUM(UMED As String)
Dim xo

UnaVez = True
If UMED = "" Then UMED = "     "

For xo = 0 To LIST1.ListCount - 1
If UMED = Mid(LIST1.List(xo), 1, 5) Then
    LIST1.ListIndex = xo
    UnaVez = False
    Exit Function
End If

Next

End Function

Function BuscaUno(UNO As String)
Dim xo

UnaVez = True
If UNO = "" Then UNO = "0"

For xo = 0 To Combo1.ListCount - 1
If UNO = Mid(Combo1.List(xo), 1, 1) Then
    Combo1.ListIndex = xo
    UnaVez = False
    Exit Function
End If

Next

End Function

