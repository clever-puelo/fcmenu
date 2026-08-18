VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TotFact 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  Totales Diarios"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11610
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
   ScaleHeight     =   6480
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Total Mensual"
      Height          =   330
      Left            =   9630
      Style           =   1  'Graphical
      TabIndex        =   136
      Top             =   480
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   330
      Left            =   9840
      TabIndex        =   128
      Top             =   4470
      Width           =   1185
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      ItemData        =   "TotFact.frx":0000
      Left            =   5310
      List            =   "TotFact.frx":0019
      TabIndex        =   127
      Text            =   "50"
      Top             =   465
      Width           =   780
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   7380
      ScaleHeight     =   3465
      ScaleWidth      =   3735
      TabIndex        =   105
      Top             =   840
      Visible         =   0   'False
      Width           =   3795
      Begin VB.Frame Frame15 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1080
         Left            =   150
         TabIndex        =   118
         Top             =   1890
         Width           =   3345
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   1560
            TabIndex        =   124
            Top             =   720
            Width           =   1650
         End
         Begin VB.Label Label38 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Unidad Prom.:"
            Height          =   240
            Left            =   150
            TabIndex        =   123
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   1545
            TabIndex        =   122
            Top             =   480
            Width           =   1650
         End
         Begin VB.Label Label41 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Prom. Diario  :"
            Height          =   240
            Left            =   150
            TabIndex        =   121
            Top             =   480
            Width           =   1440
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   1560
            TabIndex        =   120
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label43 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Proyección    :"
            Height          =   240
            Left            =   150
            TabIndex        =   119
            Top             =   240
            Width           =   1440
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         Height          =   270
         Left            =   2280
         MaskColor       =   &H00C0E0FF&
         TabIndex        =   117
         Top             =   3060
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "  Totales a la Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1230
         Left            =   150
         TabIndex        =   108
         Top             =   675
         Width           =   3345
         Begin VB.Label Label28 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Precio  Costo :"
            Height          =   240
            Left            =   150
            TabIndex        =   116
            Top             =   240
            Width           =   1440
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   1560
            TabIndex        =   115
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label30 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Precio Venta :"
            Height          =   240
            Left            =   150
            TabIndex        =   114
            Top             =   480
            Width           =   1440
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   1545
            TabIndex        =   113
            Top             =   480
            Width           =   1650
         End
         Begin VB.Label Label32 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Unidades       :"
            Height          =   240
            Left            =   150
            TabIndex        =   112
            Top             =   960
            Width           =   1440
         End
         Begin VB.Label Label33 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Venta Real     :"
            Height          =   240
            Left            =   150
            TabIndex        =   111
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   1560
            TabIndex        =   110
            Top             =   960
            Width           =   1650
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   240
            Left            =   1560
            TabIndex        =   109
            Top             =   720
            Width           =   1650
         End
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2265
         MaxLength       =   2
         TabIndex        =   107
         Top             =   105
         Width           =   780
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   2265
         MaxLength       =   2
         TabIndex        =   106
         Text            =   "1"
         Top             =   375
         Width           =   780
      End
      Begin VB.Label Label37 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Días Hábiles       :"
         Height          =   240
         Left            =   510
         TabIndex        =   126
         Top             =   105
         Width           =   1710
      End
      Begin VB.Label Label39 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Días Trabajados :"
         Height          =   240
         Left            =   510
         TabIndex        =   125
         Top             =   405
         Width           =   1710
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H003D813A&
      Height          =   3525
      Left            =   30
      ScaleHeight     =   3465
      ScaleWidth      =   7305
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   7365
      Begin VB.Frame Frame7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "  IMPORTES  "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         Left            =   120
         TabIndex        =   5
         Top             =   105
         Visible         =   0   'False
         Width           =   7125
         Begin VB.Frame Frame2 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Cta. Cte.  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   975
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   3420
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   19
               Top             =   705
               Width           =   1650
            End
            Begin VB.Label Label3 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""C"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   18
               Top             =   705
               Width           =   1440
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   17
               Top             =   480
               Width           =   1635
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   16
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label Label1 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""A"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   15
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label2 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""B"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   14
               Top             =   480
               Width           =   1440
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Mostrador   "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   975
            Left            =   3585
            TabIndex        =   6
            Top             =   240
            Width           =   3420
            Begin VB.Label Label8 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""B"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   12
               Top             =   480
               Width           =   1440
            End
            Begin VB.Label Label7 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""A"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   11
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   10
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   9
               Top             =   480
               Width           =   1635
            End
            Begin VB.Label Label9 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""C"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   8
               Top             =   705
               Width           =   1440
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   7
               Top             =   705
               Width           =   1650
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Exportación  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   600
            Left            =   1920
            TabIndex        =   20
            Top             =   1170
            Width           =   3420
            Begin VB.Label Label13 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""E"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   22
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   21
               Top             =   240
               Width           =   1650
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Unidades Promedio x Día  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   1020
            Left            =   120
            TabIndex        =   30
            Top             =   1725
            Width           =   3420
            Begin VB.Label Label16 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Por Mostrador :"
               Height          =   240
               Left            =   135
               TabIndex        =   36
               Top             =   525
               Width           =   1545
            End
            Begin VB.Label Label15 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Por Cta. Cte.    :"
               Height          =   240
               Left            =   135
               TabIndex        =   35
               Top             =   285
               Width           =   1545
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1665
               TabIndex        =   34
               Top             =   285
               Width           =   1425
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1665
               TabIndex        =   33
               Top             =   525
               Width           =   1425
            End
            Begin VB.Label Label17 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Por Tipo ""C""    :"
               Height          =   240
               Left            =   135
               TabIndex        =   32
               Top             =   750
               Width           =   1545
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1665
               TabIndex        =   31
               Top             =   750
               Width           =   1425
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Totales  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   1020
            Left            =   3585
            TabIndex        =   23
            Top             =   1725
            Width           =   3420
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1680
               TabIndex        =   29
               Top             =   735
               Width           =   1425
            End
            Begin VB.Label Label23 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Venta Real      :"
               Height          =   240
               Left            =   150
               TabIndex        =   28
               Top             =   735
               Width           =   1545
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1680
               TabIndex        =   27
               Top             =   510
               Width           =   1425
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1680
               TabIndex        =   26
               Top             =   270
               Width           =   1425
            End
            Begin VB.Label Label21 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Precio Venta   :"
               Height          =   240
               Left            =   150
               TabIndex        =   25
               Top             =   270
               Width           =   1545
            End
            Begin VB.Label Label22 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Precio Costo   :"
               Height          =   240
               Left            =   150
               TabIndex        =   24
               Top             =   510
               Width           =   1545
            End
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cerrar"
         Height          =   270
         Left            =   5985
         TabIndex        =   4
         Top             =   3015
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Importes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3075
         Width           =   990
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Unidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3075
         Width           =   990
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cpbtes."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3075
         Width           =   990
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "  COMPROBANTES  "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         Left            =   120
         TabIndex        =   68
         Top             =   105
         Visible         =   0   'False
         Width           =   7125
         Begin VB.Frame Frame11 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Exportación  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   795
            Left            =   3585
            TabIndex        =   100
            Top             =   1980
            Width           =   3420
            Begin VB.Label Label129 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""E"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   104
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label131 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   103
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label Label130 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Créditos  ""E"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   102
               Top             =   480
               Width           =   1440
            End
            Begin VB.Label Label132 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   101
               Top             =   480
               Width           =   1650
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Cta. Cte.  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   2220
            Left            =   120
            TabIndex        =   82
            Top             =   240
            Width           =   3420
            Begin VB.Label Label27 
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   99
               Top             =   705
               Width           =   1650
            End
            Begin VB.Label Label107 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""C"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   98
               Top             =   1680
               Width           =   1440
            End
            Begin VB.Label Label110 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   97
               Top             =   480
               Width           =   1635
            End
            Begin VB.Label Label109 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   96
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label Label101 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""A"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   95
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label102 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""B"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   94
               Top             =   480
               Width           =   1440
            End
            Begin VB.Label Label103 
               BackColor       =   &H00D0FBD5&
               Caption         =   "N/Créd.   ""A"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   93
               Top             =   720
               Width           =   1440
            End
            Begin VB.Label Label108 
               BackColor       =   &H00D0FBD5&
               Caption         =   "N/Créd.  ""C"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   92
               Top             =   1920
               Width           =   1440
            End
            Begin VB.Label Label104 
               BackColor       =   &H00D0FBD5&
               Caption         =   "N/Créd.   ""B"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   91
               Top             =   960
               Width           =   1440
            End
            Begin VB.Label Label106 
               BackColor       =   &H00D0FBD5&
               Caption         =   "N/Déb.    ""B"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   90
               Top             =   1440
               Width           =   1440
            End
            Begin VB.Label Label105 
               BackColor       =   &H00D0FBD5&
               Caption         =   "N/Déb.    ""A"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   89
               Top             =   1200
               Width           =   1440
            End
            Begin VB.Label Label111 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   88
               Top             =   720
               Width           =   1635
            End
            Begin VB.Label Label112 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   87
               Top             =   960
               Width           =   1635
            End
            Begin VB.Label Label113 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   86
               Top             =   1200
               Width           =   1635
            End
            Begin VB.Label Label114 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   85
               Top             =   1440
               Width           =   1635
            End
            Begin VB.Label Label115 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   84
               Top             =   1680
               Width           =   1635
            End
            Begin VB.Label Label116 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   83
               Top             =   1920
               Width           =   1635
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Mostrador   "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   1725
            Left            =   3585
            TabIndex        =   69
            Top             =   240
            Width           =   3420
            Begin VB.Label Label118 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""B"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   81
               Top             =   480
               Width           =   1440
            End
            Begin VB.Label Label117 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""A"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   80
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label123 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   79
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label Label124 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   78
               Top             =   480
               Width           =   1635
            End
            Begin VB.Label Label119 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Créditos  ""A"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   77
               Top             =   720
               Width           =   1440
            End
            Begin VB.Label Label125 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   76
               Top             =   720
               Width           =   1650
            End
            Begin VB.Label Label126 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   75
               Top             =   960
               Width           =   1635
            End
            Begin VB.Label Label127 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   74
               Top             =   1200
               Width           =   1635
            End
            Begin VB.Label Label128 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   73
               Top             =   1440
               Width           =   1635
            End
            Begin VB.Label Label120 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Créditos  ""B"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   72
               Top             =   960
               Width           =   1440
            End
            Begin VB.Label Label121 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""C"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   71
               Top             =   1200
               Width           =   1440
            End
            Begin VB.Label Label122 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Créditos  ""C"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   70
               Top             =   1440
               Width           =   1440
            End
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "  UNIDADES  "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2865
         Left            =   120
         TabIndex        =   37
         Top             =   105
         Visible         =   0   'False
         Width           =   7125
         Begin VB.Frame Frame13 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Cta. Cte.  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   1290
            Left            =   75
            TabIndex        =   59
            Top             =   240
            Width           =   3420
            Begin VB.Label Label203 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   67
               Top             =   705
               Width           =   1650
            End
            Begin VB.Label Label193 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""C"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   66
               Top             =   705
               Width           =   1440
            End
            Begin VB.Label Label202 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   65
               Top             =   480
               Width           =   1635
            End
            Begin VB.Label Label201 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   64
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label Label191 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""A"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   63
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label192 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""B"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   62
               Top             =   480
               Width           =   1440
            End
            Begin VB.Label Label194 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Devolucion   :"
               Height          =   240
               Left            =   150
               TabIndex        =   61
               Top             =   960
               Width           =   1440
            End
            Begin VB.Label Label204 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   60
               Top             =   960
               Width           =   1650
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Mostrador   "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   1275
            Left            =   3585
            TabIndex        =   50
            Top             =   240
            Width           =   3420
            Begin VB.Label Label206 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""B"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   58
               Top             =   480
               Width           =   1440
            End
            Begin VB.Label Label205 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""A"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   57
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label Label209 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   56
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label Label210 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   55
               Top             =   480
               Width           =   1635
            End
            Begin VB.Label Label207 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Facturas ""C"" :"
               Height          =   240
               Left            =   150
               TabIndex        =   54
               Top             =   705
               Width           =   1440
            End
            Begin VB.Label Label211 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   53
               Top             =   705
               Width           =   1650
            End
            Begin VB.Label Label208 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Devolución    :"
               Height          =   240
               Left            =   150
               TabIndex        =   52
               Top             =   960
               Width           =   1440
            End
            Begin VB.Label Label212 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1560
               TabIndex        =   51
               Top             =   960
               Width           =   1650
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Unidades Promedio x Día  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   1020
            Left            =   3585
            TabIndex        =   43
            Top             =   1575
            Width           =   3420
            Begin VB.Label Label218 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Por Mostrador :"
               Height          =   240
               Left            =   135
               TabIndex        =   49
               Top             =   525
               Width           =   1545
            End
            Begin VB.Label Label217 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Por Cta. Cte.    :"
               Height          =   240
               Left            =   135
               TabIndex        =   48
               Top             =   285
               Width           =   1545
            End
            Begin VB.Label Label220 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1665
               TabIndex        =   47
               Top             =   285
               Width           =   1425
            End
            Begin VB.Label Label221 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1665
               TabIndex        =   46
               Top             =   525
               Width           =   1425
            End
            Begin VB.Label Label219 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Por Tipo ""C""    :"
               Height          =   240
               Left            =   135
               TabIndex        =   45
               Top             =   750
               Width           =   1545
            End
            Begin VB.Label Label222 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1665
               TabIndex        =   44
               Top             =   750
               Width           =   1425
            End
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H00D0FBD5&
            Caption         =   "  Exportación  "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H003D813A&
            Height          =   1020
            Left            =   75
            TabIndex        =   38
            Top             =   1575
            Width           =   3420
            Begin VB.Label Label216 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1680
               TabIndex        =   42
               Top             =   615
               Width           =   1425
            End
            Begin VB.Label Label215 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00D0FBD5&
               Height          =   240
               Left            =   1680
               TabIndex        =   41
               Top             =   375
               Width           =   1425
            End
            Begin VB.Label Label213 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Factura ""E""    :"
               Height          =   240
               Left            =   150
               TabIndex        =   40
               Top             =   375
               Width           =   1545
            End
            Begin VB.Label Label214 
               BackColor       =   &H00D0FBD5&
               Caption         =   "Devolución     :"
               Height          =   240
               Left            =   150
               TabIndex        =   39
               Top             =   615
               Width           =   1545
            End
         End
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid GRILLA2 
      Height          =   3555
      Left            =   30
      TabIndex        =   129
      Top             =   840
      Width           =   11145
      _cx             =   19659
      _cy             =   6271
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"TotFact.frx":003A
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
   Begin MSComCtl2.DTPicker FECHA 
      Height          =   330
      Left            =   5250
      TabIndex        =   130
      Top             =   60
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
      Format          =   21037059
      CurrentDate     =   36877
   End
   Begin MSAdodcLib.Adodc FCIVA 
      Height          =   330
      Left            =   4290
      Top             =   5985
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   582
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
      Connect         =   "DSN=Fcmenu"
      OLEDBString     =   "DSN=Fcmenu"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "clever"
      Password        =   ""
      RecordSource    =   "Movstock"
      Caption         =   "FCIVA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc TOTL 
      Height          =   330
      Left            =   6180
      Top             =   5970
      Visible         =   0   'False
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   582
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
      Connect         =   "DSN=Fcmenu"
      OLEDBString     =   "DSN=Fcmenu"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "clever"
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "TOTL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label300 
      Caption         =   "Últimos "
      Height          =   240
      Left            =   4470
      TabIndex        =   135
      Top             =   495
      Width           =   960
   End
   Begin VB.Label Label400 
      Caption         =   "Muestra desde"
      Height          =   240
      Left            =   3780
      TabIndex        =   134
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label Label500 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D0FBD5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   2220
      TabIndex        =   133
      Top             =   4485
      Width           =   1500
   End
   Begin VB.Label Label700 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D0FBD5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   5805
      TabIndex        =   132
      Top             =   4485
      Width           =   1500
   End
   Begin VB.Label Label600 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D0FBD5&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   3990
      TabIndex        =   131
      Top             =   4485
      Width           =   1500
   End
End
Attribute VB_Name = "TotFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TotPvta As Currency, TotPcos As Currency, TotPesp As Currency
Dim LaF As Byte, diashab As Integer

Private Sub Command4_Click()
    
    MuestraTotal

End Sub

Private Sub Form_Load()

FECHA.Value = Date
Text2.Text = 1

Me.Move 0, 0, 11650, 5350
LaF = 0
DoVer2

End Sub

Private Sub Form_Deactivate()

Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyEscape Then Unload Me

If KeyCode = vbKeyF1 Or KeyCode = vbKeyReturn Then
    LaF = 0
    MuestraDetalle
End If

If KeyCode = vbKeyF2 Then
    LaF = 1
    MuestraDetalle
End If

If KeyCode = vbKeyF3 Then
    LaF = 2
    MuestraDetalle
End If

If KeyCode = vbKeyF4 Then
    MuestraTotal
End If


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

Private Sub Command3_Click()

    Picture2.Visible = False

End Sub

Private Sub FECHA_Change()

DoVer2

End Sub

Private Sub Option1_Click(Index As Integer)

Frame7.Visible = False
Frame8.Visible = False
Frame12.Visible = False

Select Case Index
    
    Case 0
        Frame7.Visible = True
    Case 1
        Frame12.Visible = True
    Case 2
        Frame8.Visible = True

End Select

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode < vbKey0 Or KeyCode > vbKey9 Then KeyCode = 0

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode < vbKey0 Or KeyCode > vbKey9 Then KeyCode = 0

End Sub

Sub DoVer2()
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2

FECHADsd = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year
TotPvta = 0: TotPcos = 0: TotPesp = 0
ElDia = 31
If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
FECHAHst = FECHA.Month & "/" & ElDia & "/" & FECHA.Year

'///           ----------------------------------------
       
LimpiaGrilla
          
TOTL.RecordSource = "Select * FROM Totales WHERE FECHA >= " & Separa & _
                    FECHADsd & Separa & " AND FECHA <= " & Separa & FECHAHst & Separa & " Order by FECHA"
TOTL.Refresh
If TOTL.Recordset.EOF Then Exit Sub

TOTL.Recordset.MoveFirst
    
paraSTATUS = "Seleccionados : " & TOTL.Recordset.RecordCount
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

       CargaGrilla

  Label500.Caption = Format(TotPvta, "###,##0.00")
  Label600.Caption = Format(TotPcos, "###,##0.00")
  Label700.Caption = Format(TotPesp, "###,##0.00")

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
Dim TOTFAC        As Double
Dim TOTNC         As Double
Dim PORCEN        As Currency
Dim PVta          As Currency
Dim PCos          As Currency
Dim PEsp          As Currency

Dim I1, i2, Tc
Dim i3 As Long

TotPvta = 0: TotPcos = 0: TotPesp = 0
TOTFAC = 0: TOTNC = 0: PORCEN = 0

GRILLA2.Col = 0

Do Until TOTL.Recordset.EOF Or i3 >= Combo1.Text

i3 = i3 + 1

FECHA = TOTL.Recordset!FECHA

PVta = TOTL.Recordset!PVta
PCos = TOTL.Recordset!PCos
PEsp = TOTL.Recordset!PEsp

TOTFAC = TOTL.Recordset!faca + TOTL.Recordset!facb + TOTL.Recordset!expa + _
         TOTL.Recordset!mosta + TOTL.Recordset!mostb

TOTNC = TOTL.Recordset!nca + TOTL.Recordset!ncb + TOTL.Recordset!ncexpa + _
         TOTL.Recordset!ncmosta + TOTL.Recordset!ncmostb


GRILLA2.AddItem FECHA & vbTab & PVta & vbTab & PCos & vbTab & PEsp & vbTab & _
            TOTFAC & vbTab & TOTNC & vbTab & 0 & vbTab & FECHA, GRILLA2.Rows
    
paraSTATUS = "Cargando : " & (GRILLA2.Rows - 1)
FCMENU.StatusBar1.Panels.item(2).Text = paraSTATUS

Me.Show
DoEvents

Salta:

TotPvta = TotPvta + PVta
TotPcos = TotPcos + PCos
TotPesp = TotPesp + PEsp

TOTL.Recordset.MoveNext
Loop

' GRILLA2.AutoSize 0, 6

If GRILLA2.Rows > 1 Then GRILLA2.Row = 1


' FG1.SetFocus

End Sub

Private Sub GRILLA2_Click()

If GRILLA2.Rows < 1 Then Exit Sub

MuestraDetalle

End Sub

Private Sub MuestraDetalle()
Dim Promed As Currency
          
TOTL.RecordSource = "Select * FROM Totales WHERE FECHA = " & Separa & _
                    GRILLA2.TextMatrix(GRILLA2.Row, 7) & Separa
TOTL.Refresh
If TOTL.Recordset.EOF Then Exit Sub

'   Importes
Label4.Caption = Format(TOTL.Recordset!PESPA, "###,##0.00")
Label5.Caption = Format(TOTL.Recordset!PESPb, "###,##0.00")
Label6.Caption = Format(TOTL.Recordset!PESPc, "###,##0.00")

Label14.Caption = Format(TOTL.Recordset!PESPEXPA, "###,##0.00")

Label10.Caption = Format(TOTL.Recordset!pespmostA, "###,##0.00")
Label11.Caption = Format(TOTL.Recordset!pespmostB, "###,##0.00")
Label12.Caption = Format(TOTL.Recordset!pespmostC, "###,##0.00")

Label24.Caption = Format(TOTL.Recordset!PVta, "###,##0.00")
Label25.Caption = Format(TOTL.Recordset!PCos, "###,##0.00")
Label26.Caption = Format(TOTL.Recordset!PEsp, "###,##0.00")

If (TOTL.Recordset!UNIDA + TOTL.Recordset!UNIDB) <> 0 Then
        Promed = (TOTL.Recordset!PESPA + TOTL.Recordset!PESPb) / _
                 (TOTL.Recordset!UNIDA + TOTL.Recordset!UNIDB)
End If
    Label18.Caption = Format(Promed, "###,##0.00"): Promed = 0

If (TOTL.Recordset!UNMOSTA + TOTL.Recordset!UNMOSTB) <> 0 Then
    Promed = (TOTL.Recordset!pespmostA + TOTL.Recordset!pespmostB) / _
             (TOTL.Recordset!UNMOSTA + TOTL.Recordset!UNMOSTB)
    End If
    Label19.Caption = Format(Promed, "###,##0.00"): Promed = 0


If (TOTL.Recordset!UNIDC + TOTL.Recordset!UNMOSTC) <> 0 Then
        Promed = (TOTL.Recordset!PESPc + TOTL.Recordset!pespmostC) / _
                 (TOTL.Recordset!UNIDC + TOTL.Recordset!UNMOSTC)
         End If
    Label20.Caption = Format(Promed, "###,##0.00"): Promed = 0

'   Unidades
Label201.Caption = Format(TOTL.Recordset!UNIDA, "###,##0")
Label202.Caption = Format(TOTL.Recordset!UNIDB, "###,##0")
Label203.Caption = Format(TOTL.Recordset!UNIDC, "###,##0")
Label204.Caption = Format(TOTL.Recordset!DEV, "###,##0")

Label209.Caption = Format(TOTL.Recordset!UNMOSTA, "###,##0")
Label210.Caption = Format(TOTL.Recordset!UNMOSTB, "###,##0")
Label211.Caption = Format(TOTL.Recordset!UNMOSTC, "###,##0")
Label212.Caption = Format(TOTL.Recordset!MOSTDEV, "###,##0")

Label215.Caption = Format(TOTL.Recordset!unexpA, "###,##0")
Label216.Caption = Format(TOTL.Recordset!expdev, "###,##0")

If (TOTL.Recordset!faca + TOTL.Recordset!facb) <> 0 Then
        Promed = (TOTL.Recordset!UNIDA + TOTL.Recordset!UNIDB) / _
                 (TOTL.Recordset!faca + TOTL.Recordset!facb)
End If
        Label220.Caption = Format(Promed, "###,##0.00"): Promed = 0
        
If (TOTL.Recordset!mosta + TOTL.Recordset!mostb) <> 0 Then
        Promed = (TOTL.Recordset!UNMOSTA + TOTL.Recordset!UNMOSTB) / _
                 (TOTL.Recordset!mosta + TOTL.Recordset!mostb)
End If
        Label221.Caption = Format(Promed, "###,##0.00"): Promed = 0
        
If (TOTL.Recordset!facc + TOTL.Recordset!mostc) <> 0 Then
        Promed = (TOTL.Recordset!UNIDC + TOTL.Recordset!UNMOSTC) / _
                 (TOTL.Recordset!facc + TOTL.Recordset!mostc)
End If
        Label222.Caption = Format(Promed, "###,##0.00"): Promed = 0

'   Comprobantes
Label109.Caption = Format(TOTL.Recordset!faca, "###,##0")
Label110.Caption = Format(TOTL.Recordset!facb, "###,##0")
Label111.Caption = Format(TOTL.Recordset!nca, "###,##0")
Label112.Caption = Format(TOTL.Recordset!ncb, "###,##0")
Label113.Caption = Format(TOTL.Recordset!nda, "###,##0")
Label114.Caption = Format(TOTL.Recordset!ndb, "###,##0")
Label115.Caption = Format(TOTL.Recordset!facc, "###,##0")
Label116.Caption = Format(TOTL.Recordset!ncc, "###,##0")

Label123.Caption = Format(TOTL.Recordset!mosta, "###,##0")
Label124.Caption = Format(TOTL.Recordset!mostb, "###,##0")
Label125.Caption = Format(TOTL.Recordset!ncmosta, "###,##0")
Label126.Caption = Format(TOTL.Recordset!ncmostb, "###,##0")
Label127.Caption = Format(TOTL.Recordset!mostc, "###,##0")
Label128.Caption = Format(TOTL.Recordset!ncmostc, "###,##0")

Label131.Caption = Format(TOTL.Recordset!expa, "###,##0")
Label132.Caption = Format(TOTL.Recordset!expdev, "###,##0")

      Picture1.Visible = True

Option1.item(LaF).Value = True

End Sub

Private Sub MuestraTotal()
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2
Dim Proyec As Currency

FECHADsd = FECHA.Value
ElDia = 31
If FECHA.Month = 4 Or FECHA.Month = 6 Or FECHA.Month = 9 Or FECHA.Month = 11 Then ElDia = 30
    
    If FECHA.Month = 2 Then
           ElDia = DateAdd("m", 1, "31/01/" & FECHA.Year)
    End If
    
FECHAHst = ElDia & "/" & FECHA.Month & "/" & FECHA.Year


Text1.Text = DiasHabiles(Val(ElDia))
Text2.Text = DiasHabiles(FECHA.Day)

'///           ----------------------------------------
    
    TOTL.RecordSource = "Select  SUM(pvta) as TotVta, SUM(pcos) as TotCos, SUM(pesp)" & _
                        " as TotEsp, SUM(UNIDA + UNIDB + UNIDC + UNMOSTA + UNMOSTB +" & _
                        " UNMOSTC - DEV - MOSTDEV) as TotUni" & _
                        " FROM Totales WHERE FECHA >= " & Separa & _
                        FECHADsd & Separa & " AND FECHA < " & Separa & FECHAHst & Separa

TOTL.Refresh
If TOTL.Recordset.EOF Then Exit Sub
If IsNull(TOTL.Recordset!totvta) And _
    IsNull(TOTL.Recordset!totcos) And _
    IsNull(TOTL.Recordset!totesp) And _
    IsNull(TOTL.Recordset!totuni) Then Exit Sub

Label29.Caption = Format(TOTL.Recordset!totvta, "###,##0.00")
Label31.Caption = Format(TOTL.Recordset!totcos, "###,##0.00")
Label35.Caption = Format(TOTL.Recordset!totesp, "###,##0.00")
Label34.Caption = Format(TOTL.Recordset!totuni, "###,##0.00")

Proyec = (TOTL.Recordset!totesp / Text1.Text) * Text2.Text
Label42.Caption = Format(Proyec, "###,##0.00")
Proyec = TOTL.Recordset!totesp / Text2.Text
Label40.Caption = Format(Proyec, "###,##0.00")
Proyec = TOTL.Recordset!totuni / Text2.Text
Label36.Caption = Format(Proyec, "###,##0.00")

Picture2.Visible = True
Text1.SetFocus

End Sub

Private Function DiasHabiles(DiaHst As Integer) As Integer
Dim FECHADsd As String, FECHAHst As String, I1%

FECHAHst = DiaHst & "/" & FECHA.Month & "/" & FECHA.Year

DiasHabiles = 0
For I1 = 1 To Val(Mid(FECHAHst, 1, 2))

FECHADsd = I1 & "/" & FECHA.Month & "/" & FECHA.Year

If Weekday(FECHADsd) <> 1 Then DiasHabiles = DiasHabiles + 1

Next I1

If DiasHabiles = 0 Then DiasHabiles = 1

End Function




