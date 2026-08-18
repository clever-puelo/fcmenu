VERSION 5.00
Begin VB.Form ABMClte 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   "Baja"
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
      Left            =   2835
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Dar de Alta un Cliente nuevo"
      Top             =   4860
      Width           =   1035
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Alta"
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
      Left            =   3960
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Dar de Alta un Cliente nuevo"
      Top             =   4860
      Width           =   1035
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cambio"
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
      Left            =   1710
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Buscar un Cliente"
      Top             =   4860
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   29
      Top             =   4830
      Width           =   1300
   End
   Begin VB.Frame Frame5 
      Caption         =   "   Datos de Cobranzas  "
      Enabled         =   0   'False
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
      Height          =   1155
      Left            =   90
      TabIndex        =   34
      Top             =   3645
      Width           =   7935
      Begin VB.TextBox Text302 
         Alignment       =   1  'Right Justify
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
         Left            =   5580
         LinkTimeout     =   5
         MaxLength       =   5
         TabIndex        =   22
         ToolTipText     =   "Comisión Especial "
         Top             =   210
         Width           =   1125
      End
      Begin VB.ComboBox Combo7 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5580
         Style           =   2  'Dropdown List
         TabIndex        =   24
         ToolTipText     =   "Tipo de Cliente"
         Top             =   480
         Width           =   2115
      End
      Begin VB.ComboBox Combo6 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5580
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Tipo de Cliente"
         Top             =   765
         Width           =   2115
      End
      Begin VB.ComboBox Combo5 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4500
         Style           =   2  'Dropdown List
         TabIndex        =   59
         ToolTipText     =   "Tipo de Cliente"
         Top             =   1275
         Width           =   2115
      End
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   25
         ToolTipText     =   "Tipo de Cliente"
         Top             =   780
         Width           =   2115
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "Abmclte.frx":0000
         Left            =   1590
         List            =   "Abmclte.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   23
         ToolTipText     =   "Tipo de Cliente"
         Top             =   480
         Width           =   2565
      End
      Begin VB.TextBox Text301 
         Alignment       =   1  'Right Justify
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
         Left            =   1590
         MaxLength       =   10
         TabIndex        =   21
         ToolTipText     =   "Crédito otorgado"
         Top             =   225
         Width           =   1740
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Comisión Espec.:"
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
         Left            =   4140
         TabIndex        =   62
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label304 
         Alignment       =   1  'Right Justify
         Caption         =   "Categoría :"
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
         Left            =   3150
         TabIndex        =   56
         Top             =   1335
         Width           =   1305
      End
      Begin VB.Label Label306 
         Alignment       =   1  'Right Justify
         Caption         =   "Cobrador :"
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
         Left            =   4230
         TabIndex        =   55
         Top             =   795
         Width           =   1290
      End
      Begin VB.Label Label305 
         Alignment       =   1  'Right Justify
         Caption         =   "Zona :"
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
         Left            =   4230
         TabIndex        =   54
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label303 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendededor :"
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
         Left            =   135
         TabIndex        =   53
         Top             =   795
         Width           =   1425
      End
      Begin VB.Label Label302 
         Alignment       =   1  'Right Justify
         Caption         =   "Cond. Venta :"
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
         Left            =   135
         TabIndex        =   52
         Top             =   540
         Width           =   1425
      End
      Begin VB.Label Label10 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5940
         TabIndex        =   39
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label301 
         Alignment       =   1  'Right Justify
         Caption         =   "Crédito :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   36
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
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
      Height          =   420
      Left            =   6690
      TabIndex        =   28
      Top             =   4830
      Width           =   1300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5250
      TabIndex        =   27
      Top             =   4830
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Datos Nominales  "
      Enabled         =   0   'False
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
      Height          =   2235
      Left            =   90
      TabIndex        =   32
      Top             =   0
      Width           =   7950
      Begin VB.ComboBox Combo1 
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
         Height          =   315
         ItemData        =   "Abmclte.frx":0004
         Left            =   1590
         List            =   "Abmclte.frx":0017
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1860
         Width           =   2565
      End
      Begin VB.TextBox Text109 
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
         Left            =   5520
         MaxLength       =   11
         TabIndex        =   11
         ToolTipText     =   "Número CUIT/CUIL/DNI (sin guiones)"
         Top             =   1875
         Width           =   1935
      End
      Begin VB.TextBox Text108 
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
         Height          =   285
         Index           =   2
         Left            =   5535
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "Fax"
         Top             =   1575
         Width           =   1935
      End
      Begin VB.TextBox Text108 
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
         Index           =   1
         Left            =   3555
         MaxLength       =   20
         TabIndex        =   8
         ToolTipText     =   "Teléfono 2"
         Top             =   1575
         Width           =   1935
      End
      Begin VB.TextBox Text108 
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
         Index           =   0
         Left            =   1575
         MaxLength       =   20
         TabIndex        =   7
         ToolTipText     =   "Teléfono 1"
         Top             =   1575
         Width           =   1935
      End
      Begin VB.TextBox Text107 
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
         Left            =   5535
         LinkTimeout     =   5
         MaxLength       =   2
         TabIndex        =   6
         ToolTipText     =   "Código de Provincia"
         Top             =   1305
         Width           =   555
      End
      Begin VB.TextBox Text106 
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
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   10
         TabIndex        =   5
         ToolTipText     =   "Código Postal"
         Top             =   1305
         Width           =   1020
      End
      Begin VB.TextBox Text105 
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
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Localidad del Cliente"
         Top             =   1035
         Width           =   4530
      End
      Begin VB.TextBox Text104 
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
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   3
         ToolTipText     =   "Dirección del Cliente"
         Top             =   765
         Width           =   4530
      End
      Begin VB.TextBox Text103 
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
         Left            =   6750
         LinkTimeout     =   5
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Clasificación alfabética adicional"
         Top             =   495
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.TextBox Text102 
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
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Razón Social del Cliente"
         Top             =   495
         Width           =   4515
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
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
         Left            =   1575
         MaxLength       =   5
         TabIndex        =   0
         ToolTipText     =   "Código de Cliente"
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label107a 
         Alignment       =   1  'Right Justify
         Caption         =   "Provincia :"
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
         Left            =   4230
         TabIndex        =   57
         Top             =   1350
         Width           =   1305
      End
      Begin VB.Label Label110 
         Alignment       =   1  'Right Justify
         Caption         =   "C.U.I.T. :"
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
         Left            =   4245
         TabIndex        =   48
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Label Label109 
         Alignment       =   1  'Right Justify
         Caption         =   "Condic. I.V.A. :"
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
         Left            =   150
         TabIndex        =   47
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label108 
         Alignment       =   1  'Right Justify
         Caption         =   "Teléfonos :"
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
         Left            =   135
         TabIndex        =   44
         Top             =   1620
         Width           =   1425
      End
      Begin VB.Label Label107b 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5265
         TabIndex        =   43
         Top             =   1350
         Width           =   2505
      End
      Begin VB.Label Label106 
         Alignment       =   1  'Right Justify
         Caption         =   "Cod. Postal :"
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
         Left            =   135
         TabIndex        =   42
         Top             =   1350
         Width           =   1425
      End
      Begin VB.Label Label105 
         Alignment       =   1  'Right Justify
         Caption         =   "Localidad :"
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
         Left            =   135
         TabIndex        =   41
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label104 
         Alignment       =   1  'Right Justify
         Caption         =   "Dirección :"
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
         Left            =   135
         TabIndex        =   40
         Top             =   810
         Width           =   1425
      End
      Begin VB.Label Label102 
         Alignment       =   1  'Right Justify
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
         Left            =   135
         TabIndex        =   38
         Top             =   540
         Width           =   1425
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Código :"
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
         Left            =   135
         TabIndex        =   37
         Top             =   270
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "  Varios  "
      Enabled         =   0   'False
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
      Left            =   5550
      TabIndex        =   49
      Top             =   2220
      Width           =   2475
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "Hab. Fact.  :"
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
         Left            =   300
         TabIndex        =   20
         ToolTipText     =   "Se le envía Etiquetas"
         Top             =   990
         Width           =   1905
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "Lista          :"
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
         Left            =   300
         TabIndex        =   19
         ToolTipText     =   "Se le envía Lista de Precios"
         Top             =   750
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Remito      :"
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
         Left            =   300
         TabIndex        =   17
         ToolTipText     =   "Se emite Remito junto con la Factura"
         Top             =   270
         Width           =   1905
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Flete         :"
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
         Left            =   300
         TabIndex        =   18
         ToolTipText     =   "Paga Flete"
         Top             =   510
         Width           =   1905
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "   Datos de Ventas"
      Enabled         =   0   'False
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
      Left            =   90
      TabIndex        =   33
      Top             =   2220
      Width           =   5445
      Begin VB.CommandButton Command6 
         Caption         =   "Cargar Descuentos"
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
         Left            =   1590
         TabIndex        =   60
         Top             =   1050
         Width           =   3765
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   58
         ToolTipText     =   "Tipo de Cliente"
         Top             =   1590
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.TextBox Text203b 
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
         Left            =   2520
         LinkTimeout     =   5
         MaxLength       =   30
         TabIndex        =   15
         ToolTipText     =   "Detalle del Transporte"
         Top             =   765
         Width           =   2820
      End
      Begin VB.TextBox Text205 
         Alignment       =   1  'Right Justify
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
         Left            =   4635
         LinkTimeout     =   5
         MaxLength       =   2
         TabIndex        =   16
         ToolTipText     =   "Cantidad de Copias de Factura"
         Top             =   1575
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox Text203a 
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
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   3
         TabIndex        =   14
         ToolTipText     =   "Código de Transporte"
         Top             =   765
         Width           =   885
      End
      Begin VB.TextBox Text202 
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
         Left            =   1575
         MaxLength       =   40
         TabIndex        =   13
         ToolTipText     =   "Persona de Contacto para Cobranza"
         Top             =   510
         Width           =   3780
      End
      Begin VB.TextBox Text201 
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
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   12
         ToolTipText     =   "Persona de Contacto para Ventas"
         Top             =   240
         Width           =   3780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Descuentos :"
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
         Left            =   90
         TabIndex        =   61
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label210 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo de Cliente   :"
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
         Left            =   135
         TabIndex        =   51
         Top             =   1620
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label205 
         Alignment       =   1  'Right Justify
         Caption         =   "Copias Fc.  :"
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
         Left            =   3600
         TabIndex        =   50
         Top             =   1620
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label203 
         Alignment       =   1  'Right Justify
         Caption         =   "Transporte :"
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
         Left            =   135
         TabIndex        =   46
         Top             =   810
         Width           =   1425
      End
      Begin VB.Label Label202 
         Alignment       =   1  'Right Justify
         Caption         =   "Contacto Cob.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   45
         ToolTipText     =   "Calificación  del Artículo"
         Top             =   540
         Width           =   1425
      End
      Begin VB.Label Label201 
         Alignment       =   1  'Right Justify
         Caption         =   "Contacto Vta.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   35
         ToolTipText     =   "Calificación  del Artículo"
         Top             =   270
         Width           =   1425
      End
   End
End
Attribute VB_Name = "ABMClte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mensaje As String
Dim CAMPOS      As String, VeNotaC As Boolean
Dim s$

Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset

Private Sub Command7_Click()

TipoMov3 = 2

Grabacion

Unload Me

If BusClte.Visible Then
    BusClte.Command1_Click
End If

End Sub

Private Sub Form_Load()
Dim I1%

Me.Move 0, 0, 8250, 5670

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

VeNotaC = False

Command3.Visible = False

If Not VerDetClt Then
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame3.Enabled = True
    Frame5.Enabled = True
End If
    ' Tabla de Condiciones de Venta
    
miSQL2 = "SELECT * FROM FCTabla1 WHERE CTAB = 'CV   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

I1 = 0

If Not RgTABL.EOF Then
        Do Until RgTABL.EOF
        RSet ClteCVta = Val(RgTABL!COD)
        Combo3.AddItem ClteCVta & "-" & RgTABL!Descri
        Combo3.ItemData(I1) = RgTABL!COD
        RgTABL.MoveNext
        I1 = I1 + 1
        Loop
End If

RgTABL.Close

    ' Tabla de Vendedores
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'VD   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

I1 = 0

If Not RgTABL.EOF Then
        Do Until RgTABL.EOF
        RSet ClteCVta = Val(RgTABL!COD)
        Combo4.AddItem ClteCVta & "-" & RgTABL!Descri
        Combo4.ItemData(I1) = RgTABL!COD
        RgTABL.MoveNext
        I1 = I1 + 1
        Loop
End If

RgTABL.Close

    ' Tabla de zona
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'ZN   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

I1 = 0

If Not RgTABL.EOF Then
        Do Until RgTABL.EOF
        RSet ClteCVta = Val(RgTABL!COD)
        Combo7.AddItem ClteCVta & "-" & RgTABL!Descri
        Combo7.ItemData(I1) = RgTABL!COD
        RgTABL.MoveNext
        I1 = I1 + 1
        Loop
End If

RgTABL.Close

Select Case TipoMov3

Case 1

    Command2.Visible = True
    Blanquea
    
Case 2, 3

    Command3.Visible = True
    Command4.Visible = True
    Command5.Visible = True
    
    miSQL2 = "SELECT  * FROM Clientes WHERE CODIGO = " & CodCLTE
    RgCLTE.Open miSQL2, dbCLTE, adOpenForwardOnly, adLockReadOnly
    
    If RgCLTE.EOF Then
        RgCLTE.Close
        Exit Sub
    End If

    Text101.Enabled = False
    Text101.Text = RgCLTE!CODIGO
    Text102.Text = RgCLTE!Nomb
    Text104.Text = RgCLTE!Dir
    Text105.Text = RgCLTE!Loc
    Text106.Text = RgCLTE!CP
    Text107.Text = RgCLTE!PCIA
    Text108(0).Text = RgCLTE!tel1
    Text108(1).Text = RgCLTE!TEL2
    Text108(2).Text = RgCLTE!FAX
    Text109.Text = RgCLTE!Cuit
    Combo1.ListIndex = RgCLTE!CIVA - 1
   
    Check1.Value = 0: Check2.Value = 0: Check3.Value = 0: Check4.Value = 0:
'    If rgCLTE!MCARMTO = 1 Then Check1.Value = 1
'    If rgCLTE!MCAFLET = 1 Then Check2.Value = 1
'    If rgCLTE!mcalis = 1 Then Check3.Value = 1
     If RgCLTE!CANAL = 9 Then Check4.Value = 1
    
    Text301.Text = RgCLTE!CREDIT
    Text302.Text = RgCLTE!PROIND
    
'If Combo3.ListCount > (RgCLTE!CVTA - 1) Then
        Combo3.ListIndex = BuscaEnCombo(Combo3, RgCLTE!CVTA)
'        Combo3.ListIndex = RgCLTE!CVTA - 1
'Else
'        LaTabla = "Cond. Vta."
'        ErrorEnTabla
'       End If
        
'If Combo4.ListCount > (RgCLTE!VEND - 1) Then
        Combo4.ListIndex = BuscaEnCombo(Combo4, RgCLTE!VEND)
'        Combo4.ListIndex = RgCLTE!VEND - 1
'Else
'        LaTabla = "Vendedores"
'        ErrorEnTabla
'        End If
        
        
'If Combo7.ListCount > (RgCLTE!ZONA - 1) Then
        Combo7.ListIndex = BuscaEnCombo(Combo7, RgCLTE!ZONA)
'       Combo7.ListIndex = RgCLTE!zona - 1
'Else
'        LaTabla = "Zonas"
'        ErrorEnTabla
'         End If
        
RgCLTE.Close
        
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbCLTE.Close
' CodCLTE = 0

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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF2 Then
    DeDonde = 1
    BusClte.Show
    Unload Me
        End If


If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 39 Then SendKeys "{TAB}"
If KeyCode = 38 Or KeyCode = 37 Then SendKeys "+{TAB}"
If KeyCode = 27 Then End

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Command1_Click()

Unload DtosxClte
Unload Me

End Sub

Private Sub Command2_Click()

Grabacion
Unload Me

End Sub

Private Sub Command3_Click()

NOTACLTE.Show

End Sub

Private Sub Command4_Click()

If Frame1.Enabled Then
                Caption = "Detalle del Cliente "
                Frame1.Enabled = False
                Frame2.Enabled = False
                Frame3.Enabled = False
                Frame5.Enabled = False
                Command2.Visible = False
                Command4.Caption = "Cambio"
'                Command5.Visible = False
'                Text102.SetFocus
Else

                Caption = "Cambios al Cliente "
                Frame1.Enabled = True
                Frame2.Enabled = True
                Frame3.Enabled = True
                Frame5.Enabled = True
                Command2.Visible = True
                Command4.Caption = "Sólo Ver"
'                Command5.Visible = True
                Text102.SetFocus
End If

End Sub

Private Sub Command5_Click()

Command3.Visible = False
Command2.Visible = True
Command5.Visible = False
                
                Frame1.Enabled = True
                Frame2.Enabled = True
                Frame3.Enabled = True
                Frame5.Enabled = True

    Blanquea

End Sub

Private Sub Command6_Click()

DtosxClte.Show

End Sub

'*************************************
'   Ingresa Código de Cliente
'*************************************

Private Sub Text101_GotFocus()

Text101.SelStart = 0
Text101.SelLength = Len(Text101.Text)

End Sub

Private Sub Text102_GotFocus()

Text102.SelStart = 0
Text102.SelLength = Len(Text102.Text)

End Sub

Private Sub Text103_GotFocus()

Text103.SelStart = 0
Text103.SelLength = Len(Text103.Text)

End Sub

Private Sub Text104_GotFocus()

Text104.SelStart = 0
Text104.SelLength = Len(Text104.Text)

End Sub

Private Sub Text105_GotFocus()

Text105.SelStart = 0
Text105.SelLength = Len(Text105.Text)

End Sub

Private Sub Text101_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode <> vbKeyReturn Then Exit Sub

End Sub

Private Sub Text101_LostFocus()
    
    miSQL2 = "SELECT * FROM Clientes WHERE CODIGO = " & Val(Text101.Text)
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

If Not RgTABL.EOF And TipoMov3 = 1 Then
  vbMsgBoxTitle = "Error de Ingreso"
  vbMsgBoxText = "  " & vbCrLf & "Código  YA  Existe  "
  vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
   MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle
   Text101.SetFocus
End If

If RgTABL.EOF And TipoMov3 <> 1 Then
  vbMsgBoxTitle = "Error de Ingreso"
  vbMsgBoxText = "  " & vbCrLf & "Código  NO  Existe  "
  vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
   MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle
   Text101.SetFocus
End If
   
   RgTABL.Close

End Sub


'*************************************
'   Ingresa Cod.Postal
'*************************************

Private Sub Text106_GotFocus()

Text106.SelStart = 0
Text106.SelLength = Len(Text106.Text)

End Sub

'*************************************
'   Ingresa Provincia
'*************************************

Private Sub Text107_GotFocus()

Text107.SelStart = 0
Text107.SelLength = Len(Text107.Text)

End Sub

'*************************************
'   Ingresa Teléfonos y Fax
'*************************************

Private Sub Text108_GotFocus(Index As Integer)

Text108(Index).SelStart = 0
Text108(Index).SelLength = Len(Text108(Index).Text)

End Sub

'*************************************
'   Ingresa CUIT
'*************************************

Private Sub Text109_GotFocus()

Text109.SelStart = 0
Text109.SelLength = Len(Text109.Text)
End Sub

Private Sub Text109_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text109_Validate(Cancel As Boolean)

If Not Cuit(Text109.Text) Then Cancel = True

End Sub

'*************************************
'   Ingresa Contacto de Ventas
'*************************************

Private Sub Text201_GotFocus()

Text201.SelStart = 0
Text201.SelLength = Len(Text201.Text)

End Sub

'*************************************
'   Ingresa Contacto de Cobranzas
'*************************************

Private Sub Text202_GotFocus()

Text202.SelStart = 0
Text202.SelLength = Len(Text202.Text)

End Sub

'*************************************
'   Ingresa Cod.Transporte
'*************************************

Private Sub Text203a_GotFocus()

Text203a.SelStart = 0
Text203a.SelLength = Len(Text203a.Text)

End Sub

Private Sub Text203a_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'*************************************
'   Ingresa Detalle Transporte
'*************************************

Private Sub Text203b_GotFocus()

Text203b.SelStart = 0
Text203b.SelLength = Len(Text203b.Text)

End Sub

'*************************************
'   Ingresa Copias de Factura
'*************************************

Private Sub Text205_GotFocus()

Text205.SelStart = 0
Text205.SelLength = Len(Text205.Text)

End Sub

Private Sub Text205_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'*************************************
'   Ingresa Crédito
'*************************************

Private Sub Text301_GotFocus()

Text301.SelStart = 0
Text301.SelLength = Len(Text301.Text)

End Sub

Private Sub Text301_KeyPress(KeyAscii As Integer)


If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub
            
            If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
            If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Grabacion()
Dim I1%, i2%, i3%

Select Case TipoMov3
        Case 1
                Mensaje = "Desea Grabar ?"
        Case 2
                Mensaje = "Desea Eliminar ?"
        Case 3
                Mensaje = "Desea Cambiar ? "
End Select

Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "ABM de Clientes")

If Respuesta = vbNo Then Exit Sub
If Respuesta = vbCancel Then
         Exit Sub
         Me.Show
         End If

miSQL2 = "SELECT  * FROM Clientes WHERE CODIGO = " & CodCLTE
RgCLTE.Open miSQL2, dbCLTE, adOpenDynamic, adLockPessimistic

If TipoMov3 = 2 Then
        RgCLTE.Delete
        Exit Sub
        End If
        
     
'******************************************
'           Graba en Clientes
'******************************************

    If RgCLTE.EOF = True Then
        RgCLTE.AddNew
        
        RgCLTE!CODIGO = CodCLTE
        RgCLTE!NGAN = 0
        RgCLTE!IB = 0
        RgCLTE!corr1 = 0
        RgCLTE!CORR2 = 0
        RgCLTE!FALTA = Date
        RgCLTE!DEUDA = 0
        RgCLTE!CREDIT = 0
    End If
    
    RgCLTE!CODIGO = Text101.Text
    RgCLTE!Nomb = Text102.Text
    RgCLTE!Dir = Text104.Text
    RgCLTE!Loc = Text105.Text
    RgCLTE!CP = Text106.Text
    RgCLTE!PCIA = Text107.Text
    RgCLTE!tel1 = Text108(0).Text
    RgCLTE!TEL2 = Text108(1).Text
    RgCLTE!FAX = Text108(2).Text
    RgCLTE!Cuit = Text109.Text
'    RgCLTE!CONTAC = Text201.Text
    RgCLTE!email = Text202.Text
'    RgCLTE!CODTRANS = Text203a.Text
'    RgCLTE!TRANSP = Text203b.Text
    RgCLTE!dto1 = 0
    RgCLTE!dto2 = 0
    RgCLTE!dto3 = 0
    RgCLTE!dto4 = 0
    RgCLTE!dto5 = 0

RgCLTE!CIVA = Combo1.ListIndex + 1
'RgCLTE!TIPOCLTE = Val(Mid(Combo2.Text, 1, 1))
RgCLTE!CVTA = Val(Mid(Combo3.Text, 1, 2))
RgCLTE!VEND = Val(Mid(Combo4.Text, 1, 2))
'RgCLTE!Categ = Val(Mid(Combo5.Text, 1, 1))
RgCLTE!ZONA = Val(Mid(Combo7.Text, 1, 2))
'rgCLTE!COBR = Val(Mid(COMBO7.Text, 1, 2))

RgCLTE!FACTUAL = Date
RgCLTE!USUARIO = Mid$(Red_Usuario, 1, 6)
    
    RgCLTE!CANAL = 0
If Check4.Value = 1 Then RgCLTE!CANAL = 9
    
    RgCLTE!PROIND = Text302.Text

RgCLTE.Update
RgCLTE.Close

End Sub

Sub MuestraError()
          
        vbMsgBoxTitle = "Error de Ingreso"
        vbMsgBoxText = "  " & vbCrLf & "Código  NO  Existe  "
        vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
        MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle

End Sub

Sub Blanquea()

miSQL2 = "SELECT  TOP 1 CODIGO FROM Clientes WHERE CODIGO < 99900 " _
                                & "ORDER BY CODIGO DESC"
RgCLTE.Open miSQL2, dbCLTE, adOpenForwardOnly, adLockReadOnly

    If RgCLTE.EOF = True Then
        RgCLTE!CODIGO = 0
    End If

    Text101.Text = RgCLTE!CODIGO + 1
    Text102.Text = ""
    Text103.Text = ""
    Text104.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text107.Text = ""
    Text108(0).Text = ""
    Text108(1).Text = ""
    Text108(2).Text = ""
    Text109.Text = ""
    Combo1.ListIndex = 0
    
    Text201.Text = ""
    Text202.Text = ""
    Text203a.Text = 0
    Text302.Text = 0
    Text203b.Text = ""
'    Combo2.ListIndex = 3
    Text205.Text = 1
   
    Check1.Value = 0: Check2.Value = 0: Check3.Value = 0: Check4.Value = 0
    
    Text301.Text = 0
    Combo3.ListIndex = 5
'    Combo4.ListIndex = 1
'    Combo5.ListIndex = 0
'    COMBO6.ListIndex = 0
    Combo7.ListIndex = 0

Text101.Enabled = True

Command4.Visible = False
Command5.Visible = False

CodCLTE = RgCLTE!CODIGO + 1
ClteNomb = ""

Text102.TabIndex = 0

RgCLTE.Close

End Sub

Sub ErrorEnTabla()
                
Mensaje = "Desea Continuar ? "
Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNo, "Error en Tabla de " & LaTabla)

If Respuesta = vbNo Then Exit Sub
If Respuesta = vbCancel Then
         Exit Sub
         Me.Show
         End If

End Sub

Private Sub Text302_GotFocus()

Text302.SelStart = 0
Text302.SelLength = Len(Text302.Text)

End Sub

Private Sub Text302_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub
            
            If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
            If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub
