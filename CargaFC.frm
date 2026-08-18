VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CargaFC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Carga de Facturas"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9150
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
   ScaleHeight     =   5310
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   7410
      TabIndex        =   22
      Top             =   4845
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5985
      TabIndex        =   21
      Top             =   4845
      Width           =   1320
   End
   Begin VB.Frame V 
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
      Height          =   4680
      Left            =   390
      TabIndex        =   24
      Top             =   90
      Width           =   8445
      Begin VB.TextBox Text1 
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
         Index           =   12
         Left            =   5370
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   48
         Top             =   4320
         Width           =   2100
      End
      Begin VB.TextBox Text1 
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
         Index           =   11
         Left            =   1620
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   47
         Top             =   4320
         Width           =   2100
      End
      Begin VB.ComboBox Combo8 
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
         ItemData        =   "CargaFC.frx":0000
         Left            =   4710
         List            =   "CargaFC.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   2670
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.TextBox Text1 
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
         Index           =   10
         Left            =   5370
         LinkTimeout     =   5
         MaxLength       =   5
         TabIndex        =   20
         Top             =   4050
         Width           =   1050
      End
      Begin VB.TextBox Text1 
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
         Index           =   9
         Left            =   1620
         LinkTimeout     =   5
         MaxLength       =   2
         TabIndex        =   19
         Top             =   4050
         Width           =   600
      End
      Begin VB.TextBox Text1 
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
         Index           =   8
         Left            =   1620
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   18
         Top             =   3780
         Width           =   2100
      End
      Begin VB.TextBox Text1 
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
         Index           =   7
         Left            =   1620
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   16
         Top             =   3510
         Width           =   2100
      End
      Begin VB.TextBox Text1 
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
         Index           =   6
         Left            =   5370
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   17
         Top             =   3510
         Width           =   2100
      End
      Begin VB.ComboBox Combo7 
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
         ItemData        =   "CargaFC.frx":0048
         Left            =   4710
         List            =   "CargaFC.frx":004A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2670
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.ComboBox Combo6 
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
         ItemData        =   "CargaFC.frx":004C
         Left            =   1620
         List            =   "CargaFC.frx":004E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2670
         Width           =   2310
      End
      Begin VB.ComboBox Combo5 
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
         ItemData        =   "CargaFC.frx":0050
         Left            =   4710
         List            =   "CargaFC.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2370
         Width           =   2310
      End
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "CargaFC.frx":0054
         Left            =   1620
         List            =   "CargaFC.frx":0056
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2370
         Width           =   2310
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
         ItemData        =   "CargaFC.frx":0058
         Left            =   1620
         List            =   "CargaFC.frx":006B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2070
         Width           =   1935
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Index           =   3
         Left            =   4710
         MaxLength       =   13
         TabIndex        =   9
         Top             =   2070
         Width           =   2265
      End
      Begin VB.TextBox Text1 
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
         Index           =   1
         Left            =   2670
         MaxLength       =   30
         TabIndex        =   6
         ToolTipText     =   "Razón Social del Cliente"
         Top             =   1800
         Width           =   4305
      End
      Begin VB.TextBox Text1 
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
         Index           =   5
         Left            =   5370
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3240
         Width           =   2100
      End
      Begin VB.TextBox Text1 
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
         Index           =   4
         Left            =   1620
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   14
         Top             =   3240
         Width           =   2100
      End
      Begin VB.TextBox Text104 
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
         Left            =   2340
         MaxLength       =   8
         TabIndex        =   4
         ToolTipText     =   "Nro. de Comprobante"
         Top             =   1140
         Width           =   1245
      End
      Begin VB.TextBox Text1 
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
         Left            =   1605
         MaxLength       =   5
         TabIndex        =   5
         ToolTipText     =   "Código de Cliente"
         Top             =   1800
         Width           =   1035
      End
      Begin VB.TextBox Text1 
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
         Index           =   2
         Left            =   7860
         MaxLength       =   1
         TabIndex        =   7
         ToolTipText     =   "Letra de Provincia"
         Top             =   1800
         Width           =   435
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Borrar"
         Height          =   345
         Left            =   6990
         TabIndex        =   23
         ToolTipText     =   "Elimina definitivamente el Registro"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "CargaFC.frx":00AA
         Left            =   1560
         List            =   "CargaFC.frx":00C3
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   540
         Width           =   1665
      End
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "CargaFC.frx":0122
         Left            =   1560
         List            =   "CargaFC.frx":0132
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   675
      End
      Begin VB.TextBox Text103 
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
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "0001"
         ToolTipText     =   "Punto de Venta o Prefijo"
         Top             =   1140
         Width           =   615
      End
      Begin MSComCtl2.DTPicker FECHA 
         Height          =   330
         Left            =   1560
         TabIndex        =   0
         Top             =   210
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
         CustomFormat    =   "ddd d MMM yyy"
         Format          =   16842755
         CurrentDate     =   36877
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Impte. I.B. :"
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
         Index           =   20
         Left            =   3930
         TabIndex        =   50
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "% I.B. :"
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
         Index           =   19
         Left            =   180
         TabIndex        =   49
         Top             =   4350
         Width           =   1335
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
         Index           =   18
         Left            =   180
         TabIndex        =   45
         Top             =   3795
         Width           =   1335
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   300
         X2              =   8175
         Y1              =   1605
         Y2              =   1605
      End
      Begin VB.Line Line3 
         X1              =   290
         X2              =   8175
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Motivo :"
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
         Index           =   17
         Left            =   3570
         TabIndex        =   44
         Top             =   2700
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cond. Vta. :"
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
         Index           =   16
         Left            =   450
         TabIndex        =   43
         Top             =   2670
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Index           =   15
         Left            =   4050
         TabIndex        =   42
         Top             =   2430
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendedor :"
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
         Index           =   14
         Left            =   450
         TabIndex        =   41
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "CUIT :"
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
         Index           =   13
         Left            =   3960
         TabIndex        =   40
         Top             =   2160
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cod. IVA :"
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
         Index           =   12
         Left            =   420
         TabIndex        =   39
         Top             =   2130
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tot.Unid. :"
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
         Index           =   11
         Left            =   3930
         TabIndex        =   38
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Items :"
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
         Left            =   180
         TabIndex        =   37
         Top             =   4050
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Exento :"
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
         Index           =   10
         Left            =   180
         TabIndex        =   36
         Top             =   3540
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6150
         TabIndex        =   35
         Top             =   780
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Pcia :"
         Height          =   225
         Index           =   6
         Left            =   6840
         TabIndex        =   34
         Top             =   1830
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Letra :  "
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   33
         Top             =   900
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "IVA No Insc. :"
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
         Index           =   9
         Left            =   3930
         TabIndex        =   32
         Top             =   3510
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "IVA Inscr. :"
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
         Index           =   8
         Left            =   3915
         TabIndex        =   31
         Top             =   3270
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Gravado :"
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
         Index           =   7
         Left            =   165
         TabIndex        =   30
         Top             =   3270
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro. Cpbte.:"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   29
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cliente : "
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   28
         Top             =   1860
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha :"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   27
         Top             =   255
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo Cpbte. :  "
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1425
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   2070
         TabIndex        =   25
         Top             =   600
         Width           =   2715
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   8220
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   8220
         Y1              =   3075
         Y2              =   3075
      End
   End
End
Attribute VB_Name = "CargaFC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Mensaje As String

Dim dbFCIVA As New ADODB.Connection, RgFCIVA As New ADODB.Recordset
Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset

Private Sub Combo1_Click()

Blanquea

If Mid(Combo1.Text, 1, 1) = 1 Or Mid(Combo1.Text, 1, 1) = 4 Then
    Label1(17).Caption = "F. Ped.:"
    Combo7.Visible = False
    Combo8.Visible = True
Else
    Label1(17).Caption = "Motivo :"
    Combo7.Visible = True
    Combo8.Visible = False
End If

End Sub

Private Sub FECHA_GotFocus()

Blanquea

End Sub

Private Sub Form_Load()
Dim I1%

Me.Move 0, 0, 9300, 5700

dbFCIVA.ConnectionString = BDatos1
dbFCIVA.Open

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

CargaVend
CargaZona

CargaCVta
CargaMoti

Combo1.ListIndex = 0
Combo2.ListIndex = 0
Text103.Text = "0001"

FECHA.Value = Date

Command2.Enabled = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

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

Private Sub Form_Unload(Cancel As Integer)

dbFCIVA.Close
dbCLTE.Close

End Sub

Private Sub Command1_Click()

'   Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Command2_Click()
 
  vbMsgBoxTitle = " Carga de Datos a la Tabla "
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

'-------------------------------------------------------------

miSQL2 = "SELECT * FROM FCIVAVta WHERE LETRA = '" & _
         Combo2.List(Combo2.ListIndex) & "' AND TIPO = '" & _
         Mid(Combo1.List(Combo1.ListIndex), 1, 1) & "' AND PTOVTA = " & _
         Val(Text103.Text) & " AND CPBTE = " & Val(Text104.Text)
RgFCIVA.Open miSQL2, dbFCIVA, adOpenDynamic, adLockPessimistic

If RgFCIVA.EOF Then
    RgFCIVA.AddNew
    RgFCIVA!FECHA = FECHA.Value
    RgFCIVA!TIPO = Mid(Combo1.List(Combo1.ListIndex), 1, 1)
    RgFCIVA!Letra = Combo2.List(Combo2.ListIndex)
    RgFCIVA!PtoVta = Text103.Text
    RgFCIVA!Cpbte = Text104.Text
    RgFCIVA!totcos = 0
    RgFCIVA!GRNOINs = 0
    RgFCIVA!COMIS = 0
    RgFCIVA!NOIMPR = 0
    End If
    
    RgFCIVA!CLTE = Text1(0).Text
    RgFCIVA!Nomb = Text1(1).Text
    RgFCIVA!PCIA = Text1(2).Text
    RgFCIVA!Cuit = Text1(3).Text
    
    RgFCIVA!GRINS = Text1(4).Text
    RgFCIVA!IVAIns = Text1(5).Text
    RgFCIVA!IVANOINS = Text1(6).Text
    RgFCIVA!EXENTO = Text1(7).Text
    RgFCIVA!bon = Text1(8).Text
    RgFCIVA!PORCIB = Text1(11).Text
    RgFCIVA!TotIB = Text1(12).Text
    
    RgFCIVA!ITEMS = Text1(9).Text
    RgFCIVA!TotCan = Text1(10).Text
    
    RgFCIVA!CIVA = Combo3.ListIndex
    RgFCIVA!VEND = Mid(Combo4.List(Combo4.ListIndex), 1, 2)
    RgFCIVA!ZONA = Mid(Combo5.List(Combo5.ListIndex), 1, 2)
    RgFCIVA!CVTA = Mid(Combo6.List(Combo6.ListIndex), 1, 2)
    
If RgFCIVA!TIPO = 1 Or RgFCIVA!TIPO = 4 Then
    RgFCIVA!moti = Mid(Combo8.List(Combo8.ListIndex), 1, 1)
Else
    RgFCIVA!moti = Mid(Combo7.List(Combo7.ListIndex), 1, 2)
End If

    RgFCIVA.Update

RgFCIVA.Close

Salir1:

FECHA.SetFocus

Salir2:
   
End Sub

Private Sub Command3_Click()

  vbMsgBoxTitle = " Eliminar Datos de la Tabla "
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

miSQL2 = "SELECT * FROM FCIVAVta WHERE LETRA = '" & _
         Combo2.List(Combo2.ListIndex) & "' AND TIPO = '" & _
         Mid(Combo1.List(Combo1.ListIndex), 1, 1) & "' AND PTOVTA = " & _
         Val(Text103.Text) & " AND CPBTE = " & _
         Val(Text104.Text)
RgFCIVA.Open miSQL2, dbFCIVA, adOpenDynamic, adLockPessimistic

If RgFCIVA.EOF Then Exit Sub

RgFCIVA.Delete

RgFCIVA.Close

Salir1:

FECHA.SetFocus

Salir2:

End Sub


Private Sub text1_GotFocus(Index As Integer)

Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)

End Sub

Private Sub text1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

If Index = 1 Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text103_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text104_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub text104_lostfocus()

'Validar
    
laLOGIK1 = FECHA.Month & "/" & FECHA.Day & "/" & FECHA.Year
miSQL2 = "SELECT * FROM FCIVAVta WHERE LETRA = '" & _
         Combo2.List(Combo2.ListIndex) & "' AND TIPO = '" & _
         Mid(Combo1.List(Combo1.ListIndex), 1, 1) & "' AND PTOVTA = " & _
         Val(Text103.Text) & " AND CPBTE = " & _
         Val(Text104.Text) & " AND FECHA >= " & Separa & _
         laLOGIK1 & Separa
RgFCIVA.Open miSQL2, dbFCIVA, adOpenForwardOnly, adLockReadOnly

If Not RgFCIVA.EOF Then
    Label4.Caption = "": Label4.Visible = False
    Text1(0).Text = RgFCIVA!CLTE
    Text1(1).Text = RgFCIVA!Nomb
    Text1(2).Text = RgFCIVA!PCIA
    Text1(3).Text = RgFCIVA!Cuit
    Text1(4).Text = RgFCIVA!GRINS
    Text1(5).Text = RgFCIVA!IVAIns
    Text1(6).Text = RgFCIVA!IVANOINS
    Text1(7).Text = RgFCIVA!EXENTO
    Text1(8).Text = RgFCIVA!bon
    Text1(9).Text = RgFCIVA!ITEMS
    Text1(10).Text = RgFCIVA!TotCan
    
    If IsNull(RgFCIVA!PORCIB) Then
        Text1(11).Text = 0
    Else
        Text1(11).Text = RgFCIVA!PORCIB
    End If
    
    If IsNull(RgFCIVA!TotIB) Then
        Text1(12).Text = 0
    Else
        Text1(12).Text = RgFCIVA!TotIB
    End If
    
    Motivo = RgFCIVA!moti
    If Motivo = " " Or Motivo = "  " Then Motivo = " 1"
    
    Combo3.ListIndex = RgFCIVA!CIVA
    Combo4.ListIndex = BuscaEnCombo(Combo4, RgFCIVA!VEND)
    Combo5.ListIndex = BuscaEnCombo(Combo5, RgFCIVA!ZONA)
    Combo6.ListIndex = BuscaEnCombo(Combo6, RgFCIVA!CVTA)
    

    If Mid(Combo1.Text, 1, 1) = 1 Or Mid(Combo1.Text, 1, 1) = 4 Then
        Combo8.ListIndex = BuscaEnCombo(Combo8, Motivo)
    Else
        Combo7.ListIndex = BuscaEnCombo(Combo7, Motivo)
    End If
    
    Command3.Visible = True
    
    GoTo Salirse
    End If
   
    Command3.Visible = False
    
   Label4.Caption = " N U E V O": Label4.Visible = True
    
Text1(0) = 0
Text1(1) = ""
Text1(2) = 0
Text1(3) = 0
Text1(4) = 0
Text1(5) = 0
Text1(6) = 0
Text1(7) = 0
Text1(8) = 0
Text1(9) = 0
Text1(10) = 0
Text1(11) = 0
Text1(12) = 0

Salirse:
   
Text1(0).SetFocus

RgFCIVA.Close

End Sub

Private Sub Text1_LostFocus(Index As Integer)

If Index <> 0 Then GoTo AValidar
    
miSQL2 = "SELECT * FROM Clientes WHERE CODIGO = " & Val(Text1(Index).Text)
RgCLTE.Open miSQL2, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgCLTE.EOF Then
  vbMsgBoxTitle = "Error de Ingreso"
  vbMsgBoxText = "  " & vbCrLf & "Cliente  NO  Existe  "
  vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
   MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle
   RgCLTE.Close
    Text1(1).SetFocus
   Exit Sub
    End If
    
Text1(1).Text = RgCLTE!Nomb
Text1(2).Text = RgCLTE!PCIA
Text1(3).Text = RgCLTE!Cuit
Combo3.ListIndex = RgCLTE!CIVA
Combo4.ListIndex = BuscaEnCombo(Combo4, RgCLTE!VEND)
Combo5.ListIndex = BuscaEnCombo(Combo5, RgCLTE!ZONA)
Combo6.ListIndex = BuscaEnCombo(Combo6, RgCLTE!CVTA)

Text1(4).SetFocus

RgCLTE.Close

AValidar:

Validar

End Sub

Private Sub CargaVend()
Dim I1

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

End Sub

Private Sub CargaZona()
Dim I1

    ' Tabla de Zona
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'ZN   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

I1 = 0

If Not RgTABL.EOF Then
        Do Until RgTABL.EOF
        RSet ClteCVta = Val(RgTABL!COD)
        Combo5.AddItem ClteCVta & "-" & RgTABL!Descri
        Combo5.ItemData(I1) = RgTABL!COD
        RgTABL.MoveNext
        I1 = I1 + 1
        Loop
End If

RgTABL.Close

End Sub

Private Sub CargaCVta()
Dim I1

    ' Tabla de Vendedores
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'CV   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

I1 = 0

If Not RgTABL.EOF Then
        Do Until RgTABL.EOF
        RSet ClteCVta = Val(RgTABL!COD)
        Combo6.AddItem ClteCVta & "-" & RgTABL!Descri
        Combo6.ItemData(I1) = RgTABL!COD
        RgTABL.MoveNext
        I1 = I1 + 1
        Loop
End If

RgTABL.Close

End Sub

Private Sub CargaMoti()
Dim I1

    ' Tabla de Vendedores
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'MT   ' ORDER BY COD"
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

End Sub
'--------------------------------------------
'    Rutina de Validación y consistencia
'--------------------------------------------

Private Sub Validar()

If TIPOMov <> 2 Then Command2.Enabled = False

If Text104.Text = "" Or Text1(0).Text = "" Or Text1(1).Text = "" Or _
   Text1(2).Text = "" Then
    Exit Sub
End If

Command2.Enabled = True
    
End Sub

Private Sub Blanquea()

Text103.Text = ""
Text104.Text = ""
Text1(0).Text = ""
Text1(1).Text = ""
Text1(2).Text = ""
Text1(3).Text = ""
Text1(4).Text = ""
Text1(5).Text = ""
Text1(6).Text = ""
Text1(7).Text = ""
Text1(8).Text = ""
Text1(9).Text = ""
Text1(10).Text = ""
Text1(11).Text = ""
Text1(12).Text = ""

End Sub

