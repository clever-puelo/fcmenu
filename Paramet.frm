VERSION 5.00
Begin VB.Form Paramet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "   Modifica Parámetros  "
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7995
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
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "  e-mail de Resguardo "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   915
      Left            =   375
      TabIndex        =   64
      Top             =   4845
      Width           =   7185
      Begin VB.TextBox Text106 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3030
         MaxLength       =   5
         TabIndex        =   24
         ToolTipText     =   "Porcentaje para Responsables Inscriptos"
         Top             =   255
         Width           =   690
      End
      Begin VB.TextBox Text106 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   23
         ToolTipText     =   "Porcentaje para Responsables Inscriptos"
         Top             =   555
         Width           =   2685
      End
      Begin VB.TextBox Text106 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   4725
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   26
         ToolTipText     =   "Porcentaje para Responsables Inscriptos"
         Top             =   540
         Width           =   2325
      End
      Begin VB.TextBox Text106 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4725
         MaxLength       =   20
         TabIndex        =   25
         ToolTipText     =   "Porcentaje para Responsables Inscriptos"
         Top             =   225
         Width           =   2370
      End
      Begin VB.TextBox Text106 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1005
         MaxLength       =   10
         TabIndex        =   22
         ToolTipText     =   "Porcentaje para Responsables Inscriptos"
         Top             =   270
         Width           =   1245
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Hora :  "
         Height          =   225
         Index           =   12
         Left            =   2250
         TabIndex        =   69
         Top             =   285
         Width           =   825
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Archivo :  "
         Height          =   225
         Index           =   11
         Left            =   -210
         TabIndex        =   68
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Pass. :  "
         Height          =   225
         Index           =   10
         Left            =   3525
         TabIndex        =   67
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuario :  "
         Height          =   225
         Index           =   9
         Left            =   3525
         TabIndex        =   66
         Top             =   285
         Width           =   1245
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Nomb. :  "
         Height          =   225
         Index           =   8
         Left            =   -195
         TabIndex        =   65
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "  Varios  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1425
      Left            =   5115
      TabIndex        =   60
      Top             =   720
      Width           =   2460
      Begin VB.OptionButton Option1 
         Caption         =   "No"
         Height          =   225
         Index           =   1
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   1110
         Width           =   540
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Si"
         Height          =   225
         Index           =   0
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   1110
         Width           =   540
      End
      Begin VB.TextBox Text105 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   10
         ToolTipText     =   "Cantidad de dias para aviso en el facturador"
         Top             =   780
         Width           =   1200
      End
      Begin VB.TextBox Text105 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   9
         ToolTipText     =   "Monto Mínimo de Deuda para Aviso en el Facturador"
         Top             =   525
         Width           =   1200
      End
      Begin VB.TextBox Text105 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   8
         ToolTipText     =   "Nombre de la Empresa"
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "C/Percep.:  "
         Height          =   225
         Index           =   15
         Left            =   60
         TabIndex        =   72
         Top             =   1110
         Width           =   1125
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Dias Vto.:  "
         Height          =   225
         Index           =   14
         Left            =   135
         TabIndex        =   63
         Top             =   780
         Width           =   1125
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Monto:  "
         Height          =   225
         Index           =   13
         Left            =   120
         TabIndex        =   62
         Top             =   525
         Width           =   1125
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa:  "
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   61
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "  Punto de Venta "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   915
      Left            =   360
      TabIndex        =   58
      Top             =   2220
      Width           =   2295
      Begin VB.TextBox Text104 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         MaxLength       =   4
         TabIndex        =   11
         ToolTipText     =   "Porcentaje para Responsables Inscriptos"
         Top             =   390
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   6240
      TabIndex        =   28
      Top             =   5850
      Width           =   1320
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   345
      Left            =   4815
      TabIndex        =   27
      Top             =   5850
      Width           =   1320
   End
   Begin VB.Frame Frame3 
      Caption         =   "  Datos  "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1635
      Left            =   360
      TabIndex        =   36
      Top             =   3180
      Width           =   7245
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5175
         MaxLength       =   20
         TabIndex        =   21
         ToolTipText     =   "Salida de Impresión (por Defecto LPT1:)"
         Top             =   1170
         Width           =   1995
      End
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   17
         ToolTipText     =   "Nombre del DSN"
         Top             =   1170
         Width           =   1875
      End
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5175
         MaxLength       =   30
         TabIndex        =   20
         ToolTipText     =   "Porcentaje para Responsables Inscriptos"
         Top             =   900
         Width           =   1995
      End
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5190
         MaxLength       =   30
         TabIndex        =   18
         ToolTipText     =   "Carpeta deonde se encuentra la Base de Datos"
         Top             =   330
         Width           =   1995
      End
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5190
         MaxLength       =   8
         TabIndex        =   19
         ToolTipText     =   "Porcentaje para Responsables Inscriptos"
         Top             =   630
         Width           =   1995
      End
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1905
         MaxLength       =   8
         TabIndex        =   16
         ToolTipText     =   "Nombre del DSN"
         Top             =   900
         Width           =   1875
      End
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   15
         ToolTipText     =   "Descripción del DSN"
         Top             =   630
         Width           =   1875
      End
      Begin VB.TextBox Text103 
         Alignment       =   2  'Center
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   14
         ToolTipText     =   "Tipo de Base de Datos : A=ACCESS - S=SQL"
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Salida Impr. :  "
         Height          =   225
         Index           =   7
         Left            =   3720
         TabIndex        =   59
         Top             =   1170
         Width           =   1545
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Servidor/Carpeta:  "
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   57
         Top             =   1170
         Width           =   1905
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Disco/Carp. :  "
         Height          =   225
         Index           =   5
         Left            =   3750
         TabIndex        =   56
         ToolTipText     =   "Nombre de La Base de Datos"
         Top             =   930
         Width           =   1515
      End
      Begin VB.Label Label103 
         Caption         =   "Desc.DSN bkp:  "
         Height          =   225
         Index           =   4
         Left            =   3780
         TabIndex        =   55
         Top             =   375
         Width           =   1455
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "DSN Backup :  "
         Height          =   225
         Index           =   3
         Left            =   3780
         TabIndex        =   54
         ToolTipText     =   "Drive o Disco donde se encuentra la Base de Datos"
         Top             =   660
         Width           =   1470
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre DSN :  "
         Height          =   225
         Index           =   2
         Left            =   420
         TabIndex        =   53
         Top             =   900
         Width           =   1545
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripción DSN :  "
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   52
         Top             =   660
         Width           =   1725
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         Caption         =   "Driver :  "
         Height          =   225
         Index           =   0
         Left            =   720
         TabIndex        =   51
         Top             =   390
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "  Porcentaje de I.V.A.  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   915
      Left            =   2670
      TabIndex        =   35
      Top             =   2220
      Width           =   4935
      Begin VB.TextBox Text102 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3780
         MaxLength       =   5
         TabIndex        =   13
         ToolTipText     =   "Porcentaje para Responsables No Inscriptos"
         Top             =   420
         Width           =   1035
      End
      Begin VB.TextBox Text102 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1260
         MaxLength       =   5
         TabIndex        =   12
         ToolTipText     =   "Porcentaje para Responsables Inscriptos"
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label Label102 
         Alignment       =   1  'Right Justify
         Caption         =   "No Inscripto :  "
         Height          =   225
         Index           =   1
         Left            =   2370
         TabIndex        =   50
         Top             =   450
         Width           =   1425
      End
      Begin VB.Label Label102 
         Alignment       =   1  'Right Justify
         Caption         =   "Inscripto :  "
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   49
         Top             =   450
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Números de Comprobante  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1425
      Left            =   360
      TabIndex        =   34
      Top             =   720
      Width           =   4740
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5970
         MaxLength       =   6
         TabIndex        =   32
         ToolTipText     =   "Último Nro. de Factura Impreso"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   5970
         MaxLength       =   6
         TabIndex        =   31
         ToolTipText     =   "Último Nro. de Factura Impreso"
         Top             =   810
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   5970
         MaxLength       =   6
         TabIndex        =   30
         ToolTipText     =   "Último Nro. de Factura Impreso"
         Top             =   540
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5970
         MaxLength       =   6
         TabIndex        =   29
         ToolTipText     =   "Último Nro. de Factura Impreso"
         Top             =   270
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3570
         MaxLength       =   6
         TabIndex        =   7
         ToolTipText     =   "Último Nro. de Factura Impreso"
         Top             =   1065
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3570
         MaxLength       =   6
         TabIndex        =   6
         ToolTipText     =   "Último Nro. de Nota de Débito ""B"" "
         Top             =   810
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3570
         MaxLength       =   6
         TabIndex        =   5
         ToolTipText     =   "Último Nro. de Nota de Crédito ""B"""
         Top             =   540
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3570
         MaxLength       =   6
         TabIndex        =   4
         ToolTipText     =   "Último Nro. de Factura ""B"" Impreso"
         Top             =   270
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   3
         ToolTipText     =   "Último Nro. de Cotización Impreso"
         Top             =   1065
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   2
         ToolTipText     =   "Último Nro. de Nota de Débito ""A"""
         Top             =   810
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   1
         ToolTipText     =   "Último Nro. de Nota de Crédito ""A"""
         Top             =   540
         Width           =   1035
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   0
         ToolTipText     =   "Último Nro. de Factura ""A"" Impreso"
         Top             =   270
         Width           =   1035
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Factura :  "
         Height          =   225
         Index           =   11
         Left            =   4800
         TabIndex        =   48
         Top             =   1110
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Factura :  "
         Height          =   225
         Index           =   10
         Left            =   4800
         TabIndex        =   47
         Top             =   840
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Factura :  "
         Height          =   225
         Index           =   9
         Left            =   4800
         TabIndex        =   46
         Top             =   570
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Factura :  "
         Height          =   225
         Index           =   8
         Left            =   4800
         TabIndex        =   45
         Top             =   300
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Recibo :  "
         Height          =   225
         Index           =   7
         Left            =   2400
         TabIndex        =   44
         Top             =   1110
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "N. Déb. B :  "
         Height          =   225
         Index           =   6
         Left            =   2400
         TabIndex        =   43
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "N.Créd. B :  "
         Height          =   225
         Index           =   5
         Left            =   2400
         TabIndex        =   42
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Factura B :  "
         Height          =   225
         Index           =   4
         Left            =   2400
         TabIndex        =   41
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Cotizac. :  "
         Height          =   225
         Index           =   3
         Left            =   30
         TabIndex        =   40
         Top             =   1110
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "N.Déb. A :  "
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   39
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "N.Créd. A :  "
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   38
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         Caption         =   "Factura A :  "
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   37
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Modificación de Parámetros"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1530
      TabIndex        =   33
      Top             =   165
      Width           =   4905
   End
End
Attribute VB_Name = "Paramet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Me.Move 500, 500, 8000, 6700

CargaDatos

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 123 And Shift = 1 Then
    Frame3.Enabled = True
    Frame6.Enabled = True
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Exit Sub

End Sub


Private Sub Command1_Click()
    
'    Unload Me

Call SlideWindow(Me, 50)
    
End Sub

Private Sub Command2_Click()
'On Error GoTo Salta

  vbMsgBoxTitle = " Actualización de Parámetros "
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

miSQL2 = "SELECT * FROM Parametro WHERE CLAVE = '1'"
RgTABL.Open miSQL2, dbTABL, adOpenDynamic, adLockPessimistic

If RgTABL.EOF Then GoTo Salta
    
    RgTABL.Fields!nume1 = Val(Text101(0).Text)
    RgTABL.Fields!nume2 = Val(Text101(1).Text)
    RgTABL.Fields!nume3 = Val(Text101(2).Text)
    RgTABL.Fields!nume4 = Val(Text101(3).Text)
    RgTABL.Fields!nume5 = Val(Text101(4).Text)
    RgTABL.Fields!nume6 = Val(Text101(5).Text)
    RgTABL.Fields!nume7 = Val(Text101(6).Text)
    RgTABL.Fields!nume8 = Val(Text101(7).Text)
    RgTABL.Fields!nume9 = Val(Text101(8).Text)
    RgTABL.Fields!nume10 = Val(Text101(9).Text)
    RgTABL.Fields!nume11 = Val(Text101(10).Text)
    RgTABL.Fields!nume12 = Val(Text101(11).Text)
    
    RgTABL.Fields!IVAIns = Val(Text102(0).Text)
    RgTABL.Fields!IVANI = Val(Text102(1).Text)
    
    RgTABL.Fields!PtoVta = Val(Text104(0).Text)
    
    RgTABL.Fields!NomEmpr = Text105(0).Text
    RgTABL.Fields!limvta = Val(Text105(1).Text)
    RgTABL.Fields!nume20 = Val(Text105(2).Text)
    
    If Option1(0).Value Then
      RgTABL.Fields!mcaib = 1
    Else
      RgTABL.Fields!mcaib = 2
    End If
    

    RgTABL.Update

Salta:

RgTABL.Close

Err = CreateNewKey(HKEY_CURRENT_USER, "FCMENU\Datos\DSN")
Err = CreateNewKey(HKEY_CURRENT_USER, "FCMENU\Datos\Impresora")
Err = CreateNewKey(HKEY_CURRENT_USER, "FCMENU\Datos\Backup")
Err = CreateNewKey(HKEY_CURRENT_USER, "FCMENU\Datos\Email")

Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos", "Driver", Text103(0).Text, REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Descrip", Text103(1).Text, REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Nombre", Text103(2).Text, REG_SZ)

If Params(3) = "A" Then
    Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Carpeta", Text103(6).Text, REG_SZ)
Else
    Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Servidor", Text103(6).Text, REG_SZ)
End If

Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Nombre", Text103(3).Text, REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Descrip", Text103(4).Text, REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Carpeta", Text103(5).Text, REG_SZ)

Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Impresora", "1", Text103(7).Text, REG_SZ)

Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "NombZIP", Text106(0).Text, REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "NombArch", Text106(1).Text, REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Usuario", Text106(2).Text, REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Pass", Text106(3).Text, REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Hora", Text106(4).Text, REG_SZ)
Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Realizado", Realizado, REG_SZ)

NomZIP = Text106(0).Text
ArchEmail = Text106(1).Text
UsuEmail = Text106(2).Text
PassEmail = Text106(3).Text
HoraEmail = Text106(4).Text

Frame6.Caption = "  e-mail de Resguardo  (" & Realizado & ")"

Salir1:

 Unload Me

Salir2:

End Sub

Private Sub CargaDatos()
Dim I1%
' On Error GoTo Salto1

    miSQL1 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
    RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockOptimistic
    
    If RgTABL.EOF Then
        RgTABL.Close
        For I1 = 0 To 12
        Text101(I1).Text = 0
        Next I1
        Text102(0).Text = 0: Text102(1).Text = 0
        GoTo Salto1
    End If
    
Text101(0).Text = RgTABL!nume1
Text101(1).Text = RgTABL!nume2
Text101(2).Text = RgTABL!nume3
Text101(3).Text = RgTABL!nume4
Text101(4).Text = RgTABL!nume5
Text101(5).Text = RgTABL!nume6
Text101(6).Text = RgTABL!nume7
Text101(7).Text = RgTABL!nume8
Text101(8).Text = RgTABL!nume9
Text101(9).Text = RgTABL!nume10
Text101(10).Text = RgTABL!nume11
Text101(11).Text = RgTABL!nume12

Text102(0).Text = RgTABL!IVAIns
Text102(1).Text = RgTABL!IVANI

Text104(0).Text = RgTABL!PtoVta

    Text105(0).Text = RgTABL.Fields!NomEmpr
    
    Text105(1).Text = RgTABL.Fields!limvta
    Text105(2).Text = RgTABL.Fields!nume20
    
    If RgTABL.Fields!mcaib = 1 Then
       Option1(0).Value = True
    Else
       Option1(1).Value = True
    End If

RgTABL.Close

Salto1:

Text103(0).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos", "Driver")
Text103(1).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Descrip")
Text103(2).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Nombre")

If Text103(0).Text = "A" Then
    Text103(6).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Carpeta")
Else
    Text103(6).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Servidor")
End If

Text103(3).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Nombre")
Text103(4).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Descrip")
Text103(5).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Backup", "Carpeta")

Text103(7).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Impresora", "1")

Text106(0).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "NombZIP")
Text106(1).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "NombArch")
Text106(2).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Usuario")
Text106(3).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Pass")
Text106(4).Text = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Hora")

NomZIP = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "NombZIP")
ArchEmail = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "NombArch")
UsuEmail = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Usuario")
PassEmail = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Pass")
HoraEmail = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Hora")
Realizado = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Realizado")

Frame6.Caption = "  e-mail de Resguardo  (" & Realizado & ")"

End Sub

