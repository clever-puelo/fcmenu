VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.MDIForm FCMENU 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "FCMENU  - Facturación y Cuenta Corriente (Electrónica)"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   14325
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1155
      Top             =   1620
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5700
      Top             =   2175
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   51
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":15D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":18EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":1C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":2056
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":2930
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":31B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":34D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":3E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":42D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":4726
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":4B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":4FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":58A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":5BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":6010
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":6880
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":7444
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":775E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":7BB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":848A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":88DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":8D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":9608
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":A45A
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":A8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":ACFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":B150
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":B5A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":B8BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":C48E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":C7A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":D37A
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":D7CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":E7D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":EAEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":EE06
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":F754
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":1002E
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":10348
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":10662
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":10AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":10F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":11358
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":11C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":11F4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":12266
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FCMENU.frx":126B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      ScaleHeight     =   720
      ScaleWidth      =   14265
      TabIndex        =   1
      Top             =   0
      Width           =   14325
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   720
         Left            =   8610
         TabIndex        =   4
         Top             =   0
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   1270
         ButtonWidth     =   1217
         ButtonHeight    =   1164
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Caption         =   "Herram."
               Key             =   "HERRAM"
               Object.ToolTipText     =   "Recicla"
               ImageIndex      =   43
               Style           =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "SALIR"
               Object.ToolTipText     =   "Termina el Programa"
               ImageIndex      =   50
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ayuda"
               Key             =   "AYUDA"
               Object.ToolTipText     =   "Ayuda"
               ImageIndex      =   40
               Style           =   5
               Object.Width           =   1e-4
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AYUCONT"
                     Text            =   "Contenido ..."
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AYUACE"
                     Text            =   "Acerca de FCMENU 1.0 ..."
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   720
         Left            =   3780
         TabIndex        =   3
         Top             =   0
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   1270
         ButtonWidth     =   1429
         ButtonHeight    =   1164
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " A-B-M "
               Key             =   "ABMS"
               Object.ToolTipText     =   "Altas, Bajas y Modificaciones"
               ImageIndex      =   17
               Style           =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Consultas"
               Key             =   "CONSULTA"
               Object.ToolTipText     =   "Consultas y Visualizaciones"
               ImageIndex      =   15
               Style           =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ingreso"
               Key             =   "INGRESO"
               Object.ToolTipText     =   "Programas de Ingreso o Carga"
               ImageIndex      =   44
               Style           =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Listados"
               Key             =   "LIST"
               Object.ToolTipText     =   "Listados Varios"
               ImageIndex      =   48
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   720
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   1270
         ButtonWidth     =   1667
         ButtonHeight    =   1164
         Appearance      =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Calendario"
               Key             =   "CALEN"
               Object.ToolTipText     =   "Calendario"
               ImageIndex      =   31
               Style           =   2
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Calculadora"
               Key             =   "CALCU"
               Object.ToolTipText     =   "Calculadora"
               ImageIndex      =   5
               Style           =   2
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Notas"
               Key             =   "NOTAS"
               Object.ToolTipText     =   "Notas Varias"
               ImageIndex      =   8
               Style           =   5
               Object.Width           =   1e-4
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "NOTCLTE"
                     Text            =   "Clientes"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "NOTART"
                     Text            =   "Artículos"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCCIÓN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   390
         Left            =   12180
         TabIndex        =   5
         Top             =   180
         Width           =   4530
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   7170
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu HH 
      Caption         =   "Herramientas"
      Begin VB.Menu CALCU 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu CALEN 
         Caption         =   "Calendario"
      End
      Begin VB.Menu X0 
         Caption         =   "-"
      End
      Begin VB.Menu NOTA 
         Caption         =   "Notas Varias"
         Begin VB.Menu NOTCLT1 
            Caption         =   "Clientes"
         End
         Begin VB.Menu NOTART1 
            Caption         =   "Artículos"
            WindowList      =   -1  'True
         End
      End
      Begin VB.Menu SALIR 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu CC 
      Caption         =   "Tareas"
      Begin VB.Menu abm 
         Caption         =   "ABM´s"
         Begin VB.Menu ABMCLTE1 
            Caption         =   "Clientes"
         End
         Begin VB.Menu DTOXCLT1 
            Caption         =   "Descuentos del Cliente"
         End
         Begin VB.Menu ABMART1 
            Caption         =   "Artículos"
         End
         Begin VB.Menu MODPREC1 
            Caption         =   "Modif. de Precios"
         End
         Begin VB.Menu ABMTAB1 
            Caption         =   "Tablas Varias"
         End
         Begin VB.Menu ET5 
            Caption         =   "-"
         End
         Begin VB.Menu CARGAAFIP 
            Caption         =   "Constancias AFIP"
         End
         Begin VB.Menu ABMCOT 
            Caption         =   "Cotización del Dolar"
         End
         Begin VB.Menu PARAMET1 
            Caption         =   "Parametros"
         End
      End
      Begin VB.Menu X1 
         Caption         =   "-"
      End
      Begin VB.Menu CONSULTA 
         Caption         =   "Consultas por Pantalla"
         Begin VB.Menu VC 
            Caption         =   "Clientes"
            Begin VB.Menu BUSCLTE1 
               Caption         =   "Busqueda"
            End
            Begin VB.Menu VISCLTE1 
               Caption         =   "Detalle del Cliente"
            End
            Begin VB.Menu VISCCTE1 
               Caption         =   "Cuentas Ctes."
            End
            Begin VB.Menu VISEST1 
               Caption         =   "Ventas Por Artículo"
            End
         End
         Begin VB.Menu VISART1 
            Caption         =   "Artículos"
            Begin VB.Menu DETAL 
               Caption         =   "Detalle del Artículo"
            End
            Begin VB.Menu VTASART 
               Caption         =   "Ventas"
            End
            Begin VB.Menu VISSTO1 
               Caption         =   "Stock"
            End
            Begin VB.Menu VISDESPA 
               Caption         =   "Despachos"
            End
            Begin VB.Menu ESTADIS1 
               Caption         =   "Ventas por Cliente"
            End
         End
         Begin VB.Menu VERCOB1 
            Caption         =   "Cobranzas por Zona"
         End
         Begin VB.Menu VERFAC1 
            Caption         =   "Facturas Emitidas"
         End
         Begin VB.Menu TOTFAC1 
            Caption         =   "Totales de Facturación"
         End
         Begin VB.Menu VERCHEQ1 
            Caption         =   "Cheques"
         End
      End
      Begin VB.Menu A1 
         Caption         =   "-"
      End
      Begin VB.Menu INGRESO 
         Caption         =   "Ingreso"
         Begin VB.Menu INGFAC 
            Caption         =   "Factura"
         End
         Begin VB.Menu INGREC 
            Caption         =   "Recibos"
         End
         Begin VB.Menu NCINT 
            Caption         =   "N.Crédito Interna"
            Enabled         =   0   'False
         End
         Begin VB.Menu INGSTO1 
            Caption         =   "Stock"
         End
      End
      Begin VB.Menu X2 
         Caption         =   "-"
      End
      Begin VB.Menu LISTADO1 
         Caption         =   "Listados"
      End
   End
   Begin VB.Menu COMPLEM 
      Caption         =   "Complentos"
      Visible         =   0   'False
      Begin VB.Menu RECICLA1 
         Caption         =   "Arreglos"
         Begin VB.Menu ARRECCTE1 
            Caption         =   "Cuenta Corriente"
         End
         Begin VB.Menu ARRESUBDV1 
            Caption         =   "Subdiario Ventas"
         End
         Begin VB.Menu ARRECHEQ1 
            Caption         =   "Cheques"
         End
         Begin VB.Menu ARREEST1 
            Caption         =   "Estadísticas"
         End
      End
      Begin VB.Menu pasaarch1 
         Caption         =   "Pasaje de Fact.Manuales"
      End
      Begin VB.Menu cocina1 
         Caption         =   "Corrige Fechas"
      End
      Begin VB.Menu cocina2 
         Caption         =   "Arregla Estadísticas"
      End
      Begin VB.Menu FORZBACK 
         Caption         =   "Backup (Forzado)"
      End
   End
   Begin VB.Menu AYUDA 
      Caption         =   "Ayuda"
      Begin VB.Menu VERAYU 
         Caption         =   "Contenido ..."
      End
      Begin VB.Menu SOBRE1 
         Caption         =   "Acerca de  FCMENU 1.0 ..."
      End
   End
End
Attribute VB_Name = "FCMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim YaFue As Boolean
'  Any string argument automatically begins the 30-day trial.
'  Create a new email object
'Dim mailman As New ChilkatMailMan2
Public WithEvents mailman As ChilkatMailMan2
Attribute mailman.VB_VarHelpID = -1
Dim email As New ChilkatEmail2

Dim success As Long, NomArch As String

Private Sub BUSCLTE1_Click()

DeDonde = 9
'BusClte.Show
                ExplodeForm BusClte

End Sub

Private Sub coso_Click()
'PrnScr
End Sub

Private Sub cocina1_Click()

    Zcocina1.Show

End Sub

Private Sub cocina2_Click()

    Zcocina2.Show

End Sub

Private Sub ESTADIS1_Click()

    ESTADIST.Show

End Sub

Private Sub INGSTO1_Click()

Stock.Show


End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub MODPREC1_Click()

ModPrec.Show

End Sub

Private Sub Timer1_Timer()
On Error GoTo SALIR

StatusBar1.Panels.item(1).Text = Time

If Mid(str(Time), 1, 5) = HoraEmail Or FuerzaBack Then
   MandaBackup
End If

SALIR:

End Sub

Private Sub VISDESPA_Click()

    VerDesp.Show
    
End Sub

Private Sub VISSTO1_Click()

    VerStock.Show
    
End Sub

Private Sub VTASART_Click()

    VTAXART.Show

End Sub

Private Sub MDIForm_Load()
Dim NomDir As String, Raiz1 As String, Raiz2 As String, FecDir As String

Me.Move 100, 100, 11500, 8250

Set mailman = New ChilkatMailMan2

Me.Caption = Me.Caption & " - Versión " & App.Major & "." & App.Minor & "." & App.Revision & _
             Space$(10) & "(" & Params(4) & ")"

If HOMO Then
    Label6.Caption = "Homologación / Prueba"
Else
    Label6.Caption = "PRODUCCIÓN"
End If

Correa = "": RESTO = ""
YaFue = False
HoraEmail = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Hora")

'EsFecha = Separa & FECHA.Value & Separa
laLOGIK1 = Format(Year(Date), "0000") & "/" & Format(Month(Date), "00") & "/" & Format(Day(Date), "00")

miSQL1 = "SELECT * FROM Cotizacion WHERE fecha = " & Separa & laLOGIK1 & Separa
RgTABL.Open miSQL1, dbTABL, adOpenForwardOnly, adLockReadOnly

If RgTABL.EOF Then
    RgTABL.Close
    Cotizac.Show
'    Exit Sub
Else
    LaCotiz = RgTABL!dolar
    StatusBar1.Panels.item(3).Text = "  1 Dólar = " & Format(RgTABL!dolar, "$ ##,##0.00")
    RgTABL.Close
    End If
    
' Verifica si existen los directorios principales - Sino los crea
If Dir$("C:\Factu", vbDirectory) = "" Then MkDir ("C:\Factu")
If Dir$("C:\Factu\Comprob", vbDirectory) = "" Then MkDir ("C:\Factu\Comprob")
If Dir$("C:\Factu\Logs", vbDirectory) = "" Then MkDir ("C:\Factu\Logs")

Raiz1 = "C:\Factu\Comprob\": Raiz2 = "C:\Factu\Logs\"

'FecDir = Mid(Date, 7, 4) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 1, 2)
FecDir = Year(Date) & "-" & _
        Format(Month(Date), "00") & "-" & _
        Format(Day(Date), "00")

' Verifica si existen los direct. de la fecha, sino los crea
If Dir$(Raiz1 & FecDir, vbDirectory) = "" Then MkDir (Raiz1 & FecDir)
If Dir$(Raiz2 & FecDir, vbDirectory) = "" Then MkDir (Raiz2 & FecDir)

FCMENU.StatusBar1.Panels.item(4) = TipoDB & " [" & ServidorSQL & "] [" & SaleImpr & "]"

'Toolbar2.Buttons(1).Enabled = False
'Toolbar2.Buttons(3).Enabled = False
'Toolbar2.Buttons(4).Enabled = False
'Toolbar2.Buttons(5).Enabled = False

'DBEngine.DefaultType = dbUseODBC
'Set dbWORKSP = DBEngine.CreateWorkspace("NewSpace", "Clever", "", dbUseODBC)
' Set dbCONEX = dbWORKSP.OpenConnection("artic", dbDriverPrompt, False, "ODBC;DATABASE=MAESTROS;;PWD=;DSN=Art")
' Set dbRECORD = dbCONEX.OpenRecordset("SELECT * FROM Articulo")

'PieFact.Show
'DetFact.Show
'CabFact.Show
'BusClte.Show
           

End Sub

Private Sub PARAMET1_Click()
  
'  Paramet.Show
 
 ExplodeForm Paramet
  
End Sub

Private Sub pasaarch1_Click()
    
    PasaFCIVA.Show

End Sub

Private Sub RECFACT_Click()

ReiniciaFac

End Sub

Private Sub RECREC_Click()

ReiniciaRec

End Sub

Private Sub SACABCO_Click()

           
'SacaBcos.Show
End Sub

Private Sub SALIR_Click()

End

End Sub


Private Sub SOBRE1_Click()

'    Inicio.Show
                ExplodeForm Inicio

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

If Mid(StatusBar1.Panels.item(2).Text, 1, 10) <> "" Then
    Exit Sub
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 ' On Error GoTo tlbToolBar_ButtonClickErr

If Toolbar1.Buttons(1).Value = tbrPressed Then
          CALEND.Show
Else
          Unload CALEND
End If

If Toolbar1.Buttons(3).Value = tbrPressed Then
       Calculator.Show
       ' laCALCU = Shell("c:\windows\calc.exe", vbNormalFocus)
Else
      Unload Calculator
      ' AppActivate laCALCU, True
      ' SendKeys "%{F4}", True
End If

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

  Select Case ButtonMenu.Key
         Case "NOTACLTE"
                NOTCLT1_Click
          
         Case "Artículos"
                NOTART1_Click
  
  End Select
            
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
 ' On Error GoTo tlbToolBar_ButtonClickErr

  Select Case Button
      
      Case " A-B-M "
        PopupMenu abm, , 3930, 740
      
      Case "Consultas"
        PopupMenu CONSULTA, , 5060, 740
      
      Case "Ingreso"
        PopupMenu INGRESO, , 6180, 740
      
      Case "Listados"
'        PopupMenu LISTADO, , 7310, 740
'        Listados.Show
         ExplodeForm Listados
  
  End Select
  
End Sub

Private Sub Toolbar2_ButtonDropDown(ByVal Button As MSComctlLib.Button)

  Select Case Button
      
      Case " A-B-M "
        PopupMenu abm, , 3930, 740
      
      Case "Consultas"
        PopupMenu CONSULTA, , 5060, 740
      
      Case "Ingreso"
        PopupMenu INGRESO, , 6180, 740
      
      Case "Listados"
 '        Listados.Show
'        PopupMenu LISTADO, , 7310, 740
        ExplodeForm Listados
  
  End Select

End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
 ' On Error GoTo tlbToolBar_ButtonClickErr

  Select Case Button
      
      Case "Herram."
        PopupMenu COMPLEM, , 8420, 740
    
     Case "Salir"
           End
               
     Case "Ayuda"
           Inicio.Show

  End Select

End Sub

Private Sub Toolbar3_ButtonDropDown(ByVal Button As MSComctlLib.Button)

  Select Case Button
      
      Case "Herram."
        PopupMenu COMPLEM, , 8420, 740
      
      Case "Ayuda"
        PopupMenu AYUDA, , 10260, 740

  End Select

End Sub

Private Sub Toolbar22_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
             
            abm.Checked = False
            ABMCLTE1.Checked = False
            Unload ABMClte
  
Select Case ButtonMenu.Key
   
    Case "ABMCLTE"
            ABMCLTE1_Click
    Case "DTOXCLTE"
            DTOXCLT1_Click
    Case "ABMART"
           ' Busqueda.Show
            ExplodeForm Busqueda
    Case "ABMTAB"
            ABMTAB1_Click
    Case "COTIZ"
            ABMTAB1_Click

'--------------------------------
   
    Case "VISCLTE"
 '           BusClte.Show
            ExplodeForm BusClte
    Case "VISART"
  '          Busqueda.Show
            ExplodeForm Busqueda
    
    Case "VISCCTE"
            VISCCTE1_Click
            
    Case "VISFACT"
            CtaCte.Show
    
    Case "VISSTO"
            CtaCte.Show

'--------------------------------
   
    Case "INGFAC"
        INGFAC_Click
    
    Case "INGREC1"
           INGREC_Click
   
    Case "NCINT"
        NCINT_Click
            
    Case "INGSTO"
        EmiFact.Show

'--------------------------------
   
    Case "LISCCTE"
            EmiFact.Show
    Case "LISVTA"
            EmiFact.Show
    Case "LISCOBR"
            EmiFact.Show
    Case "LISIB"
            EmiFact.Show
    Case "LISCOMV"
            EmiFact.Show
    Case "LISCOMC"
            EmiFact.Show
    Case "LISDPEN"
            EmiFact.Show
    Case "LISDVEN"
            EmiFact.Show
    Case "LISFVV"
            EmiFact.Show
    Case "LISSTOV"
            EmiFact.Show
    Case "LISBSTO"
            EmiFact.Show

End Select

End Sub


Private Sub Toolbar33_ButtonClick(ByVal Button As MSComctlLib.Button)
 ' On Error GoTo tlbToolBar_ButtonClickErr
   Select Case Button.Key
               
     Case "RESEL"
           Unload PieFact
           Unload DetFact
           Unload CabFact
           Unload Busqueda
           Unload BusClte
           
            PieFact.Show
            DetFact.Show
            CabFact.Show
            BusClte.Show

     Case "SALIR"
           Unload Me
               
     Case "IMPRIME"
           EmiFact.Show
               
     Case "AYUDA"
           Inicio.Show
      
      End Select

End Sub

Private Sub ARRECCTE1_Click()
    
'    CargaCC.Show
 ExplodeForm CargaCC
 
End Sub

Private Sub ARRESUBDV1_Click()
  
'    CargaFC.Show
 ExplodeForm CargaFC
  
End Sub

Private Sub LISTADO1_Click()

'Listados.Show
 ExplodeForm Listados

End Sub

Private Sub VEREST1_Click()
  
   Toolbar2.Buttons(7).Value = tbrPressed
    
    ESTADIST.Show

End Sub

Private Sub NOTART1_Click()
   
If Correa > Space(5) Then
    NOTARTIC.Show
    End If

End Sub

Private Sub NOTCLT1_Click()
   
If CodCLTE > 0 Then
    NOTACLTE.Show
    End If

End Sub

Private Sub ABMART1_Click()

'Busqueda.Show
 ExplodeForm Busqueda

End Sub

Private Sub ABMCLTE1_Click()
            
            CodCLTE = 0
                DeDonde = 0
 '               BusClte.Show
            ExplodeForm BusClte

End Sub

Private Sub DTOXCLT1_Click()
               
            If CodCLTE = 0 Then
                DeDonde = 5
 '              BusClte.Show
                ExplodeForm BusClte
           Else
'                DtosxClte.Show
                ExplodeForm DtosxClte
            End If

End Sub

Private Sub ABMTAB1_Click()
    
'    ABMTablas.Show
                ExplodeForm ABMTablas

End Sub

Private Sub INGFAC_Click()
               
                CodCLTE = 0
                
                DeDonde = 1
                CabFact.Show
                PieFact.Show
                DetFact.Show
              '  BusClte.Show
                ExplodeForm BusClte

End Sub

Private Sub NCINT_Click()
    
    CodCLTE = 0
    NCInterna.Show
    ExplodeForm BusClte
    
End Sub

Private Sub INGREC_Click()

CodCLTE = 0
           
            DeDonde = 2
'            PieRec.Show
'            DetRec.Show
'            DetPago.Show
            CabRec.Show
'            BusClte.Show
            ExplodeForm BusClte

End Sub

Sub ReiniciaFac()

If Not CabFact.Visible Then Exit Sub

CodCLTE = 0

           Unload PieFact
           Unload DetFact
           Unload CabFact
           Unload Busqueda
           Unload BusClte
           
            PieFact.Show
            DetFact.Show
            CabFact.Show
            
            DeDonde = 1
            BusClte.Show

End Sub

Sub ReiniciaRec()

If Not CabRec.Visible Then Exit Sub

CodCLTE = 0

           Unload PieRec
           Unload DetRec
           Unload DetPago
           Unload CabRec
           Unload BusClte

VeNotaC = False

        
'            DetPago.Show
'            PieRec.Show
'            DetRec.Show
            CabRec.Show
                
                DeDonde = 2
                BusClte.Show

End Sub

Private Sub TOTFAC1_Click()
    
    TotFact.Show
    
End Sub

Private Sub VERAYU_Click()
  
  vbMsgBoxTitle = " Ayuda del FCMENU 1.0 "
  vbMsgBoxText = "  " & vbCrLf & "           En este momento,            " & vbCrLf & _
                                 "    la documentación del programa      " & vbCrLf & _
                                 "      no se encuentra disponible.      "
  vbMsgBoxResp = vbOKOnly + vbExclamation
   MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle

End Sub

Private Sub VERCHEQ1_Click()

'    VerCheq.Show
                ExplodeForm VerCheq

End Sub

Private Sub VERCOB1_Click()

 '   VerCobra.Show
                ExplodeForm VerCobra

End Sub

Private Sub VERFAC1_Click()
   
 '  VerFact.Show
 ExplodeForm VerFact
   
End Sub

Private Sub VISCCTE1_Click()
               
            If CodCLTE = 0 Then
                DeDonde = 3
  '              BusClte.Show
                ExplodeForm BusClte
            Else
      '          CtaCte.Show
                ExplodeForm CtaCte
            End If

End Sub

Private Sub VISCLTE1_Click()

            If CodCLTE = 0 Then
                DeDonde = 0
  '              BusClte.Show
                ExplodeForm BusClte
           Else
                TipoMov3 = 3
                VerDetClt = True
                ABMClte.Caption = "Detalle del Cliente "
                ABMClte.Frame1.Enabled = False
                ABMClte.Frame2.Enabled = False
                ABMClte.Frame3.Enabled = False
                ABMClte.Frame5.Enabled = False
                ABMClte.Show
            End If

End Sub

Private Sub CALCU_Click()
 
 Calculator.Show
 
End Sub

Private Sub CALEN_Click()

CALEND.Show

End Sub

Sub DETAL_Click()

'Busqueda.Show
ExplodeForm Busqueda

End Sub

Private Sub CARGAAFIP_Click()

'ConstAFIP.Show
ExplodeForm ConstAFIP

End Sub

Private Sub ABMCOT_Click()

'Cotizac.Show
ExplodeForm Cotizac

End Sub

Private Sub VISEST1_Click()
               
            If CodCLTE = 0 Then
                DeDonde = 6
        '        BusClte.Show
                ExplodeForm BusClte
            Else
  '              CpraClte.Show
                ExplodeForm CpraClte
            End If


End Sub


Private Sub MandaBackup()

FuerzaBack = False

If YaFue Or HOMO Then Exit Sub
    
    mailman.UnlockComponent "MAIL!TEAMCRD!_rAy8OuCQ253w"
    EnviaEmail.ProgressBar1.Value = 0


Me.WindowState = 1
EnviaEmail.Show

' backalestel
' crovara2

NomZIP = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "NombZIP")
ArchEmail = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "NombArch")
UsuEmail = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Usuario")
PassEmail = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Pass")

'success = mailman.UnlockComponent("30-day trial")
'If (success <> 1) Then
'    MsgBox "Component unlock failed"
'    Exit Sub
'End If

Timer1.Interval = 200
StatusBar1.Panels.item(1).Text = "Enviando backup ... Espere"

'  GMail SMTP server
mailman.SmtpHost = "smtp.gmail.com"
'mailman.SmtpPort = 587
'mailman.StartTLS = 1
mailman.SmtpPort = 465
mailman.SmtpSsl = 1

'  Set SMTP login/password.
mailman.SmtpUsername = UsuEmail & "@gmail.com"
mailman.SmtpPassword = PassEmail

YaFue = True

'If Not mailman.VerifySmtpLogin Then GoTo sale

email.subject = "Backup del " & Date
email.Body = "Backup del " & Date & vbCrLf & "Hora de realización: " & Time
email.From = "backup <backalestel@gmail.com>"
email.AddTo "backup", UsuEmail & "@gmail.com"
email.AddFileAttachment ArchEmail

NomArch = NomZIP & " " & Mid(Date, 1, 2) & "-" & Mid(Date, 4, 2) & "-" & Mid(Date, 7, 4)
email.ZipAttachments NomArch & ".zip"

success = mailman.SendEmail(email)

Timer1.Interval = 1000

Sale:

Unload EnviaEmail
Me.WindowState = 2

If (success <> 1) Then
    MsgBox mailman.LastErrorText
    Exit Sub
End If

Err = SetValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Realizado", Now, REG_SZ)

End Sub
' This is the event callback that updates the progress bar with the percentage done.
Private Sub mailman_SendPercentDone(ByVal percentDone As Long, abort As Long)
    EnviaEmail.ProgressBar1.Value = percentDone
End Sub

Sub mailman_AbortCheck(abort As Long)
    ' Set abort = 1 if you wish to abort.
    If abort = 1 Then
        YaFue = False
        mailman.CloseSmtpConnection
    End If
    
    EnviaEmail.Text1.Text = EnviaEmail.Text1.Text & vbCrLf & mailman.SmtpSessionLog
    EnviaEmail.Text1.SelStart = Len(EnviaEmail.Text1.Text)
    EnviaEmail.Show
    DoEvents
    
End Sub

Private Sub mailman_ReadPercentDone(ByVal percentDone As Long, abort As Long)
    ' Set abort = 1 if you wish to abort.
End Sub

Private Sub FORZBACK_Click()

MandaBackup

End Sub
