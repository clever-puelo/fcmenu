VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Actualiza 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambios a Clientes"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "  Varios  "
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
      Height          =   1680
      Left            =   5580
      TabIndex        =   48
      Top             =   2745
      Width           =   2490
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "Etiquetas  :"
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
         Left            =   315
         TabIndex        =   52
         ToolTipText     =   "Es Artículo Hijo"
         Top             =   1260
         Width           =   1905
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "Lista         :"
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
         Left            =   315
         TabIndex        =   51
         ToolTipText     =   "Es Artículo Hijo"
         Top             =   945
         Width           =   1905
      End
      Begin VB.CheckBox Check2 
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
         Left            =   315
         TabIndex        =   50
         ToolTipText     =   "Es Artículo Hijo"
         Top             =   315
         Width           =   1905
      End
      Begin VB.CheckBox Check3 
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
         Left            =   315
         TabIndex        =   49
         ToolTipText     =   "Es Artículo Hijo"
         Top             =   630
         Width           =   1905
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Buscar"
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
      Left            =   2160
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5700
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Notas"
      Height          =   420
      Left            =   120
      TabIndex        =   6
      Top             =   5700
      Width           =   1300
   End
   Begin VB.Frame Frame5 
      Caption         =   "   Datos de Cobranzas  "
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
      Height          =   1125
      Left            =   90
      TabIndex        =   11
      Top             =   4455
      Width           =   8010
      Begin VB.ListBox List6 
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
         Height          =   255
         ItemData        =   "Actualiza.frx":0000
         Left            =   4590
         List            =   "Actualiza.frx":0002
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   63
         ToolTipText     =   "Calidad"
         Top             =   780
         Width           =   1755
      End
      Begin VB.ListBox List5 
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
         Height          =   255
         ItemData        =   "Actualiza.frx":0004
         Left            =   4590
         List            =   "Actualiza.frx":0006
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   61
         ToolTipText     =   "Calidad"
         Top             =   510
         Width           =   1755
      End
      Begin VB.ListBox List4 
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
         Height          =   255
         ItemData        =   "Actualiza.frx":0008
         Left            =   1575
         List            =   "Actualiza.frx":000A
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   59
         ToolTipText     =   "Calidad"
         Top             =   780
         Width           =   1755
      End
      Begin VB.ListBox List3 
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
         Height          =   255
         ItemData        =   "Actualiza.frx":000C
         Left            =   1575
         List            =   "Actualiza.frx":000E
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   57
         ToolTipText     =   "Calidad"
         Top             =   510
         Width           =   1755
      End
      Begin VB.TextBox Text301 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         MaxLength       =   3
         TabIndex        =   4
         ToolTipText     =   "Unidad de Medida de Compra"
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label3 
         Caption         =   "Cobrad.  :"
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
         Left            =   3735
         TabIndex        =   64
         Top             =   795
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Zona       :"
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
         Left            =   3735
         TabIndex        =   62
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label303 
         Caption         =   "Vendededor        :"
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
         TabIndex        =   60
         Top             =   795
         Width           =   1425
      End
      Begin VB.Label Label302 
         Caption         =   "Cond. Venta        :"
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
         TabIndex        =   58
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
         TabIndex        =   16
         Top             =   240
         Width           =   1875
      End
      Begin VB.Label Label301 
         Caption         =   "Crédito           :"
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
         TabIndex        =   13
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   420
      Left            =   6690
      TabIndex        =   7
      Top             =   5700
      Width           =   1300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   420
      Left            =   5250
      TabIndex        =   5
      Top             =   5700
      Width           =   1300
   End
   Begin VB.Frame Frame3 
      Caption         =   "   Datos de Ventas"
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
      Height          =   1680
      Left            =   90
      TabIndex        =   10
      Top             =   2745
      Width           =   5445
      Begin VB.TextBox Text203b 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         LinkTimeout     =   5
         MaxLength       =   30
         TabIndex        =   65
         ToolTipText     =   "Número Interno General"
         Top             =   765
         Width           =   2820
      End
      Begin VB.ListBox List2 
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
         Height          =   255
         ItemData        =   "Actualiza.frx":0010
         Left            =   1575
         List            =   "Actualiza.frx":0026
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   54
         ToolTipText     =   "Calidad"
         Top             =   1305
         Width           =   1755
      End
      Begin VB.TextBox Text211 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4635
         LinkTimeout     =   5
         MaxLength       =   2
         TabIndex        =   53
         ToolTipText     =   "Número Interno General"
         Top             =   1305
         Width           =   705
      End
      Begin VB.TextBox Text204 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4185
         LinkTimeout     =   5
         MaxLength       =   5
         TabIndex        =   47
         ToolTipText     =   "Número Interno General"
         Top             =   1035
         Width           =   570
      End
      Begin VB.TextBox Text204 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3600
         LinkTimeout     =   5
         MaxLength       =   5
         TabIndex        =   46
         ToolTipText     =   "Número Interno General"
         Top             =   1035
         Width           =   570
      End
      Begin VB.TextBox Text204 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2745
         LinkTimeout     =   5
         MaxLength       =   5
         TabIndex        =   45
         ToolTipText     =   "Número Interno General"
         Top             =   1035
         Width           =   570
      End
      Begin VB.TextBox Text204 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2160
         LinkTimeout     =   5
         MaxLength       =   5
         TabIndex        =   44
         ToolTipText     =   "Número Interno General"
         Top             =   1035
         Width           =   570
      End
      Begin VB.TextBox Text204 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   5
         TabIndex        =   43
         ToolTipText     =   "Número Interno General"
         Top             =   1035
         Width           =   570
      End
      Begin VB.TextBox Text203a 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   3
         TabIndex        =   37
         ToolTipText     =   "Número Interno General"
         Top             =   765
         Width           =   885
      End
      Begin VB.TextBox Text202 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         MaxLength       =   40
         TabIndex        =   34
         ToolTipText     =   "Calificación  del Artículo"
         Top             =   495
         Width           =   3780
      End
      Begin VB.TextBox Text201 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         MaxLength       =   40
         TabIndex        =   3
         ToolTipText     =   "Calificación  del Artículo"
         Top             =   225
         Width           =   3780
      End
      Begin VB.Label Label210 
         Caption         =   "Categoría            :"
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
         TabIndex        =   56
         Top             =   1350
         Width           =   1425
      End
      Begin VB.Label Label211 
         Caption         =   "Copias    :"
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
         Left            =   3735
         TabIndex        =   55
         Top             =   1350
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Desctos./Rgos.  :"
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
         Top             =   1080
         Width           =   1425
      End
      Begin VB.Label Label203 
         Caption         =   "Transporte         :"
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
         TabIndex        =   36
         Top             =   810
         Width           =   1425
      End
      Begin VB.Label Label202 
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
         TabIndex        =   35
         ToolTipText     =   "Calificación  del Artículo"
         Top             =   540
         Width           =   1425
      End
      Begin VB.Label Label201 
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
         TabIndex        =   12
         ToolTipText     =   "Calificación  del Artículo"
         Top             =   270
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Datos Nominales  "
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
      Height          =   2745
      Left            =   90
      TabIndex        =   9
      Top             =   0
      Width           =   7965
      Begin VB.TextBox Text110 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4635
         MaxLength       =   20
         TabIndex        =   41
         ToolTipText     =   "Número CUIT/CUIL/DNI (sin guiones)"
         Top             =   2385
         Width           =   1935
      End
      Begin VB.ListBox List1 
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
         Height          =   255
         ItemData        =   "Actualiza.frx":0081
         Left            =   1575
         List            =   "Actualiza.frx":0097
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   39
         ToolTipText     =   "Calidad"
         Top             =   2385
         Width           =   1935
      End
      Begin VB.TextBox Text108 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1575
         MaxLength       =   20
         TabIndex        =   33
         ToolTipText     =   "Número de Despacho"
         Top             =   2115
         Width           =   1935
      End
      Begin VB.TextBox Text108 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3555
         MaxLength       =   20
         TabIndex        =   32
         ToolTipText     =   "Número de Despacho"
         Top             =   2115
         Width           =   1935
      End
      Begin VB.TextBox Text108 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5535
         MaxLength       =   20
         TabIndex        =   31
         ToolTipText     =   "Número de Despacho"
         Top             =   2115
         Width           =   1935
      End
      Begin VB.TextBox Text108 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3555
         MaxLength       =   20
         TabIndex        =   30
         ToolTipText     =   "Número de Despacho"
         Top             =   1845
         Width           =   1935
      End
      Begin VB.TextBox Text108 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5535
         MaxLength       =   20
         TabIndex        =   29
         ToolTipText     =   "Número de Despacho"
         Top             =   1845
         Width           =   1935
      End
      Begin VB.TextBox Text108 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1575
         MaxLength       =   20
         TabIndex        =   27
         ToolTipText     =   "Número de Despacho"
         Top             =   1845
         Width           =   1935
      End
      Begin VB.TextBox Text107 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4635
         LinkTimeout     =   5
         MaxLength       =   2
         TabIndex        =   25
         ToolTipText     =   "Número Interno General"
         Top             =   1575
         Width           =   525
      End
      Begin VB.TextBox Text106 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   10
         TabIndex        =   23
         ToolTipText     =   "Número Interno General"
         Top             =   1575
         Width           =   1020
      End
      Begin VB.TextBox Text105 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   21
         ToolTipText     =   "Número Interno General"
         Top             =   1305
         Width           =   4530
      End
      Begin VB.TextBox Text104 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   50
         TabIndex        =   19
         ToolTipText     =   "Número Interno General"
         Top             =   1035
         Width           =   4530
      End
      Begin VB.TextBox Text103 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         LinkTimeout     =   5
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Número Interno General"
         Top             =   765
         Width           =   1020
      End
      Begin VB.TextBox Text102 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Descripción del Artículo"
         Top             =   495
         Width           =   4515
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1575
         MaxLength       =   5
         TabIndex        =   0
         ToolTipText     =   "Prefijo -1-"
         Top             =   225
         Width           =   1035
      End
      Begin VB.Label Label110 
         Caption         =   "C.U.I.T.    :"
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
         Left            =   3735
         TabIndex        =   40
         Top             =   2430
         Width           =   975
      End
      Begin VB.Label Label109 
         Caption         =   "Condic. I.V.A.    :"
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
         TabIndex        =   38
         Top             =   2430
         Width           =   1335
      End
      Begin VB.Label Label108 
         Caption         =   "Teléfonos     :"
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
         TabIndex        =   28
         Top             =   1890
         Width           =   1395
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
         TabIndex        =   26
         Top             =   1620
         Width           =   2505
      End
      Begin VB.Label Label107a 
         Caption         =   "Provcia.  :"
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
         Left            =   3735
         TabIndex        =   24
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label Label106 
         Caption         =   "Cod. Postal        :"
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
         TabIndex        =   22
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label Label105 
         Caption         =   "Localidad           :"
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
         TabIndex        =   20
         Top             =   1350
         Width           =   1335
      End
      Begin VB.Label Label104 
         Caption         =   "Dirección           :"
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
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label103 
         Caption         =   "Clasificador       :"
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
         TabIndex        =   17
         Top             =   810
         Width           =   1335
      End
      Begin VB.Label Label102 
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
         TabIndex        =   15
         Top             =   540
         Width           =   1395
      End
      Begin VB.Label Label101 
         Caption         =   "Código           :  "
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
         TabIndex        =   14
         Top             =   270
         Width           =   1425
      End
   End
   Begin MSAdodcLib.Adodc TABLA 
      Height          =   390
      Left            =   1590
      Top             =   6030
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   30
      CommandTimeout  =   60
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ART"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ART"
      OtherAttributes =   ""
      UserName        =   "clever"
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "TABLA"
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
   Begin MSAdodcLib.Adodc CLTE 
      Height          =   330
      Left            =   3600
      Top             =   6060
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ART"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ART"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "CLTE"
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
End
Attribute VB_Name = "Actualiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Canales       As String
Dim CAMPOS      As String
Dim s$
Dim CANAL As String * 2

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF2 Then
    Busqueda.Show
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

Private Sub Form_Load()

Me.Move 0, 0, 8250, 6550

Select Case TipoMov

Case 1
    Label7.Visible = False
        Text1.Visible = True
        Text2.Visible = True
        Text3.Visible = True
        Text4.Visible = True
        Text5.Visible = True
        List1.Visible = True

Case 3
        Text1.Visible = False
        Text2.Visible = False
        Text3.Visible = False
        Text4.Visible = False
        Text5.Visible = False
        List1.Visible = False
        Label7.Visible = True

    CLTE.RecordSource = "SELECT  * FROM ARTIculo WHERE ORDEN = " & Ordenador & _
    " AND CORR = " & Correlat
    CLTE.Refresh
    If CLTE.Recordset.EOF = True Then Exit Sub

    OrdenPadre = CLTE.Recordset!PadOrd
    CorrelPadre = CLTE.Recordset!PadCor
    Text1.Text = CLTE.Recordset!PREF1
    Text2.Text = CLTE.Recordset!PREF2
    Text3.Text = CLTE.Recordset!Basico
    Text4.Text = CLTE.Recordset!Sufijo
    List1.ListIndex = CLTE.Recordset!Calid - 1
    Text5.Text = CLTE.Recordset!SUBCAL
    Label7.Caption = CLTE.Recordset!NROFORD
    
    Canales = CLTE.Recordset!Canales
    Label1.Caption = Format(CLTE.Recordset!ORDEN, "00000000")
    Label2.Caption = Format(CLTE.Recordset!CORR, "0000")
    
    Text7.Text = CLTE.Recordset!descri
    Text13.Text = CLTE.Recordset!NroInt
    Text12.Text = CLTE.Recordset!PROV
    Text6.Text = CLTE.Recordset!ProvAlt
    
    Text8.Text = CLTE.Recordset!CALIF
    Text9.Text = CLTE.Recordset!ORIG
    Text10.Text = CLTE.Recordset!NRODESP1
    Text11.Text = CLTE.Recordset!NRODESP2
    
    Text19.Text = CLTE.Recordset!UMCPRA
    Text20.Text = CLTE.Recordset!UMVTA
    Combo1.Text = CLTE.Recordset!CABEC
           
    TABLA2.RecordSource = "SELECT * FROM Fctabla2 WHERE ORDENH = " & Ordenador _
                                        & " ORDER BY DESCAB"
    TABLA2.Refresh
     If Not TABLA2.Recordset.EOF Then
            Set Combo1.DataSource = TABLA2
            Set Combo1.RowSource = TABLA2
            Combo1.ListField = "DESCAB"
'            Combo1.BoundColumn = "CORRH"
            Combo1.Refresh
    End If

    GRILLA1.PROV.RecordSource = "SELECT * FROM Proveed WHERE CODIGO LIKE '%" & CLTE.Recordset!PROV & "%'"
    GRILLA1.PROV.Refresh
    If GRILLA1.PROV.Recordset.EOF = False Then
            Label3.Caption = GRILLA1.PROV.Recordset!RSOCIAL
    End If

    GRILLA1.PROV.RecordSource = "SELECT * FROM Proveed WHERE CODIGO LIKE '%" & CLTE.Recordset!ProvAlt & "%'"
    GRILLA1.PROV.Refresh
    If GRILLA1.PROV.Recordset.EOF = False Then
            Label11.Caption = GRILLA1.PROV.Recordset!RSOCIAL
    End If
     
     GRILLA1.STOCK.RecordSource = "SELECT * FROM STOCK WHERE sORDEN = " & Ordenador & _
      " AND sCORR = " & Correlat
    GRILLA1.STOCK.Refresh
    If GRILLA1.STOCK.Recordset.EOF = False Then
            Text16.Text = GRILLA1.STOCK.Recordset!STMAX
            Text17.Text = GRILLA1.STOCK.Recordset!STMIN
            Text18.Text = GRILLA1.STOCK.Recordset!STREP
   End If
        
    ' Tabla de Origenes
TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'B  ' AND COD = " & _
                                    CLTE.Recordset!ORIG
TABLA.Refresh
     If TABLA.Recordset.EOF = False Then
            Label5.Caption = TABLA.Recordset!descri
     Else
             Label5.Caption = ""
     End If
       
    ' Tabla de Calificación
TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'C  ' AND COD = " & _
                                    CLTE.Recordset!CALIF
TABLA.Refresh
     If TABLA.Recordset.EOF = False Then
            Label4.Caption = TABLA.Recordset!descri
     Else
             Label4.Caption = ""
     End If
       
    ' Tabla de Unid.Medida Compra
TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'E  ' AND COD = " & _
                                    CLTE.Recordset!UMCPRA
TABLA.Refresh
      If TABLA.Recordset.EOF = False Then
            Label9.Caption = TABLA.Recordset!descri
     Else
             Label9.Caption = ""
      End If
        
    ' Tabla de Unid.Medida Venta
TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'E  ' AND COD = " & _
                                    CLTE.Recordset!UMVTA
TABLA.Refresh
      If TABLA.Recordset.EOF = False Then
            Label10.Caption = TABLA.Recordset!descri
     Else
             Label10.Caption = ""
      End If
      
    ' Código del Padre
    Check1.Value = 0
    If OrdenPadre > 0 Then
        TABLA.RecordSource = "SELECT * FROM Articulo WHERE ORDEN = " & OrdenPadre _
                                        & " AND CORR = " & CorrelPadre
        TABLA.Refresh
        If TABLA.Recordset.EOF = False Then
            Check1.Value = 1
            Text21.Text = TABLA.Recordset!PREF1
            Text22.Text = TABLA.Recordset!PREF2
            Text23.Text = TABLA.Recordset!Basico
            Text24.Text = TABLA.Recordset!Sufijo
            List2.ListIndex = Val(TABLA.Recordset!Calid) - 1
            Text25.Text = TABLA.Recordset!SUBCAL
            CodPadre = TABLA.Recordset!NROFORD
            Label12.Caption = TABLA.Recordset!NROFORD
            Pic2.Visible = True
        End If
    End If
       
End Select

CargaCanal

TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'A  '" _
                            & " ORDER BY DESCRI"
TABLA.Refresh

s = FGCan.BuildComboList(TABLA.Recordset, "COD,*DESCRI", "descri", vbCyan)
FGCan.ColComboList(1) = s

End Sub

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

Grabacion
Unload Me

GRILLA1.DoVer
GRILLA1.FG1.TopRow = TopRow
GRILLA1.FG1.Row = LaFila

End Sub

Private Sub Command3_Click()

NOTCLTE.Show

End Sub

Private Sub Command4_Click()

Busqueda.Show

End Sub

Private Sub Check1_Click()

If Check1.Value = 1 Then
    Pic2.Visible = True
    Command4.Visible = True
    If Actualiza.Visible Then
        Pic2.Visible = False
        Pic1.Visible = True
        Text21.SetFocus
    End If
Else
    Pic1.Visible = False
    Pic2.Visible = False
    Command4.Visible = False
    End If
    
End Sub

Private Sub FGCan_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)

FinishEdit = True

End Sub

Private Sub FGCan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim Lin As Single, Cont As String, HayVac As Boolean

For Lin = 1 To FGCan.Rows - 1

Cont = FGCan.TextMatrix(Lin, 0)
If Cont = "" Or Cont = " " Or Val(Cont) = 0 Then
       HayVac = True
       Exit For
        End If
        
Next Lin

' Si no hay lineas en blanco, agrega una
If Not HayVac Then FGCan.Rows = FGCan.Rows + 1

End Sub

Private Sub FGCan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim Lin As Single

If FGCan.EditText = "" Then Exit Sub

        Select Case Col
        
            Case 0
                     Lin = FGCan.FindRow(FGCan.EditText, 0, Col)
            
            Case 1
                     Lin = FGCan.FindRow(FGCan.EditText, 1, Col, , True)
        End Select

' Verifica si ya fue cargado. Si es asi, reingreso
If Lin <> -1 And Lin <> Row Then
    Cancel = True
    Exit Sub
End If
                
' Verifica si la opcion existe en la tabla. Si no, reingreso
      If Col = 0 Then
                miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'A  ' AND COD = '" _
                                                & FGCan.EditText & "'"
                TABLA.RecordSource = miSQL2
                TABLA.Refresh
                If Not TABLA.Recordset.EOF Then
                    FGCan.TextMatrix(FGCan.Row, 1) = TABLA.Recordset!descri
      '              FGCan.Select Row, Col
                Else
                    Cancel = True
                    End If
        End If
                
' Verifica si la opcion existe en la tabla. Si no, reingreso
      If Col = 1 Then
                miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'A  ' AND DESCRI = '" _
                                                & FGCan.EditText & "'"
                TABLA.RecordSource = miSQL2
                TABLA.Refresh
                If Not TABLA.Recordset.EOF Then
                    FGCan.TextMatrix(FGCan.Row, 0) = TABLA.Recordset!cod
   '                 FGCan.Select Row, Col
                Else
                    Cancel = True
                    End If
        End If

' Se posiciona en la última linea
If Cancel = False Then FGCan.Row = FGCan.Rows - 1: FGCan.Col = 0

End Sub

Private Sub FGCan_GotFocus()

FGCan.BackColorFixed = &HE0E0E0
FGCan.ForeColorFixed = &HC00000

End Sub

Private Sub FGCan_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete Then
    FGCan.RemoveItem (FGCan.Row)
End If

End Sub

Private Sub FGCan_LostFocus()
FGCan.BackColorFixed = &H8000000F
FGCan.ForeColorFixed = &H80000012
End Sub

Private Sub Label12_Click()

BuscaelPadre
Text21.Text = TABLA.Recordset!PREF1
Text22.Text = TABLA.Recordset!PREF2
Text23.Text = TABLA.Recordset!Basico
Text24.Text = TABLA.Recordset!Sufijo
Text25.Text = TABLA.Recordset!SUBCAL
List2.ListIndex = TABLA.Recordset!Calid - 1

Pic2.Visible = False
Pic1.Visible = True

Text21.SetFocus

End Sub


Private Sub Frame5_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label1_Click()

End Sub

'*************************************
'   Ingresa Número FORD
'*************************************

Private Sub Text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text3_GotFocus()

Text3.SelStart = 0
Text3.SelLength = Len(Text3.Text)

End Sub

Private Sub Text3_LostFocus()
Dim i1%
    
 Respuesta = "        "
For i1 = Len(Text3.Text) To 1 Step -1

If Mid(Text3.Text, i1, 1) < 0 Or Mid(Text3.Text, i1, 1) > 9 Then GoTo AlNext

Mid(Respuesta, i1, 1) = Mid(Text3.Text, i1, 1)
AlNext:
Next

Ordenador = Val(Respuesta)
Label1.Caption = Format(Ordenador, "#######0")

OtraVez:

Correlat = Int((9999 - 2) * Rnd(Ordenador) + 1)
    
    TABLA.RecordSource = "SELECT * FROM Articulo WHERE ORDEN = " & Ordenador _
                                & " AND CORR = " & Correlat
    TABLA.Refresh

If Not TABLA.Recordset.EOF Then GoTo OtraVez

Label2.Caption = Format(Correlat, "###0")

Combo1.Text = ""
           
    TABLA2.RecordSource = "SELECT * FROM Fctabla2 WHERE ORDENH = " & Ordenador _
                                        & " ORDER BY DESCAB"
    TABLA2.Refresh
     If Not TABLA2.Recordset.EOF Then
            Set Combo1.DataSource = TABLA2
            Set Combo1.RowSource = TABLA2
            Combo1.ListField = "DESCAB"
            Combo1.Refresh
            Combo1.Text = TABLA2.Recordset!descab
    End If

End Sub

Private Sub Text4_GotFocus()

Text4.SelStart = 0
Text4.SelLength = Len(Text4.Text)

End Sub

Private Sub Text5_GotFocus()

Text5.SelStart = 0
Text5.SelLength = Len(Text5.Text)

End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode <> vbKeyReturn Then Exit Sub

If Len(Text1.Text) < 2 Then
    Prefijo1 = Space(2 - Len(Text1.Text)) & Text1.Text
    Else
    Prefijo1 = Text1.Text
End If

If Len(Text2.Text) < 4 Then
    Prefijo2 = Space(4 - Len(Text2.Text)) & Text2.Text
    Else
    Prefijo2 = Text2.Text
End If

If Len(Text3.Text) < 8 Then
    Basico = Space(8 - Len(Text3.Text)) & Text3.Text
    Else
    Basico = Text3.Text
End If

If Len(Text4.Text) < 6 Then
    Sufijo = Text4.Text & Space(6 - Len(Text4.Text))
    Else
    Sufijo = Text4.Text
End If

If Len(Text5.Text) < 6 Then
    SubCalidad = Text5.Text & Space(6 - Len(Text5.Text))
    Else
    SubCalidad = Text5.Text
End If

    NROFORD = Prefijo1 & "-" & Prefijo2 & "-" & Basico & "-" & Sufijo & "-" & Mid(List1.Text, 3, 3) & "-" & SubCalidad
    
    TABLA.RecordSource = "SELECT * FROM Articulo WHERE NROFORD = '" & NROFORD & "'"
    TABLA.Refresh

If Not TABLA.Recordset.EOF Then
  vbMsgBoxTitle = "Error de Ingreso"
  vbMsgBoxText = "  " & vbCrLf & "Código  YA  Existe  "
  vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
   MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle
   Text1.SetFocus
   Exit Sub
    End If

End Sub


'*************************************
'   Ingresa Descripcón
'*************************************

Private Sub Text7_GotFocus()

Text7.SelStart = 0
Text7.SelLength = Len(Text7.Text)

End Sub

'*************************************
'   Ingresa Número Interno
'*************************************

Private Sub Text13_GotFocus()

Text13.SelStart = 0
Text13.SelLength = Len(Text13.Text)

End Sub

'*************************************
'   Ingresa Proveedor Habitual
'*************************************

Private Sub Text12_GotFocus()

Text12.SelStart = 0
Text12.SelLength = Len(Text12.Text)

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text12_LostFocus()

    GRILLA1.PROV.RecordSource = "SELECT * FROM Proveed WHERE CODIGO = " & Val(Text12.Text)
    GRILLA1.PROV.Refresh

    If GRILLA1.PROV.Recordset.EOF Then
        Label3.Caption = "  ***   No Existe   ***"
        Label3.BackColor = &H8080FF: Label3.Alignment = 2
        MuestraError
        Text12.SetFocus
        Exit Sub
    End If
            
Label3.BackColor = &H8000000F: Label3.Alignment = 0
Label3.Caption = GRILLA1.PROV.Recordset!RSOCIAL

End Sub

'*************************************
'   Ingresa Proveedor para Precio
'*************************************

Private Sub Text6_GotFocus()

Text6.SelStart = 0
Text6.SelLength = Len(Text6.Text)

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text6_LostFocus()

    GRILLA1.PROV.RecordSource = "SELECT * FROM Proveed WHERE CODIGO = " & Val(Text6.Text)
    GRILLA1.PROV.Refresh

    If GRILLA1.PROV.Recordset.EOF Then
        Label11.Caption = "  ***   No Existe   ***"
        Label11.BackColor = &H8080FF: Label11.Alignment = 2
        MuestraError
        Text6.SetFocus
        Exit Sub
    End If
            
Label11.BackColor = &H8000000F: Label11.Alignment = 0
Label11.Caption = GRILLA1.PROV.Recordset!RSOCIAL

End Sub

'*************************************
'   Ingresa Línea del Proveedor
'*************************************

Private Sub Text14_GotFocus()

Text14.SelStart = 0
Text14.SelLength = Len(Text14.Text)

End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text14_LostFocus()

If Val(Text14.Text) = 0 Then Exit Sub

TABLA.RecordSource = "SELECT * FROM LINPROV WHERE PROV = " & Val(Text12.Text) & " AND LINEA = " & _
                                    Val(Text14.Text)
TABLA.Refresh

    If TABLA.Recordset.EOF Then
        Label16.Caption = "  ***   No Existe   ***"
        Label16.BackColor = &H8080FF: Label16.Alignment = 2
        MuestraError
        Text14.SetFocus
        Exit Sub
    End If
            
Label16.BackColor = &H8000000F: Label16.Alignment = 0
Label16.Caption = TABLA.Recordset!descri

End Sub

'*************************************
'   Ingresa Calificación
'*************************************

Private Sub Text8_GotFocus()

Text8.SelStart = 0
Text8.SelLength = Len(Text8.Text)

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text8_LostFocus()

TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'C  ' AND COD = " & _
                                    Val(Text8.Text)
TABLA.Refresh

    If TABLA.Recordset.EOF Then
        Label4.Caption = "  ***   No Existe   ***"
        Label4.BackColor = &H8080FF: Label4.Alignment = 2
        MuestraError
        Text8.SetFocus
        Exit Sub
    End If
            
Label4.BackColor = &H8000000F: Label4.Alignment = 0
Label4.Caption = TABLA.Recordset!descri

End Sub

'*************************************
'   Ingresa Origen del Artículo
'*************************************

Private Sub Text9_GotFocus()

Text9.SelStart = 0
Text9.SelLength = Len(Text9.Text)

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text9_LostFocus()

TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'B  ' AND COD = " & _
                                    Val(Text9.Text)
TABLA.Refresh

    If TABLA.Recordset.EOF Then
        Label5.Caption = "  ***   No Existe   ***"
        Label5.BackColor = &H8080FF: Label5.Alignment = 2
        MuestraError
        Text9.SetFocus
        Exit Sub
    End If
            
Label5.BackColor = &H8000000F: Label5.Alignment = 0
Label5.Caption = TABLA.Recordset!descri

End Sub

'*************************************
'   Ingresa Despacho 1
'*************************************

Private Sub Text10_GotFocus()

Text10.SelStart = 0
Text10.SelLength = Len(Text10.Text)

End Sub

'*************************************
'   Ingresa Despacho 2
'*************************************

Private Sub Text11_GotFocus()

Text11.SelStart = 0
Text11.SelLength = Len(Text11.Text)

End Sub

'*************************************
'   Ingresa Stock Punto de Pedido
'*************************************

Private Sub Text18_GotFocus()

Text18.SelStart = 0
Text18.SelLength = Len(Text18.Text)

End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'*************************************
'   Ingresa Stock Máximo
'*************************************

Private Sub Text16_GotFocus()

Text16.SelStart = 0
Text16.SelLength = Len(Text16.Text)

End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'*************************************
'   Ingresa Stock Mínimo
'*************************************

Private Sub Text17_GotFocus()

Text17.SelStart = 0
Text17.SelLength = Len(Text17.Text)

End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'*************************************
'   Ingresa Un.Med.Compra
'*************************************

Private Sub Text19_GotFocus()

Text19.SelStart = 0
Text19.SelLength = Len(Text19.Text)

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text19_LostFocus()

TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'E  ' AND COD = " & _
                                    Val(Text19.Text)
TABLA.Refresh

    If TABLA.Recordset.EOF Then
        Label9.Caption = "  ***   No Existe   ***"
        Label9.BackColor = &H8080FF: Label9.Alignment = 2
        MuestraError
        Text19.SetFocus
        Exit Sub
    End If
            
Label9.BackColor = &H8000000F: Label9.Alignment = 0
Label9.Caption = TABLA.Recordset!descri

End Sub

'*************************************
'   Ingresa Un.Med.Venta
'*************************************

Private Sub Text20_GotFocus()

Text20.SelStart = 0
Text20.SelLength = Len(Text20.Text)

End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text20_LostFocus()

TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'E  ' AND COD = " & _
                                    Val(Text20.Text)
TABLA.Refresh

    If TABLA.Recordset.EOF Then
        Label10.Caption = "  ***   No Existe   ***"
        Label10.BackColor = &H8080FF: Label10.Alignment = 2
        MuestraError
        Text20.SetFocus
        Exit Sub
    End If
            
Label10.BackColor = &H8000000F: Label10.Alignment = 0
Label10.Caption = TABLA.Recordset!descri

End Sub

'*************************************
'   Ingresa Número FORD del Padre
'*************************************

Private Sub Text21_GotFocus()

Text21.SelStart = 0
Text21.SelLength = Len(Text21.Text)

End Sub

Private Sub Text22_GotFocus()

Text22.SelStart = 0
Text22.SelLength = Len(Text22.Text)

End Sub

Private Sub Text23_GotFocus()

Text23.SelStart = 0
Text23.SelLength = Len(Text23.Text)

End Sub

Private Sub Text24_GotFocus()

Text24.SelStart = 0
Text24.SelLength = Len(Text24.Text)

End Sub

Private Sub Text25_GotFocus()

Text25.SelStart = 0
Text25.SelLength = Len(Text25.Text)

End Sub

Private Sub Text25_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode <> vbKeyReturn Then Exit Sub

If Len(Text21.Text) < 2 Then
    Prefijo1 = Space(2 - Len(Text21.Text)) & Text21.Text
    Else
    Prefijo1 = Text21.Text
End If

If Len(Text22.Text) < 4 Then
    Prefijo2 = Space(4 - Len(Text22.Text)) & Text22.Text
    Else
    Prefijo2 = Text22.Text
End If

If Len(Text23.Text) < 8 Then
    Basico = Space(8 - Len(Text23.Text)) & Text23.Text
    Else
    Basico = Text23.Text
End If

If Len(Text24.Text) < 6 Then
    Sufijo = Text24.Text & Space(6 - Len(Text24.Text))
    Else
    Sufijo = Text24.Text
End If

If Len(Text25.Text) < 6 Then
    SubCalidad = Text25.Text & Space(6 - Len(Text25.Text))
    Else
    SubCalidad = Text25.Text
End If

    CodPadre = Prefijo1 & "-" & Prefijo2 & "-" & Basico & "-" & Sufijo & "-" & Mid(List1.Text, 3, 3) & "-" & SubCalidad
    BuscaelPadre

Label12.Caption = CodPadre

Pic1.Visible = False
Pic2.Visible = True

End Sub

Sub BuscaelPadre()
    
    TABLA.RecordSource = "SELECT * FROM Articulo WHERE NROFORD = '" & CodPadre & "'"
    TABLA.Refresh

If TABLA.Recordset.EOF Then
    MuestraError
   Exit Sub
    End If

OrdenPadre = TABLA.Recordset!ORDEN
CorrelPadre = TABLA.Recordset!CORR

End Sub

Sub CargaCanal()

    Dim i1%, i2%, i3%, i4%
    
    FGCan.Cols = 2
    FGCan.Rows = 2
    
    FGCan.Row = 1
    FGCan.ColWidth(0) = 600
    FGCan.ColAlignment(0) = 9
    FGCan.ColAlignment(1) = 1
    FGCan.ColWidth(1) = 2100
    
    i3 = 1
    
    For i1 = 0 To 9
    
    i4 = (i1 * 2) + 1
    CANAL = Mid(Canales, i4, 2)
    
    If CANAL = 0 Then Exit For
     
    i3 = i3 + 1
    FGCan.Rows = i3

    FGCan.Row = i3 - 1
    FGCan.Col = 0
    FGCan.Text = CANAL
        
    If FGCan.Text <> 0 Then
        TABLA.RecordSource = "SELECT * FROM Fctabla1 WHERE CTAB = 'A  ' AND COD = " _
                            & CANAL
        TABLA.Refresh
        End If
        
    FGCan.Col = 1
            
    FGCan.Text = ""

    If CANAL <> 0 And TABLA.Recordset.EOF = False Then
        FGCan.Text = TABLA.Recordset.Fields!descri
        End If

    Next i1

FGCan.RowSel = 0
FGCan.ColSel = 1
FGCan.Col = 0
FGCan.Row = i3 - 1
FGCan.Sort = 3
FGCan.Rows = i3 + 1

End Sub

Private Sub Grabacion()
Dim Mensaje As String
Dim i1%, i2%, i3%

                Me.Hide

Select Case TipoMov
        Case 1
                Mensaje = "Desea Grabar ?"
        Case 2
                Mensaje = "Desea Eliminar ?"
        Case 3
                Mensaje = "Desea Cambiar ? "
End Select

Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "ABM de Artículos")

If Respuesta = vbNo Then Exit Sub
If Respuesta = vbCancel Then
         Exit Sub
         Me.Show
         End If

If TipoMov = 2 Then
        CLTE.Recordset.Delete
        Exit Sub
        End If
        
     
'******************************************
'           Graba en Artículos
'******************************************

CLTE.RecordSource = "SELECT  * FROM ARTIculo WHERE ORDEN = " & Ordenador & _
                                  " AND CORR = " & Correlat
CLTE.Refresh

    If CLTE.Recordset.EOF = True Then
        CLTE.Recordset.AddNew
        CLTE.Recordset!ORDEN = Ordenador
        CLTE.Recordset!CORR = Correlat
        CLTE.Recordset!NROFORD = NROFORD
        CLTE.Recordset!Basico = Basico
        CLTE.Recordset!PREF1 = Prefijo1
        CLTE.Recordset!PREF2 = Prefijo2
        CLTE.Recordset!Sufijo = Sufijo
        CLTE.Recordset!Calid = Val(Mid(List1.Text, 1, 1))
        CLTE.Recordset!SUBCAL = SubCalidad
        CLTE.Recordset!COMPRAD = 0
        CLTE.Recordset!FALTTE = 0
        CLTE.Recordset!CONVTA = 0
        CLTE.Recordset!ACOEFV = 0
        CLTE.Recordset!COSUCPR = 0
        CLTE.Recordset!CPP = 0
        CLTE.Recordset!CANDTO1 = 0
        CLTE.Recordset!CANDTO2 = 0
        CLTE.Recordset!PORDTO1 = 0
        CLTE.Recordset!PORDTO2 = 0
        CLTE.Recordset!FALTA = Date
        CLTE.Recordset!FUCPRA = 0
        CLTE.Recordset!FUSAL = 0
        CLTE.Recordset!AFMPREC = 0
        CLTE.Recordset!FFLTTE = 0
        CLTE.Recordset!MCAPED = 0
        CLTE.Recordset!NOVEND = 0
        CLTE.Recordset!MCACONJ = 0
        CLTE.Recordset!AVISO = 0
        CLTE.Recordset!MCAINV = 0
        CLTE.Recordset!MCANVO = " "
        CLTE.Recordset!FALTTE = 0
    End If

    Canales = "00000000000000000000"
    For i1 = 0 To FGCan.Rows - 3
        If Len(FGCan.TextMatrix(i1 + 1, 0)) < 2 Then
            CANAL = "0" & FGCan.TextMatrix(i1 + 1, 0)
        Else
            CANAL = FGCan.TextMatrix(i1 + 1, 0)
        End If
    Mid$(Canales, (i1 * 2) + 1, 2) = CANAL
    Next

i2 = InStr(Combo1.Text, "   ")
If i2 = 0 Then i2 = Len(Combo1.Text) + 1
'I3 = InStr(Text7.Text, "   ")

CLTE.Recordset!CABEC = Combo1.Text
CLTE.Recordset!descri = Text7.Text

CLTE.Recordset!descomp = Mid$(Combo1.Text, 1, i2 - 1) & " " & Text7.Text

CLTE.Recordset!NroInt = Text13.Text
CLTE.Recordset!PROV = Text12.Text
CLTE.Recordset!ProvAlt = Text6.Text
CLTE.Recordset!CALIF = Text8.Text
CLTE.Recordset!LINEA = Text14.Text
CLTE.Recordset!ORIG = Text9.Text
CLTE.Recordset!NRODESP1 = Text10.Text
CLTE.Recordset!NRODESP2 = Text11.Text
CLTE.Recordset!UMCPRA = Text19.Text
CLTE.Recordset!UMVTA = Text20.Text
CLTE.Recordset!Canales = Canales

    If Check1.Value = 0 Then
        CLTE.Recordset!PadOrd = 0
        CLTE.Recordset!PadCor = 0
    Else
        CLTE.Recordset!PadOrd = OrdenPadre
        CLTE.Recordset!PadCor = CorrelPadre
    End If

CLTE.Recordset!FACTUAL = Date
CLTE.Recordset!USUARIO = Mid$(Red_Usuario, 1, 6)

CLTE.Recordset.Update
     
'******************************************
'           Graba en STOCK
'******************************************

    If GRILLA1.STOCK.Recordset.EOF Then
        GRILLA1.STOCK.Recordset.AddNew
        GRILLA1.STOCK.Recordset!sORDEN = Ordenador
        GRILLA1.STOCK.Recordset!sCORR = Correlat
        GRILLA1.STOCK.Recordset!STANT = 0
        GRILLA1.STOCK.Recordset!STSUC1 = 0
        GRILLA1.STOCK.Recordset!STSUC2 = 0
        GRILLA1.STOCK.Recordset!STSUC3 = 0
        GRILLA1.STOCK.Recordset!ENTMES = 0
        GRILLA1.STOCK.Recordset!SALMES = 0
        GRILLA1.STOCK.Recordset!AJEMES = 0
        GRILLA1.STOCK.Recordset!AJSMES = 0
        GRILLA1.STOCK.Recordset!PedLaten = 0
        GRILLA1.STOCK.Recordset!PedPenC = 0
        GRILLA1.STOCK.Recordset!PEDPENV = 0
        GRILLA1.STOCK.Recordset!PENCJE = 0
    End If
    
            GRILLA1.STOCK.Recordset!STMAX = Text16.Text
            GRILLA1.STOCK.Recordset!STMIN = Text17.Text
            GRILLA1.STOCK.Recordset!STREP = Text18.Text
            GRILLA1.STOCK.Recordset!sFACTUAL = Date
            GRILLA1.STOCK.Recordset!sUSUARIO = Mid$(Red_Usuario, 1, 6)

GRILLA1.STOCK.Recordset.Update

'******************************************
'           Graba en CABECERAS
'******************************************
           
    TABLA2.RecordSource = "SELECT * FROM Fctabla2 WHERE ORDENH = " & Ordenador _
                    & " AND DESCAB = '" & Combo1.Text & "'"
    TABLA2.Refresh
     If TABLA2.Recordset.EOF Then
        TABLA2.Recordset.AddNew
        TABLA2.Recordset!ORDENH = Ordenador
        TABLA2.Recordset!CORRH = Correlat
        TABLA2.Recordset!descab = Combo1.Text
        TABLA2.Recordset!FACTUAL = Date
        TABLA2.Recordset!USUARIO = Mid$(Red_Usuario, 1, 6)
        TABLA2.Recordset.Update
   End If

End Sub

Sub MuestraError()
          
        vbMsgBoxTitle = "Error de Ingreso"
        vbMsgBoxText = "  " & vbCrLf & "Código  NO  Existe  "
        vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
        MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle

End Sub

Private Sub List4_Click()

End Sub
