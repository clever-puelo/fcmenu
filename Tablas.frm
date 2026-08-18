VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Tablas 
   Caption         =   "ABMTablas - Version 1.0"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   9495
   Begin VB.Frame Frame2 
      Caption         =   "Actualización"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   750
      TabIndex        =   13
      Top             =   4980
      Visible         =   0   'False
      Width           =   4155
      Begin VB.CommandButton Command5 
         Caption         =   "Cambio (F7)"
         Height          =   345
         Left            =   2850
         TabIndex        =   16
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Baja (F6)"
         Height          =   345
         Left            =   1560
         TabIndex        =   15
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Alta (F5)"
         Height          =   345
         Left            =   240
         TabIndex        =   14
         Top             =   210
         Width           =   1100
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Retorna"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6180
      TabIndex        =   12
      Top             =   5190
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7590
      TabIndex        =   10
      Top             =   5190
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
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
      Height          =   3495
      Left            =   1530
      TabIndex        =   0
      Top             =   1110
      Visible         =   0   'False
      Width           =   6495
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Para Cálculo de Prec."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Tareas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   840
         Value           =   -1  'True
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nros. de Comprobante"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2100
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Datos Varios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2520
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Vendedores"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1260
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Condic. de Venta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   3510
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   840
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Materias Primas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Origen del Pedido"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2100
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Papel / Material"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1260
         Width           =   2500
      End
   End
   Begin VB.PictureBox FG1 
      BackColor       =   &H00FEFADE&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   1920
      ScaleHeight     =   3885
      ScaleWidth      =   5865
      TabIndex        =   17
      Top             =   870
      Visible         =   0   'False
      Width           =   5925
   End
   Begin MSAdodcLib.Adodc TABLAS 
      Height          =   330
      Left            =   150
      Top             =   5700
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   "DSN=DATT1"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DATT1"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "TABLAS"
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tablas Varias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   525
      Left            =   1560
      TabIndex        =   11
      Top             =   240
      Width           =   6525
   End
End
Attribute VB_Name = "Tablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

Me.Move 0, 0, 9600, 6400

Option1(1) = False


Frame1.Visible = True
Frame2.Visible = False
FG1.Visible = False
Command1.Move 6880
Command2.Visible = False

Habilita = False

End Sub

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

FG1.Visible = False
Frame1.Visible = True
Frame2.Visible = False
Command1.Move 6880
Command2.Visible = False
Label1.Caption = "Tablas Varias"

End Sub

Private Sub Command3_Click()
coditem = 0

TipoMov = 1
ActTabla.Show

End Sub

Private Sub Command4_Click()

coditem = FG1.TextMatrix(FG1.Row, 1)

TipoMov = 2
ActTabla.Show

End Sub

Private Sub Command5_Click()
    
 coditem = FG1.TextMatrix(FG1.Row, 1)

TipoMov = 3
ActTabla.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

SGIMenu.Toolbar2.Buttons(8).Value = tbrUnpressed

End Sub

Private Sub option1_Click(Index As Integer)

Select Case Index
     
     Case 0
            CodTabla = "CP "
            Label1.Caption = "Valores para Cálc. de Precios"
     Case 1
            Label1.Caption = "Tabla de Tareas / Trabajos"
     Case 2
            CodTabla = "MT "
            Label1.Caption = "Tabla de Papeles / Materiales"
     Case 3
            CodTabla = "OP "
            Label1.Caption = "Tabla de Origen del Pedido"
     Case 4
            CodTabla = "MP "
            Label1.Caption = "Tabla de Materias Primas"
     Case 5
            CodTabla = "CV "
            Label1.Caption = "Tabla de Condic. de Venta"
     Case 6
            CodTabla = "US "
            Label1.Caption = "Tabla de Usuarios"
     Case 7
            CodTabla = "VD "
            Label1.Caption = "Tabla de Vendedores"
     Case 8
            CodTabla = "NC "
            Label1.Caption = "Tabla de Nros. de Comprobante"
     Case 9
            CodTabla = "VS "
            Label1.Caption = "Tabla de Datos Varios"

End Select

LaOpcion = Index

Option1(Index) = False
If LaOpcion = 1 Then
TABLA.RecordSource = "SELECT PREC, cod, descri FROM Tareas ORDER BY COD"
Else
TABLA.RecordSource = "SELECT * FROM Tablas WHERE CTAB = '" & CodTabla & "' " _
                                    & "ORDER BY COD"
End If
                                    
TABLA.Refresh

If TABLA.Recordset.EOF Then
    Command3_Click
    Exit Sub
    End If

    Set FG1.DataSource = TABLA
    FG1.Visible = True
    FG1.Row = 1
    
Frame1.Visible = False
Frame2.Visible = True
Command1.Move 7580
Command2.Visible = True

End Sub

Private Sub FG1_DBLClick()

Command5_Click

End Sub
