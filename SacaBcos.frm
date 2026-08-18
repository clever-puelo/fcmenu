VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form SacaBcos 
   Caption         =   "Form1"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   7635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
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
      Left            =   5370
      TabIndex        =   2
      Top             =   2880
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar"
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
      Left            =   3960
      TabIndex        =   1
      Top             =   2880
      Width           =   1230
   End
   Begin MSAdodcLib.Adodc ARTIC 
      Height          =   330
      Left            =   1185
      Top             =   3900
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
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=c1"
      OLEDBString     =   "DSN=c1"
      OLEDBFile       =   ""
      DataSourceName  =   "DATT1"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "ARTIC"
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
      Caption         =   "Saca Blancos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2685
      TabIndex        =   0
      Top             =   225
      Width           =   2355
   End
End
Attribute VB_Name = "SacaBcos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LaRow As Integer

Private Sub Command1_Click()

ARTIC.RecordSource = "SELECT  * FROM Articulo"
ARTIC.Refresh
RgART.MoveFirst

Do Until RgART.EOF

SacaLosBlancos

RgART.Update

RgART.MoveNext

Loop

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub SacaLosBlancos()
Dim val1 As Single, val2 As String

val1 = InStr(1, RgART!Cod2, " ", vbTextCompare)

val2 = Mid(RgART!Cod2, 1, val1 - 1)
RgART!Cod2 = ""
RgART!Cod2 = val2 & ""

End Sub

