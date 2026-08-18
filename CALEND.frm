VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form CALEND 
   Caption         =   "Calendario"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2415
   ScaleWidth      =   3975
   Begin MSACAL.Calendar Calendar1 
      Height          =   2220
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   3435
      _Version        =   524288
      _ExtentX        =   6059
      _ExtentY        =   3916
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2000
      Month           =   12
      Day             =   12
      DayLength       =   1
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   2
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "CALEND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Calendar1_LostFocus()

FCMENU.Toolbar1.Buttons(1).Value = tbrUnpressed
Unload Me
End Sub

Private Sub Form_Load()

' SELEC.Width = 4095
' SELEC.Height = 2820
Me.Move 0, 3550, 4095, 2820
Calendar1.Value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)

FCMENU.Toolbar1.Buttons(1).Value = tbrUnpressed

End Sub
