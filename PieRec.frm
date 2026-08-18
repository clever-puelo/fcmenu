VERSION 5.00
Begin VB.Form PieRec 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Totales"
   ClientHeight    =   765
   ClientLeft      =   45
   ClientTop       =   6285
   ClientWidth     =   10680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  A Pagar  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   5430
      TabIndex        =   10
      Top             =   30
      Width           =   1620
      Begin VB.Label Label302 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   135
         TabIndex        =   11
         Top             =   270
         Width           =   1380
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "  Descuentos  "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   3630
      TabIndex        =   8
      Top             =   30
      Width           =   1620
      Begin VB.Label Label305 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   9
         Top             =   270
         Width           =   1380
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Selección"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   1950
      TabIndex        =   2
      Top             =   30
      Width           =   1620
      Begin VB.Label Label306 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   3
         Top             =   270
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Deuda Total "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   150
      TabIndex        =   0
      Top             =   15
      Width           =   1620
      Begin VB.Label Label301 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   1
         Top             =   270
         Width           =   1380
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Pago "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   7200
      TabIndex        =   4
      Top             =   30
      Width           =   1620
      Begin VB.Label Label303 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   5
         Top             =   270
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   " Diferencia "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   9000
      TabIndex        =   6
      Top             =   30
      Width           =   1620
      Begin VB.Label Label304 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   225
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Width           =   1380
      End
   End
End
Attribute VB_Name = "PieRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim IVAIns As Double, IVANI As Double, NETO As Double

Private Sub Form_Load()

Me.Move 0, 5300, 11330, 1050

Label301.Caption = 0: Label302.Caption = 0
Label303.Caption = 0: Label304.Caption = 0
Label305.Caption = 0: Label306.Caption = 0

End Sub

Private Sub Label302_Change()

If CodCLTE = 0 Or Label303.Caption = "" Then Exit Sub

Label304.Caption = Format((CCur(Label302.Caption) - CCur(Label303.Caption)), "$ ###,##0.00")

If Label302.Caption > 0 Then
            DetRec.Command2.Enabled = True
    Else
            DetRec.Command2.Enabled = False
End If
    

End Sub

Private Sub Label303_Change()

If Label302.Caption = "" Then Exit Sub

Label304.Caption = CCur(Label302.Caption) - CCur(Label303.Caption)
   

End Sub

Private Sub Label304_Change()

If Label306.Caption = "" Or Label304.Caption = "" Then Exit Sub

If Label306.Caption > 0 And Label304.Caption = 0 Then
                CabRec.Command2.Enabled = True
            Else
                CabRec.Command2.Enabled = False
End If

End Sub

Private Sub Label306_Change()

If Label306.Caption = "" Or Label304.Caption = "" Then Exit Sub

If Label306.Caption > 0 And Label304.Caption = 0 Then
                CabRec.Command2.Enabled = True
            Else
                CabRec.Command2.Enabled = False
End If

End Sub

