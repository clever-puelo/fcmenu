VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form EnviaEmail 
   BackColor       =   &H0080FF80&
   BorderStyle     =   0  'None
   Caption         =   "EnviaEmail"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      Enabled         =   0   'False
      ForeColor       =   &H80000005&
      Height          =   2865
      Left            =   5775
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   3840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   2940
      Left            =   105
      TabIndex        =   0
      Top             =   30
      Width           =   5595
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3855
         TabIndex        =   5
         Top             =   2550
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   345
         Left            =   540
         TabIndex        =   3
         Top             =   1530
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Por favor espere ..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   570
         TabIndex        =   2
         Top             =   2055
         Width           =   4530
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Resguardo Automático de Datos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   165
         TabIndex        =   1
         Top             =   345
         Width           =   5310
      End
   End
End
Attribute VB_Name = "EnviaEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer2_Timer()
If Label2.Caption = "" Then
    Label2.Caption = "Por favor espere ..."
    Else
    Label2.Caption = ""
    End If
End Sub

Private Sub Command1_Click()
FCMENU.mailman_AbortCheck (1)
End Sub

