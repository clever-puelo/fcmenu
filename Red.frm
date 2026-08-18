VERSION 5.00
Begin VB.Form RED 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Conexión en la RED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   3465
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   5
         Left            =   2160
         TabIndex        =   17
         Top             =   1350
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   4
         Left            =   900
         TabIndex        =   16
         Top             =   1350
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   3
         Left            =   2160
         TabIndex        =   15
         Top             =   1190
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   2
         Left            =   900
         TabIndex        =   14
         Top             =   1190
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   1
         Left            =   2160
         TabIndex        =   13
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Index           =   0
         Left            =   900
         TabIndex        =   12
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   900
         TabIndex        =   11
         Top             =   860
         Width           =   1005
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   900
         TabIndex        =   10
         Top             =   690
         Width           =   1005
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   900
         TabIndex        =   9
         Top             =   520
         Width           =   1005
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   900
         TabIndex        =   8
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   900
         TabIndex        =   7
         Top             =   190
         Width           =   1005
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "S.Op.         :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   30
         TabIndex        =   6
         Top             =   195
         Width           =   850
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario     :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   30
         TabIndex        =   5
         Top             =   360
         Width           =   850
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nomb.Máq.:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   30
         TabIndex        =   4
         Top             =   520
         Width           =   850
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor    :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   30
         TabIndex        =   3
         Top             =   690
         Width           =   850
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Dominio    :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   30
         TabIndex        =   2
         Top             =   860
         Width           =   850
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Grupos Aut:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   30
         TabIndex        =   1
         Top             =   1020
         Width           =   850
      End
   End
End
Attribute VB_Name = "RED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_Load()
Me.Move 8200, 6800, 3510, 1450

Label3.Caption = VNet.OSPlataf
Label4.Caption = VNet.UserName
Label5.Caption = VNet.ComputerName
Label6.Caption = VNet.Server
Label7.Caption = VNet.LogonDomain
For i = 0 To VNet.GroupCount
Label8(i).Caption = VNet.Group(i)
Next

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   
 Unload Me

End Sub

Private Sub Frame1_Click()

    Unload Me

End Sub


