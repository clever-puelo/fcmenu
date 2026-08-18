VERSION 5.00
Begin VB.Form Espera 
   BorderStyle     =   0  'None
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2205
      Left            =   0
      ScaleHeight     =   2145
      ScaleWidth      =   4935
      TabIndex        =   0
      Top             =   0
      Width           =   4995
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   90
         TabIndex        =   1
         Top             =   660
         Width           =   4755
      End
   End
End
Attribute VB_Name = "Espera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Me.Move 3000, 4000
Me.Show
' DoEvents

End Sub
