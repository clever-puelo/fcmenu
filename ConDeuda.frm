VERSION 5.00
Begin VB.Form ConDeuda 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aviso de Deuda"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   3825
      TabIndex        =   1
      Top             =   3015
      Width           =   2310
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente con Deuda Mayor a 90 días    Consulte"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1710
      Left            =   2760
      TabIndex        =   0
      Top             =   915
      Width           =   4365
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ConDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Form_Load()

Me.Move 1000, 1610, 9700, 4500

End Sub
