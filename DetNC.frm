VERSION 5.00
Begin VB.Form DetNC 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle de la Nota de Crédito / Débito"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8445
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DAFEFE&
      Enabled         =   0   'False
      Height          =   345
      Index           =   2
      Left            =   6420
      MaxLength       =   12
      TabIndex        =   5
      Top             =   3030
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DAFEFE&
      Enabled         =   0   'False
      Height          =   345
      Index           =   1
      Left            =   6450
      MaxLength       =   12
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00DAFEFE&
      Enabled         =   0   'False
      Height          =   345
      Index           =   0
      Left            =   6450
      MaxLength       =   12
      TabIndex        =   1
      Top             =   570
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SIN IVA"
      Enabled         =   0   'False
      Height          =   405
      Index           =   2
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SIN IVA"
      Enabled         =   0   'False
      Height          =   405
      Index           =   1
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2250
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "SIN IVA"
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1020
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00DAFEFE&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   2
      Left            =   390
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "DetNC.frx":0000
      Top             =   2790
      Width           =   5925
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00DAFEFE&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   1
      Left            =   390
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "DetNC.frx":0006
      Top             =   1560
      Width           =   5925
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00DAFEFE&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   390
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "DetNC.frx":000C
      Top             =   330
      Width           =   5925
   End
End
Attribute VB_Name = "DetNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

Me.Move 0, 1610, 8450, 4900
Text1(0).Text = "": Text1(1).Text = "": Text1(2).Text = ""
PieFact.Label306.Caption = 0: PieFact.Label301.Caption = 0

'EstaFactu = True
VolverArt = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)

'EstaFactu = False
VolverArt = 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Check1_Click(Index As Integer)

If Check1(Index).Value = 1 Then
    Check1(Index).BackColor = &HC0E0FF
    Check1(Index).Caption = "CON IVA"
Else
    Check1(Index).BackColor = &HC0FFC0
    Check1(Index).Caption = "SIN IVA"
End If

Text2(Index).SetFocus
PieFact.Label301.Caption = Val(PieFact.Label301.Caption) + 1

If Check1(Index).Value = 1 Then
    Importe = Importe + Val(Text2(Index).Text)
Else
    Importe = Importe - Val(Text2(Index).Text)
End If
    
PieFact.Label301.Caption = Format(Val(PieFact.Label301.Caption) - 1, "$ ###,##0.00 ")

End Sub

Private Sub Text1_Change(Index As Integer)

If Text1(Index).Text <> "" Then
    Text2(Index).Enabled = True
    Check1(Index).Enabled = True
Else
    Text2(Index).Enabled = False
    Check1(Index).Enabled = False
End If

End Sub

Private Sub Text2_Change(Index As Integer)

PieFact.Label301.Caption = Format(Val(Text2(0).Text) + Val(Text2(1).Text) + _
                            Val(Text2(2).Text), "$ ###,##0.00 ")

If Check1(Index).Value = 1 Then
    Importe = Importe + Val(Text2(Index).Text)
End If

End Sub
