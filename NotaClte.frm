VERSION 5.00
Begin VB.Form NOTACLTE 
   Appearance      =   0  'Flat
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Notas del Cliente"
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6900
   DrawStyle       =   2  'Dot
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
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   4
      Top             =   2310
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4560
      TabIndex        =   3
      Top             =   2310
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5730
      TabIndex        =   2
      Top             =   2310
      Width           =   1095
   End
   Begin VB.PictureBox SP1 
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   6735
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   6795
      Begin VB.Frame Frame1 
         Caption         =   "Principal"
         Height          =   2085
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   3360
         Begin VB.TextBox Text1 
            BackColor       =   &H00DDFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   0
            Top             =   210
            Width           =   3285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Adicional"
         Height          =   2085
         Left            =   3390
         TabIndex        =   6
         Top             =   30
         Width           =   3270
         Begin VB.TextBox Text2 
            BackColor       =   &H00DDFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   1
            Top             =   210
            Width           =   3195
         End
      End
   End
End
Attribute VB_Name = "NOTACLTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Form_Deactivate()
' Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

If DetFact.Visible Then DetFact.SetFocus

End Sub

Private Sub Form_Load()

Me.Move 4300, 1600, 7000, 3100
         
miSQL2 = "SELECT * FROM Notaclte WHERE CLTE = " & CodCLTE
RgTABL.Open miSQL2, dbTABL, adOpenDynamic, adLockOptimistic

Text1.Text = "": Text2.Text = ""

If RgTABL.EOF = True Then
    RgTABL.Close
    Exit Sub
End If

If RgTABL!nota1 <> "" Then
    Text1.Text = RgTABL!nota1
    End If
If RgTABL!nota2 <> "" Then
    Text2.Text = RgTABL!nota2
    End If
    
RgTABL.Close
    
End Sub

Private Sub Command2_Click()

'  grabacion

Respuesta = MsgBox("Desea Grabar", vbYesNoCancel, "Grabación de Notas")
If Respuesta = vbNo Then Exit Sub
If Respuesta = vbCancel Then
         Exit Sub
         Load Me
         End If
         
miSQL2 = "SELECT * FROM NotaClte WHERE CLTE = " & CodCLTE
RgTABL.Open miSQL2, dbTABL, adOpenDynamic, adLockPessimistic

If RgTABL.EOF = True Then
    RgTABL.AddNew
    RgTABL!CLTE = CodCLTE
    End If

If Text1.Text = "" And Text2 = "" Then
    RgTABL.Delete
Else
    RgTABL!nota1 = Text1.Text
    RgTABL!nota2 = Text2.Text
    RgTABL.Update
End If

RgTABL.Close

End Sub

