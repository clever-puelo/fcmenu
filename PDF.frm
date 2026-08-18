VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form PDF 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   10875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Configurar"
      Height          =   330
      Left            =   2355
      TabIndex        =   3
      Top             =   5445
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Imprimir"
      Height          =   330
      Left            =   345
      TabIndex        =   2
      Top             =   5445
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   330
      Left            =   9420
      TabIndex        =   1
      Top             =   5445
      Width           =   1200
   End
   Begin AcroPDFLibCtl.AcroPDF PDF1 
      Height          =   5130
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   10485
      _cx             =   5080
      _cy             =   5080
   End
End
Attribute VB_Name = "PDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If EsProvisorio Then
    EsProvisorio = False
    Unload Me
    Exit Sub
End If

    FCMENU.ReiniciaFac
    EmiFact.Command4.Enabled = False

Unload EmiFact
Unload Me

End Sub

Private Sub Command2_Click()

PDF1.printAll

Sleep 6000

If EsProvisorio Then
    Unload Me
    Exit Sub
End If

    FCMENU.ReiniciaFac
    EmiFact.Command4.Enabled = False
    Unload EmiFact

Unload Me

End Sub

Private Sub Command3_Click()

PDF1.printWithDialog

End Sub

Private Sub Form_Load()

Me.Move 1600, 1500, 11000, 6300

End Sub

Private Sub Form_Resize()

If PDF1.Width < 3800 Then PDF1.Width = 3800

PDF1.Width = Me.Width - 500
PDF1.Height = Me.Height - 1000

Command1.top = Me.Height - 800
Command2.top = Me.Height - 800
Command3.top = Me.Height - 800

Command1.Left = Me.Width - 2000

End Sub
