VERSION 5.00
Begin VB.Form NOTARTIC 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Notas del Artículo"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6285
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
   ScaleHeight     =   3675
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   400
      Left            =   5130
      TabIndex        =   1
      Top             =   1755
      Width           =   1100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   400
      Left            =   5130
      TabIndex        =   0
      Top             =   1200
      Width           =   1100
   End
   Begin VB.PictureBox SP3 
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2115
      ScaleWidth      =   5010
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   5070
      Begin VB.Frame Frame3 
         Caption         =   "Ventas"
         Height          =   2055
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2430
         Begin VB.TextBox Text3 
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
            Height          =   2000
            Left            =   40
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   200
            Width           =   2320
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Especial"
         Height          =   2055
         Left            =   2460
         TabIndex        =   3
         Top             =   0
         Width           =   2505
         Begin VB.TextBox Text4 
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
            Height          =   2000
            Left            =   40
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   200
            Width           =   2320
         End
      End
   End
End
Attribute VB_Name = "NOTARTIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbNTAR As New ADODB.Connection, RgNTAR As New ADODB.Recordset

Private Sub Form_Load()

Me.Move 4700, 2600, 6400, 2650
' THBCoolCaption1.hWnd = Me.hWnd

dbNTAR.ConnectionString = BDatos1
dbNTAR.Open
    
miSQL2 = "SELECT * FROM Notartic WHERE COD1 = '" & Correa & _
         "' and COD2 = '" & RESTO & "'"
RgNTAR.Open miSQL2, dbTABL, adOpenDynamic, adLockPessimistic

Text3.Text = "": Text4.Text = ""

If RgNTAR.EOF Then
    RgNTAR.Close
    Exit Sub
End If

If RgNTAR!NOTA3 <> "" Then
    Text3.Text = RgNTAR!NOTA3
    End If
If RgNTAR!NOTA4 <> "" Then
    Text4.Text = RgNTAR!NOTA4
    End If
    
RgNTAR.Close
    
End Sub


Private Sub Command1_Click()

'Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Command2_Click()

'  grabacion

Respuesta = MsgBox("Desea Grabar", vbYesNoCancel, "Grabación de Notas")
If Respuesta = vbNo Then Exit Sub
If Respuesta = vbCancel Then
         Exit Sub
         End If

RgNTAR.Open miSQL2, dbTABL, adOpenDynamic, adLockPessimistic

If RgNTAR.EOF = True Then
    RgNTAR.AddNew
    RgNTAR!COD1 = Correa
    RgNTAR!Cod2 = RESTO
    End If
    
If Text3.Text = "" And Text4 = "" Then
    RgNTAR.Delete
Else
    RgNTAR!NOTA3 = Text3.Text
    RgNTAR!NOTA4 = Text4.Text
    RgNTAR.Update
End If

RgNTAR.Close

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbNTAR.Close

End Sub


