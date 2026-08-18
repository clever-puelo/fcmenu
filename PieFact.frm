VERSION 5.00
Begin VB.Form PieFact 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Totales"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   6285
   ClientWidth     =   2715
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
   ScaleHeight     =   4545
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Descuento"
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
      Height          =   585
      Left            =   180
      TabIndex        =   12
      Top             =   1530
      Width           =   2400
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
         Left            =   150
         TabIndex        =   13
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Impte. Neto"
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
      Height          =   585
      Left            =   180
      TabIndex        =   8
      Top             =   3870
      Width           =   2400
      Begin VB.Label Label304 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9FEF1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "IVA No Inscr."
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
      Height          =   585
      Left            =   195
      TabIndex        =   6
      Top             =   3285
      Width           =   2400
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
         Left            =   150
         TabIndex        =   7
         Top             =   270
         Width           =   2100
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "IVA Inscr."
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
      Height          =   585
      Left            =   180
      TabIndex        =   2
      Top             =   2700
      Width           =   2400
      Begin VB.Label Label302 
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
         Left            =   150
         TabIndex        =   3
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Subtotal"
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
      Height          =   585
      Left            =   180
      TabIndex        =   0
      Top             =   2115
      Width           =   2400
      Begin VB.Label Label305 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E9FEF1&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impte. Bruto"
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
      Height          =   585
      Left            =   180
      TabIndex        =   4
      Top             =   945
      Width           =   2400
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
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   1530
      TabIndex        =   10
      Top             =   705
      Width           =   1005
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000007&
      X1              =   45
      X2              =   2790
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      X1              =   45
      X2              =   2790
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E9FEF1&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   90
      TabIndex        =   11
      Top             =   0
      Width           =   2520
   End
End
Attribute VB_Name = "PieFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SubTotal As Double, TIVAIns As Double, TIVANI As Double, NETO As Double

Private Sub Form_Load()

Me.Move 8450, 1610, 2880, 4900

End Sub

Private Sub Label301_Change()

SubTotal = Importe - CCur(Label306.Caption)
Label305.Caption = Format(SubTotal, "$ ###,##0.00")

If ClteCIVA < 3 And CltePCIA <> "V " Then
    TIVAIns = SubTotal * IVAIns
    
    If CIVA = 2 Then
        TIVANI = SubTotal * IVANI
        End If
End If
    
NETO = SubTotal + TIVAIns + TIVANI

If Motivo = 1 Then
'    Label305.Caption = Format(DetFact.FG1.Rows - 1, " ##,##0")
End If

Label302.Caption = Format(TIVAIns, "$ ###,##0.00")
Label303.Caption = Format(TIVANI, "$ ###,##0.00")
Label304.Caption = Format(NETO, "$ ###,##0.00")

End Sub

Private Sub Label304_Change()

CabFact.Command1.Enabled = True

End Sub
