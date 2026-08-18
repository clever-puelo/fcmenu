VERSION 5.00
Begin VB.Form DETALLE 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle del Artículo"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Artículo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   945
      Left            =   90
      TabIndex        =   12
      Top             =   30
      Width           =   7965
      Begin VB.Label Label5 
         Caption         =   "XXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5970
         TabIndex        =   17
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label4 
         Caption         =   "XXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4875
         TabIndex        =   16
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label3 
         Caption         =   "XXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3750
         TabIndex        =   15
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "XXXXXXXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1860
         TabIndex        =   14
         Top             =   330
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "XXXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   705
         TabIndex        =   13
         Top             =   330
         Width           =   960
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Notas"
      Height          =   300
      Left            =   4560
      TabIndex        =   11
      Top             =   2250
      Width           =   1300
   End
   Begin VB.Frame Frame3 
      Caption         =   "Stock"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1005
      Left            =   3885
      TabIndex        =   2
      Top             =   1080
      Width           =   4140
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1290
         TabIndex        =   10
         Top             =   615
         Width           =   660
      End
      Begin VB.Label Label16 
         Caption         =   "Mínimo  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   285
         TabIndex        =   9
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1230
         TabIndex        =   8
         Top             =   285
         Width           =   720
      End
      Begin VB.Label Label14 
         Caption         =   "Actual   :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   300
         TabIndex        =   7
         Top             =   285
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Precios"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   990
      Left            =   90
      TabIndex        =   1
      Top             =   1080
      Width           =   3675
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "0,00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   1575
         TabIndex        =   6
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label12 
         Caption         =   "Costo  :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   690
         TabIndex        =   5
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "0,00"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1635
         TabIndex        =   4
         Top             =   330
         Width           =   705
      End
      Begin VB.Label Label10 
         Caption         =   "Lista   :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   690
         TabIndex        =   3
         Top             =   330
         Width           =   705
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Default         =   -1  'True
      Height          =   300
      Left            =   6660
      TabIndex        =   0
      Top             =   2235
      Width           =   1300
   End
End
Attribute VB_Name = "DETALLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbART As New ADODB.Connection, RgART As New ADODB.Recordset

Private Sub Form_Load()
Dim Total As Currency

Me.Move 3010, 0, 8300, 4000

dbART.ConnectionString = BDatos1
dbART.Open

'FCMENU.Toolbar2.Buttons(2).Value = tbrPressed

If DetFact.FG1.Row < 1 Then Exit Sub

    Label1.Caption = Correa
    
    If Correa = "PL" Or Correa = "PLS" Then
        Label3.Caption = "Milim."
        Label5.Caption = "Telas"
        Label2.Caption = Format(Milim, "#0")
        Label4.Caption = Format(Telas, "#0")
        GoTo Saltodet1
        End If
    
    If Correa = "SF" Then
        Label3.Caption = "Pulgadas"
        Label5.Caption = "Metros"
        Label2.Caption = Format(Pulg, "#00.0")
        Label4.Caption = Format(Mtros, "#0.00")
        GoTo Saltodet1
       End If
    
    If Correa = "BL" Or Correa = "BLN" Then
        Label3.Caption = "Pulgadas"
        Label5.Caption = "Telas"
        Label2.Caption = Format(Pulg, "#00.0")
        Label4.Caption = Format(Telas, "#0")
        GoTo Saltodet1
        End If
            
    If Correa = "H2" Or Correa = "H4" Or Correa = "H5" Or _
       Correa = "H6" Or Correa = "H8" Or Correa = "H9" Or _
       Correa = "H10" Or Correa = "H12" Or Correa = "H15" _
            Then
        Label3.Caption = "Pulgadas"
        Label2.Caption = Format(Pulg, "#00.0")
        GoTo Saltodet1
        End If

        Label3.Caption = "Número"
        Label2.Caption = NroCor
    
Saltodet1:
            
    miSQL1 = "SELECT STOCK, PREC, PCOS, STMIN FROM ARTIculo" & _
            " WHERE COD1 = " & Correa & " AND COD2 = " & RESTO

RgART.Open miSQL1, dbART, adOpenDynamic, adLockPessimistic

If RgART.EOF Then
    RgART.Close
    Exit Sub
    End If
            
            Label11.Caption = Format(RgART!prec, "$ ###,##0.00")
            Label13.Caption = Format(RgART!PCos, "$ ###,##0.00")
            Label15.Caption = Format(RgART!STOCK, "#,##0.00")
            Label17.Caption = Format(RgART!Stmin, "#,##0.00")
     
miSQL1 = "SELECT * FROM Notartic WHERE COD1 = " & Correa & _
      " AND COD2 = " & RESTO
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic
   
If RgTABL.EOF Then
    RgTABL.Close
    NOTARTIC.Show
    End If

RgTABL.Close

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyF1 Then NOTARTIC.Show
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

'FCMENU.Toolbar2.Buttons(2).Value = tbrUnpressed
Unload NOTARTIC

dbART.Close

End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()

NOTARTIC.Show
    
End Sub


