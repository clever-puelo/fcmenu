VERSION 5.00
Begin VB.Form CabRec 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cabecera del Recibo"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11220
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
   ScaleHeight     =   3135
   ScaleMode       =   0  'User
   ScaleWidth      =   11225.4
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Ver C.C."
      Height          =   585
      Left            =   10530
      TabIndex        =   19
      Top             =   660
      Width           =   585
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nota Clte."
      Height          =   585
      Left            =   10530
      TabIndex        =   18
      Top             =   60
      Width           =   585
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Nuevo"
      Height          =   330
      Left            =   7920
      TabIndex        =   17
      Top             =   840
      Width           =   795
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00F9FADC&
      Height          =   315
      ItemData        =   "CabRec.frx":0000
      Left            =   3030
      List            =   "CabRec.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2550
      Width           =   1935
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00F9FADC&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5460
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2580
      Width           =   2310
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Grabar"
      Enabled         =   0   'False
      Height          =   570
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   720
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   330
      Left            =   9600
      TabIndex        =   13
      Top             =   810
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
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
      Height          =   690
      Left            =   90
      TabIndex        =   1
      Top             =   0
      Width           =   10350
      Begin VB.Label Label104 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   6630
         TabIndex        =   5
         Top             =   405
         Width           =   3630
      End
      Begin VB.Label Label103 
         Height          =   225
         Left            =   6630
         TabIndex        =   4
         Top             =   180
         Width           =   3630
      End
      Begin VB.Label Label102 
         Height          =   225
         Left            =   225
         TabIndex        =   3
         Top             =   405
         Width           =   6345
      End
      Begin VB.Label Label101 
         Height          =   225
         Left            =   225
         TabIndex        =   2
         Top             =   180
         Width           =   6345
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      Height          =   615
      Left            =   90
      TabIndex        =   6
      Top             =   660
      Width           =   2685
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1245
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Recibo"
         Height          =   225
         Left            =   150
         TabIndex        =   7
         Top             =   270
         Width           =   1125
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2820
      TabIndex        =   8
      Top             =   660
      Width           =   4995
      Begin VB.Label Label2 
         Caption         =   "Crédito"
         Height          =   180
         Left            =   150
         TabIndex        =   12
         Top             =   270
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   225
         Left            =   945
         TabIndex        =   11
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label4 
         Caption         =   "Deuda"
         Height          =   180
         Left            =   2640
         TabIndex        =   10
         Top             =   270
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         Height          =   225
         Left            =   3360
         TabIndex        =   9
         Top             =   270
         Width           =   1500
      End
   End
End
Attribute VB_Name = "CabRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset
Dim RgCCTE As New ADODB.Recordset

Private Sub Form_Load()

Me.Move 0, 0, 11330, 1650

VeNotaC = False

NroRec = 0

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

CargaCVta


End Sub

Private Sub Form_Unload(Cancel As Integer)

dbCLTE.Close

End Sub

Private Sub Form_Activates()      ''  Fuera de Uso
Dim LaConstAFIP As String, DescCVTA As String

If BusClte.Visible Then
    BusClte.SetFocus
End If
    
    miSQL2 = "SELECT  * FROM Clientes WHERE CODIGO = " & CodCLTE
    RgCLTE.Open miSQL2, dbCLTE, adOpenForwardOnly, adLockReadOnly
    
    If RgCLTE.EOF Then
        RgCLTE.Close
        Exit Sub
    End If

        Label101 = "( " & CodCLTE & " )  " & ClteNomb & RgCLTE!tel1
        Label102 = RgCLTE!Dir & " ( " & RgCLTE!CP & " )  " & RgCLTE!Loc
        Label103 = "IVA " & Combo1.List(RgCLTE!CIVA - 1) & "     CUIT : " & RgCLTE!Cuit
'        Label104 = "Condición de Venta : " & RgCLTE!CVTA & "  " & Mid(Combo3.List(RgCLTE!CVTA - 1), 5, 12)
        Label3 = Format(RgCLTE!CREDIT, "###,##0.00")
        Label5 = Format(RgCLTE!DEUDA, "###,##0.00")
            
        If Combo3.ListCount > (ClteCVta - 1) Then
                Combo3.ListIndex = BuscaEnCombo(Combo3, ClteCVta)
        '        Combo3.ListIndex = RgCLTE!CVTA - 1
                DescCVTA = Mid(Combo3.List(BuscaEnCombo(Combo3, ClteCVta)), 4, 15)
        Else
                DescCVTA = "Cond. Vta."
               End If

        Label104 = "Condición de Venta : " & ClteCVta & " - " & DescCVTA
        
RgCLTE.Close

    ' Busca Próximo Nro. de Factura
    miSQL1 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
    RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockOptimistic

Text1.Text = 0

If Not RgTABL.EOF Then
    Text1.Text = Format((RgTABL!nume8 + 1), "###,##0")
End If

RgTABL.Close

' Text1.SetFocus
If Command1.Enabled Then Command1.SetFocus
If Command2.Enabled Then Command2.SetFocus

If VeNotaC Then Exit Sub

VeNotaC = True

miSQL2 = "SELECT * FROM Notaclte WHERE CLTE = " & CodCLTE
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
   
If Not RgTABL.EOF Then
    RgTABL.Close
    ExplodeForm NOTACLTE
    ' NOTACLTE.Show
    Exit Sub
End If

RgTABL.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
'If KeyCode = 40 Then SendKeys "{TAB}"
'If KeyCode = 38 Then SendKeys "+{TAB}"
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Command3_Click()

CodCLTE = 0

           Unload PieRec
           Unload DetRec
           Unload DetPago
           Unload CabRec
           Unload BusClte


VeNotaC = False

        
'            DetPago.Show
'            PieRec.Show
'            DetRec.Show
            CabRec.Show
                
                DeDonde = 2
'                BusClte.Show
            ExplodeForm BusClte

End Sub

Private Sub Command4_Click()

'NOTACLTE.Show
ExplodeForm NOTACLTE

End Sub

Private Sub Command5_Click()

'CtaCte.Show
ExplodeForm CtaCte

End Sub

Private Sub Command1_Click()

Unload PieRec
Unload DetRec
Unload DetPago
Unload Me

End Sub

Private Sub Command2_Click()

'    EmiRec.Show
ExplodeForm EmiRec

End Sub

Private Sub text1_Change()

If Text1.Text > 0 Then
    NroRec = Text1.Text
    DetRec.SetFocus
End If

End Sub

Private Sub text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode <> vbKeyReturn Then Exit Sub

If Text1.Text > 0 Then
    NroRec = Text1.Text
    DetRec.SetFocus
End If

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

If KeyAscii <> vbKeyReturn Then Exit Sub

If Text1.Text > 0 Then
    DetRec.SetFocus
End If

End Sub

Private Sub CargaCVta()
Dim I1

Combo3.Clear

    ' Tabla de Condiciones de Venta
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'CV   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

I1 = 0

If Not RgTABL.EOF Then
        Do Until RgTABL.EOF
        RSet ClteCVta = Val(RgTABL!COD)
        Combo3.AddItem ClteCVta & "-" & RgTABL!Descri
        Combo3.ItemData(I1) = RgTABL!COD
        RgTABL.MoveNext
        I1 = I1 + 1
        Loop
End If

RgTABL.Close

End Sub
