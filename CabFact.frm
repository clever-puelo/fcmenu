VERSION 5.00
Begin VB.Form CabFact 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cabecera del Comprobante"
   ClientHeight    =   4275
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
   ScaleHeight     =   4275
   ScaleMode       =   0  'User
   ScaleWidth      =   11225.4
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E9FEF1&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   150
      ScaleHeight     =   1215
      ScaleWidth      =   10935
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10995
      Begin VB.CheckBox Check1 
         Caption         =   "En Dólares"
         Height          =   315
         Left            =   450
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   150
         Width           =   1410
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0FBD5&
         Caption         =   "Nota de Débito"
         Height          =   510
         Index           =   3
         Left            =   5340
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   660
         Width           =   1800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Exportación"
         Height          =   360
         Index           =   5
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0FBD5&
         Caption         =   "Cotización"
         Height          =   510
         Index           =   4
         Left            =   7140
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   660
         Width           =   1800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0FBD5&
         Caption         =   "Nota de Crédito"
         Height          =   510
         Index           =   2
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   660
         Width           =   1800
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0FBD5&
         Caption         =   "Factura"
         Height          =   510
         Index           =   1
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00D0FBD5&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "aA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2445
         TabIndex        =   5
         Top             =   90
         Width           =   5880
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nueva"
      Height          =   345
      Left            =   8640
      TabIndex        =   27
      Top             =   810
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Emitir"
      Enabled         =   0   'False
      Height          =   555
      Left            =   9450
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   690
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   345
      Left            =   10290
      TabIndex        =   11
      Top             =   810
      Width           =   795
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
      ItemData        =   "CabFact.frx":0000
      Left            =   5490
      List            =   "CabFact.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2880
      Width           =   2310
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00F9FADC&
      Height          =   315
      ItemData        =   "CabFact.frx":0004
      Left            =   3030
      List            =   "CabFact.frx":0017
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2850
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7590
      ScaleHeight     =   180
      ScaleWidth      =   2310
      TabIndex        =   7
      Top             =   -30
      Width           =   2370
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Constancia de AFIP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   2370
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nota Clte."
      Height          =   585
      Left            =   10500
      TabIndex        =   6
      Top             =   90
      Width           =   585
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
      Left            =   150
      TabIndex        =   13
      Top             =   0
      Width           =   10320
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5025
         TabIndex        =   28
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label Label101 
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
         Left            =   225
         TabIndex        =   17
         Top             =   195
         Width           =   4755
      End
      Begin VB.Label Label102 
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
         Left            =   225
         TabIndex        =   16
         Top             =   405
         Width           =   6465
      End
      Begin VB.Label Label103 
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
         Left            =   6750
         TabIndex        =   15
         Top             =   180
         Width           =   3450
      End
      Begin VB.Label Label104 
         BackColor       =   &H00FFFFFF&
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
         Left            =   6750
         TabIndex        =   14
         Top             =   405
         Width           =   3450
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos Administrativos "
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
      Height          =   555
      Left            =   150
      TabIndex        =   18
      Top             =   660
      Width           =   6165
      Begin VB.ComboBox Combo2 
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
         Left            =   1095
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   180
         Width           =   2040
      End
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "CabFact.frx":0056
         Left            =   4230
         List            =   "CabFact.frx":0069
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   180
         Width           =   1860
      End
      Begin VB.ComboBox Combo5 
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
         ItemData        =   "CabFact.frx":009E
         Left            =   4230
         List            =   "CabFact.frx":00A0
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   180
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   23
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3150
         TabIndex        =   22
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Últ. [Comprob.] Próx"
      Height          =   555
      Left            =   6330
      TabIndex        =   24
      Top             =   660
      Width           =   2295
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00F9FADC&
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         MaxLength       =   20
         TabIndex        =   25
         Top             =   180
         Width           =   1905
      End
   End
End
Attribute VB_Name = "CabFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PorVend As String, UltVend As Single, CodCIVA As String
Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset
Dim RgCCTE As New ADODB.Recordset

Private Sub Check1_Click()
Dim I1

If Check1.Value = 1 Then
        EnDolares = True
        Check1.BackColor = vbRed
        PieFact.Label4.Caption = Label3.Caption & vbCrLf & " en Dólares"
        Label3.Caption = Label3.Caption & " en Dólares"
    Else
        EnDolares = False
        Check1.BackColor = &H8000000F
        I1 = InStr(1, Label3.Caption, " en") - 1
        Label3.Caption = Mid$(Label3.Caption, 1, I1)
        PieFact.Label4.Caption = vbCrLf & Mid$(Label3.Caption, 1, I1)
        End If

End Sub

Private Sub Form_Load()

Me.Move 0, 0, 11330, 1600

EnDolares = False

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

CodCLTEf = 0

CargaVend

CargaCVta

Option1(Val(TipoFac)).Value = True
Picture2.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
 
 dbCLTE.Close

End Sub

Private Sub Command1_Click()

If TipoFac = 2 Then
'    ImputFC.Show
    ExplodeForm ImputFC
    Exit Sub
End If

EmiFact.Show

End Sub

Private Sub Command2_Click()

Unload PieFact
Unload DetFact
Unload DetNC
Unload Me

End Sub

Private Sub Command3_Click()

NOTACLTE.Show

End Sub

Private Sub Command4_Click()

CodCLTEf = 0
Motivo = 0

 VeNotaC = True
 
           Unload PieFact
           Unload DetFact
           Unload DetNC
           Unload CabFact
           Unload Busqueda
           Unload BusClte
           Unload NOTACLTE
           
            PieFact.Show
            DetFact.Show
            CabFact.Show
            
            DeDonde = 1
            BusClte.Show

End Sub

Private Sub Text101_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    DetFact.Enabled = True
    DetFact.SetFocus
    End If

End Sub

'*************************************
'   Ingresa Vendedor
'*************************************

Private Sub Combo2_KeyPress(KeyAscii As Integer)

If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then Exit Sub

PorVend = PorVend & Chr(KeyAscii)

If Val(PorVend) > UltVend Then
    PorVend = Chr(KeyAscii)
End If

BuscaEnCombo Combo2, (String(3 - Len(PorVend), "0") & PorVend)

End Sub

Private Sub Combo2_Click()

If TipoFac = 1 Or TipoFac = 4 Then
    Combo4.TabIndex = 0
End If

If TipoFac = 2 Or TipoFac = 3 Then
    Combo5.TabIndex = 0
End If

End Sub

Private Sub Label4_Click()

'ConstAFIP.Show

ExplodeForm ConstAFIP

ConstAFIP.Text1 = ClteCUIT

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Then Exit Sub

If KeyAscii = 13 Then
    NroCpbte = Text1.Text
    If PieFact.Label304.Caption <> "" Then
        If CCur(PieFact.Label304.Caption) <> 0 Then
            Command1.Enabled = True
        End If
    End If
    If DetFact.Visible Then
        DetFact.Enabled = True
        DetFact.FG1.SetFocus
    End If
End If
'If Not BusClte.Visible And DetFact.Visible Then
'    DetFact.Enabled = True
'    DetFact.FG1.SetFocus
'End If

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

'*************************************
'   Ingresa Tipo de Venta
'*************************************

Private Sub Combo4_KeyPress(KeyAscii As Integer)
Dim ElNro As String, I1%

If KeyAscii = 13 Then
'    Text1.SetFocus
    DetFact.Enabled = True
    DetFact.SetFocus
End If

If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then Exit Sub

ElNro = Chr(KeyAscii)

For I1 = 0 To Combo4.ListCount - 1

Combo4.ListIndex = I1

If Mid(Combo4.Text, 1, 1) = ElNro Then
   Exit For
End If

Next

TipoVta = Combo4.ListIndex + 1

DetFact.Enabled = True
DetFact.SetFocus

End Sub

Private Sub Combo4_Click()

TipoVta = Combo4.ListIndex + 1

'If Text1.Visible Then Text1.SetFocus

End Sub

'*************************************
'   Ingresa Motivo de N.Crédito
'*************************************

Private Sub Combo5_KeyPress(KeyAscii As Integer)
Dim ElNro As String, I1%

If KeyAscii = 13 Then
'    Text1.SetFocus
End If

If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then Exit Sub

ElNro = Chr(KeyAscii)

For I1 = 0 To Combo5.ListCount - 1

'Combo5.ListIndex = i1

If Val(Mid(Combo5.List(I1), 1, 2)) = ElNro Then
   Combo5.ListIndex = I1
   Exit For
End If

Next

End Sub

Private Sub Combo5_Click()
    
Unload DetFact
Unload DetNC
Importe = 0

'If Text1.Visible Then Text1.SetFocus
Motivo = Mid(Combo5.Text, 1, 2)
If Motivo = 1 Then
    DetFact.Show
Else
    DetNC.Show
End If

If BusClte.Visible Then
    BusClte.SetFocus
Else
    Picture2.Visible = False
End If

End Sub

Private Sub Label101_Change()

If CodCLTEf = 0 Then Exit Sub
        
        PieFact.Show
        DetFact.Show

'miSQL2 = "SELECT * FROM Notaclte WHERE CLTE = " & CodCLTEf
'RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
   
'If RgTABL.EOF = False Then
'    RgTABL.Close
'    NOTACLTE.Show
'    Exit Sub
'End If

'RgTABL.Close

End Sub

Private Sub Option1_Click(Index As Integer)
        
Check1.Value = 0
        
Picture2.BackColor = &HE9FEF1
PieFact.Label4.BackColor = &HE9FEF1

TipoFac = Index
Motivo = 1

Select Case Index
    
    Case 1
        Frame3.Caption = "  Últ.  [Factura] Próx."
        Label3.Caption = "Factura"
        PieFact.Label4.Caption = vbCrLf & "Factura"
'        PieFact.Label4.Caption = vbCrLf & "Saldo: " & Format(SaldoClte, "$###,##0.00")
        FCMENU.StatusBar1.Panels.item(1).Text = "Factura"
        
    Case 2
        Frame3.Caption = "  Últ.  [N.Créd.] Próx."
        Label3.Caption = "Nota de Crédito"
        PieFact.Label4.Caption = vbCrLf & "Nota de Crédito"
        FCMENU.StatusBar1.Panels.item(1).Text = "Nota de Crédito"
        Picture2.BackColor = &HC0FFFF
        PieFact.Label4.BackColor = &HC0FFFF
    Case 3
        Frame3.Caption = "  Últ.  [N.Déb.] Próx."
        Label3.Caption = "Nota de Débito"
        PieFact.Label4.Caption = vbCrLf & "Nota de Débito"
        FCMENU.StatusBar1.Panels.item(1).Text = "Nota de Débito"
        Picture2.BackColor = &HC0FFFF
        PieFact.Label4.BackColor = &HC0FFFF
        
    Case 4
        Label3.Caption = "Cotización"
        Frame3.Caption = "  Últ.  [Cotizac.] Próx."
        PieFact.Label4.Caption = vbCrLf & "Cotización"
        FCMENU.StatusBar1.Panels.item(1).Text = "Cotización"
        
    Case 5
        Label3.Caption = "FACTURA"
        Frame3.Caption = " Últ. [Export.] Próx."
        PieFact.Label4.Caption = vbCrLf & "FACTURA"
        FCMENU.StatusBar1.Panels.item(1).Text = "FACTURA"
        Picture2.BackColor = &HC0C0FF
        PieFact.Label4.BackColor = &HC0C0FF

End Select

BuscaUltima

If BusClte.Visible Then
    BusClte.SetFocus
Else
    Picture2.Visible = False
End If

Select Case TipoFac
    Case 1, 4
        Label2.Caption = "Forma Ped.:"
        Combo4.Visible = True
        Combo5.Visible = False
        Combo4.ListIndex = 1
    Case 2
        CargaMoti
        Label2.Caption = "Motivo NC:"
        Combo5.Visible = True
        Combo4.Visible = False
        Combo5.ListIndex = 0
    Case 3
        CargaMoti
        Label2.Caption = "Motivo ND:"
        Combo5.Visible = True
        Combo4.Visible = False
        Combo5.ListIndex = 0
    Case 4
        Label2.Caption = ""
        Combo5.Visible = False
        Combo4.Visible = False
        Combo4.ListIndex = 1

End Select

End Sub

Private Sub CargaVend()
Dim I1

Combo2.Clear

    ' Tabla de Vendedores
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'VD   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

Do Until RgTABL.EOF
Combo2.AddItem Mid(RgTABL!COD, 1, 3) & "-" & RgTABL!Descri
UltVend = RgTABL!COD
RgTABL.MoveNext
Loop

RgTABL.Close

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

Private Sub CargaMoti()
Dim I1

Combo5.Clear

    ' Tabla de Motivos de NC y ND
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'MT   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

I1 = 0

If Not RgTABL.EOF Then
        Do Until RgTABL.EOF
        RSet ClteCVta = Val(RgTABL!COD)
        Combo5.AddItem ClteCVta & "-" & RgTABL!Descri
        Combo5.ItemData(I1) = RgTABL!COD
        RgTABL.MoveNext
        I1 = I1 + 1
        Loop
End If

RgTABL.Close

End Sub

Sub BuscaUltima()
Dim ElNro As String

    ' Busca Último Nro. de Factura

Text1.Text = "": ElNro = ""

If ClteCIVA < "1" Then Exit Sub

        If ClteCIVA < 3 Or (ClteCIVA = 4 And CltePCIA = "V ") Then
            LaLetra = "A"
        Else
            LaLetra = "B"
        End If

'miSQL2 = "SELECT CPBTE  FROM FCIVAVTA WHERE TIPO = '" & TipoFac & _
'        "' AND LETRA = '" & LaLetra & "' ORDER BY CPBTE DESC"
miSQL2 = "SELECT CPBTE  FROM FCIVAVTA WHERE LETRA = '" & LaLetra & _
         "' ORDER BY CPBTE DESC"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

If Not RgTABL.EOF Then
    Text1.Text = Format(RgTABL!Cpbte, "###,##0")
Else
    Text1.Text = "0"
End If

Salir1:

RgTABL.Close

    ' Busca Próximo Nro. de Factura
    miSQL1 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
    RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockOptimistic

If RgTABL.EOF Then GoTo Salir2

Select Case TipoFac
    
    Case 1
        If LaLetra = "A" Then
            ElNro = Format((RgTABL!nume1 + 1), "###,##0")
        Else
            ElNro = Format((RgTABL!nume5 + 1), "###,##0")
        End If
        
    Case 2
        If LaLetra = "A" Then
            ElNro = Format((RgTABL!nume1 + 1), "###,##0")
        Else
            ElNro = Format((RgTABL!nume5 + 1), "###,##0")
        End If
    
    Case 3
        If LaLetra = "A" Then
            ElNro = Format((RgTABL!nume1 + 1), "###,##0")
        Else
            ElNro = Format((RgTABL!nume5 + 1), "###,##0")
        End If
    
    Case 4
            ElNro = Format((RgTABL!nume4 + 1), "###,##0")

End Select

Text1.Text = Text1.Text & " / " & ElNro

Salir2:

RgTABL.Close

End Sub


