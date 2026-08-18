VERSION 5.00
Begin VB.Form ABMArt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambios a los Items"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8205
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Caption         =   "Baja"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3825
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2805
      Width           =   915
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Alta"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1890
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2805
      Width           =   915
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cambio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2850
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2805
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   195
      TabIndex        =   11
      Top             =   2805
      Width           =   1300
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
      Height          =   315
      Left            =   6735
      TabIndex        =   9
      Top             =   2805
      Width           =   1300
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
      Height          =   315
      Left            =   5325
      TabIndex        =   8
      Top             =   2805
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Caption         =   "  Datos del Ítem"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2745
      Left            =   135
      TabIndex        =   12
      Top             =   15
      Width           =   7965
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   26
         ToolTipText     =   "Seleccione la Sección a Visualizar"
         Top             =   150
         Width           =   1335
      End
      Begin VB.TextBox Text109 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
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
         Height          =   285
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   22
         ToolTipText     =   "Stock de Reposición"
         Top             =   2280
         Width           =   1980
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         MaxLength       =   8
         TabIndex        =   0
         Top             =   150
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FEFADE&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#9.99"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6750
         MaxLength       =   5
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox Text108 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
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
         Height          =   285
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   7
         ToolTipText     =   "Stock de Reposición"
         Top             =   1980
         Width           =   1980
      End
      Begin VB.TextBox Text107 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
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
         Height          =   285
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   6
         ToolTipText     =   "Stock Máximo"
         Top             =   1680
         Width           =   1980
      End
      Begin VB.TextBox Text106 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
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
         Height          =   285
         Left            =   1620
         MaxLength       =   8
         TabIndex        =   5
         ToolTipText     =   "Stock Mínimo"
         Top             =   1380
         Width           =   1980
      End
      Begin VB.TextBox Text105 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1620
         MaxLength       =   15
         TabIndex        =   4
         ToolTipText     =   "Precio de Costo"
         Top             =   1080
         Width           =   1980
      End
      Begin VB.TextBox Text104 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FEFADE&
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1620
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   "Precio de Venta"
         Top             =   780
         Width           =   1980
      End
      Begin VB.TextBox Text102 
         BackColor       =   &H00FEFADE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1620
         MaxLength       =   25
         TabIndex        =   2
         ToolTipText     =   "Descripción (o dejar en Blanco)"
         Top             =   480
         Width           =   4515
      End
      Begin VB.Label Label7 
         Caption         =   "Stock             :"
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
         Left            =   135
         TabIndex        =   23
         Top             =   2310
         Width           =   1515
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Pulgadas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2940
         TabIndex        =   21
         Top             =   210
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Metros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5220
         TabIndex        =   20
         Top             =   210
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Stock Ant .     :"
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
         Left            =   135
         TabIndex        =   19
         Top             =   2025
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Stock Mín.      :"
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
         Left            =   135
         TabIndex        =   18
         Top             =   1725
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "Prec. Venta 2 :"
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
         Left            =   135
         TabIndex        =   17
         Top             =   1125
         Width           =   1470
      End
      Begin VB.Label Label1 
         Caption         =   "Prec. Costo    :"
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
         Left            =   135
         TabIndex        =   16
         ToolTipText     =   "Cantidad a Preparar"
         Top             =   1425
         Width           =   1395
      End
      Begin VB.Label Label109 
         Caption         =   "Precio de Venta :"
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
         Left            =   135
         TabIndex        =   15
         Top             =   840
         Width           =   1470
      End
      Begin VB.Label Label102 
         Caption         =   "Descripción   :"
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
         Left            =   135
         TabIndex        =   14
         Top             =   570
         Width           =   1395
      End
      Begin VB.Label Label101 
         Caption         =   "Sección         :  "
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
         Left            =   135
         TabIndex        =   13
         Top             =   270
         Width           =   1425
      End
   End
End
Attribute VB_Name = "ABMArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mensaje As String, VeNotaA As Boolean
Dim CAMPOS  As String, Lgo1, Lgo2, Deci1, Deci2
Dim s$

Dim dbART As New ADODB.Connection, RgART As New ADODB.Recordset

Private Sub Form_Activate()

If VeNotaA Then Exit Sub

VeNotaA = True

miSQL2 = "SELECT * FROM Notartic WHERE cod1 = '" & Correa & _
                    "' and cod2 = '" & RESTO & "'"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
   
If Not RgTABL.EOF Then
    RgTABL.Close
    NOTARTIC.Show
    Exit Sub
End If
    
RgTABL.Close

End Sub

Private Sub Form_Load()
Dim I1%

    ' Tabla de Secciones
miSQL2 = "SELECT COD, DESCRI  FROM Fctabla1 WHERE CTAB = 'SC   ' ORDER BY COD"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

I1 = 0

If Not RgTABL.EOF Then
        Do Until RgTABL.EOF
        Combo2.AddItem RgTABL!COD
        Combo2.ItemData(I1) = I1
        RgTABL.MoveNext
        I1 = I1 + 1
        Loop
End If

RgTABL.Close

Combo2.ListIndex = 0

Me.Move 0, 0, 8300, 3550

dbART.ConnectionString = BDatos1
dbART.Open
   
Command3.Visible = False
VeNotaA = False
'        TipoMov2 = 1

Select Case TipoMov2

Case 1

    Frame1.Enabled = True
    Blanquea
    ABMArt.Caption = "Alta de un Artículo"

Case 2, 3
    
    ABMArt.Caption = "Cambios a un Artículo"

    Command3.Visible = True

miSQL2 = "SELECT  * FROM Articulo WHERE COD1 = '" & Correa & "' and COD2 = '" & _
                         RESTO & "'"
RgART.Open miSQL2, dbART, adOpenForwardOnly, adLockReadOnly
   
If RgART.EOF Then
    RgART.Close
    Exit Sub
End If
    

    Combo2.Enabled = False

For I1 = 0 To Combo2.ListCount
If Trim(Combo2.List(I1)) = Trim(RgART!COD1) Then
    Combo2.ListIndex = I1
    Exit For
End If
Next I1
    
    Text102.Text = RgART!Descri
    Text104.Text = RgART!Prec
    Text105.Text = RgART!vta1
    Text106.Text = RgART!PCos
    Text107.Text = RgART!Stmin
    Text108.Text = RgART!Stant
    Text109.Text = RgART!Stock
        
End Select

If TipoMov2 = 2 Then
                ABMArt.Caption = "Baja de un Item"
                Frame1.Enabled = False
                
                Command2.Enabled = True
                Command2.Caption = "Eliminar"
End If

If VerDetArt Then
                ABMArt.Caption = "Consulta de un Artículo"
                Frame1.Enabled = False
                
                Command2.Visible = False
End If

If TipoMov2 <> 1 Then
  BuscaSeccion
  Text2.Text = Mid(RESTO, 1, Lgo1)
  Text3.Text = Mid(RESTO, Lgo1 + 1, Lgo2)
  Text2.Enabled = False
  Text3.Enabled = False
  RgART.Close
  End If
  
' If NotasITEM Then NOTAITEM.Show


End Sub

Private Sub Form_Unload(Cancel As Integer)

dbART.Close

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF2 Then
    Busqueda.Show
    Unload Me
        End If

If KeyCode = 13 Or KeyCode = 40 Or KeyCode = 39 Then SendKeys "{TAB}"
If KeyCode = 38 Or KeyCode = 37 Then SendKeys "+{TAB}"
If KeyCode = 27 Then Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

End Sub

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command2_Click()

Grabacion
' Unload Me

' CONSULTA.DoVer1

End Sub

Private Sub Command3_Click()

NOTARTIC.Show

End Sub

Private Sub Command5_Click()

If Frame1.Enabled Then
                Caption = "Detalle del Artículo "
                Frame1.Enabled = False
                Command2.Visible = False
                Command5.Caption = "Cambio"
'                Command5.Visible = False
'                Text102.SetFocus
Else

                Caption = "Cambios al Artículo "
                Frame1.Enabled = True
                Command2.Visible = True
                Command5.Caption = "Sólo Ver"
'                Command5.Visible = True
                Text102.SetFocus
End If

End Sub

Private Sub Command4_Click()

Frame1.Enabled = True

    Blanquea

End Sub

Private Sub Command6_Click()

Command4.Visible = False
TipoMov2 = 2
Grabacion

Unload Me

Busqueda.LimpiaGrilla
Busqueda.BuscaDesdeAfuera

End Sub


'--------------------
'   Sección
Private Sub Combo2_Click()

Label6.Visible = False: Label3.Visible = False: Text2.Visible = False: Text3.Visible = False

miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'SC   ' AND COD = '" & Combo2.Text & "'"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
   
If RgTABL.EOF Then
        vbMsgBoxTitle = "Error de Ingreso"
        vbMsgBoxText = "  " & vbCrLf & "No Existe Sección"
        vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
         MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle
         Combo2.SetFocus
    RgTABL.Close
    Exit Sub
End If
    
RgTABL.Close

Command1.Enabled = True
    
BuscaSeccion

End Sub

'----
'   Primer Valor de la Clave

Private Sub Text2_KeyPress(KeyAscii As Integer)


If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'----
'   Segundo Valor de la Clave

Private Sub Text3_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

'*************************************
'   Descripción
'*************************************

Private Sub Text102_GotFocus()

Text102.SelStart = 0
Text102.SelLength = Len(Text102.Text)

End Sub

'*************************************
'   Ingresa Precios
'*************************************

Private Sub Text104_GotFocus()

Text104.SelStart = 0
Text104.SelLength = Len(Text104.Text)

End Sub

Private Sub Text105_GotFocus()

Text105.SelStart = 0
Text105.SelLength = Len(Text105.Text)

End Sub

Private Sub Text106_GotFocus()

Text106.SelStart = 0
Text106.SelLength = Len(Text106.Text)

End Sub

'*************************************
'   Stocks
'*************************************

Private Sub Text107_GotFocus()

Text107.SelStart = 0
Text107.SelLength = Len(Text107.Text)

End Sub

Private Sub Text108_KeyPress(KeyAscii As Integer)

        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
            KeyAscii = KeyAscii + 32
            End If

End Sub

Private Sub Text109_KeyPress(KeyAscii As Integer)

        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
            KeyAscii = KeyAscii + 32
            End If

End Sub

Private Sub Text104_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text105_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text106_KeyPress(KeyAscii As Integer)

If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii = Asc(".") Or KeyAscii = Asc("-") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Grabacion()
Dim I1%, i2%, i3%, Codi1 As String, Codi2 As String

Select Case TipoMov2
        Case 1
                Mensaje = "Desea Grabar ?"
        Case 2
                Mensaje = "Desea Eliminar ?"
        Case 3
                Mensaje = "Desea Cambiar ? "
End Select

Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "ABM de ITEMS")

If Respuesta = vbNo Then GoTo Final1
If Respuesta = vbCancel Then
         Exit Sub
         End If
        
     
'******************************************
'           Graba en Items
'******************************************
'Codi2 = String(Lgo1 - Len(Text2.Text), "0") & Text2.Text & String(Lgo2 - Len(Text3.Text), "0") & Text3.Text
I1 = InStr(1, Text2.DataFormat.Format, ",")
i2 = InStr(1, Text2.Text, ",")
If I1 > 0 And i2 = 0 Then
    Text2.Text = Text2.Text & "," & String(Deci1, "0")
End If

I1 = InStr(1, Text3.DataFormat.Format, ",")
i2 = InStr(1, Text3.Text, ",")
If I1 > 0 And i2 = 0 Then
    Text3.Text = Text3.Text & "," & String(Deci2, "0")
End If

Codi1 = Replace(Text2.Text, ",", "")

'----  Si la U.Med. es Número lo saltea
If Lgo1 = 5 And Lgo2 = 0 Then GoTo EsPorNro
        
        If Len(Codi1) < Lgo1 Then
            Codi1 = String(Lgo1 - Len(Codi1), "0") & Codi1
        End If
        
        Codi2 = Replace(Text3.Text, ",", "")
        If Len(Codi2) < Lgo2 Then
            Codi1 = String(Lgo2 - Len(Codi2), "0") & Codi2
        End If

EsPorNro:

Codi2 = Codi1 & Codi2

miSQL2 = "SELECT  * FROM Articulo WHERE COD1 = '" & Combo2.Text & _
                     "' and COD2 = '" & Codi2 & "'"
RgART.Open miSQL2, dbART, adOpenDynamic, adLockPessimistic
   

    If RgART.EOF = True Then
        RgART.AddNew
        
        RgART!COD1 = Trim(Combo2.Text)
'        RgART!Cod2 = Codi2
        RgART!Cod2 = Codi2
        RgART!Descri = ""
        RgART!Prec = 0
        RgART!vta1 = 0
        RgART!vta2 = 0
        RgART!PCos = 0
        RgART!dto1 = 0
        RgART!Stock = 0
        RgART!Stmin = 0
        RgART!mcalis = 0
        RgART!Stant = 0
        RgART!gptipo = 0
        RgART!gpdsd = 0
        RgART!gphst = 0
    End If

If TipoMov2 = 2 Then
        RgART.Delete
        RgART.Close
        Exit Sub
        End If
    
If Text102.Text <> "" Then
    RgART!Descri = RTrim(Text102.Text)
Else
    RgART!Descri = Combo2.Text & " " & Codi2
End If

    RgART!Prec = Text104.Text
    RgART!vta1 = Text105.Text
    RgART!PCos = Text106.Text
    RgART!Stmin = Val(Text107.Text)
    RgART!Stant = Val(Text108.Text)
    RgART!Stock = Val(Text109.Text)

RgART!FACTUAL = Date
RgART!USUARIO = Mid$(Red_Usuario, 1, 6)

RgART.Update
RgART.Close

Final1:

Blanquea
Unload Me

Final2:

End Sub

Sub MuestraError()
          
        vbMsgBoxTitle = "Error de Ingreso"
        vbMsgBoxText = "  " & vbCrLf & "Código  NO  Existe  "
        vbMsgBoxResp = vbOKOnly + vbCritical + vbSystemModal + vbDefaultButton1
        MsgBox vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle

End Sub

Sub Blanquea()
    
Frame1.Enabled = True

    Text102.Text = ""
    Text109.Text = ""
    Text104.Text = ""
    Text105.Text = ""
    Text106.Text = ""
    Text107.Text = ""
    Text108.Text = ""

Combo2.Enabled = True
Text102.Enabled = True

If Combo2.Visible Then
    Combo2.SetFocus
End If

End Sub

Private Sub BuscaSeccion()

Dim CAMPO1 As String, CAMPO2 As String, CAMPO3 As String
Dim UMED1 As String, UMED2 As String, UMED3 As String
Dim Codi As String
 
miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'SC   ' AND COD = '" & Combo2.Text & "'"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
   
If RgTABL.EOF Then
    RgTABL.Close
    Exit Sub
End If

UMED1 = RgTABL!ALF1
UMED2 = RgTABL!ALF2
UMED3 = RgTABL!ALF3
    
RgTABL.Close

If UMED1 <> "     " Then
    miSQL2 = "SELECT * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED1 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    Lgo1 = RgTABL!NumSD1 + RgTABL!NumSD2
    Deci1 = RgTABL!NumSD2
    CAMPO1 = String(RgTABL!NumSD1 - 1, "#") & "0"
    If RgTABL!NumSD2 > 0 Then CAMPO1 = CAMPO1 & "," & String(RgTABL!NumSD2, "0")
    Text2.DataFormat.Format = CAMPO1
    Text2.MaxLength = Len(CAMPO1)
    Label6.Caption = RgTABL!Descri
    Label6.Visible = True: Text2.Visible = True
    RgTABL.Close
'    Text2.SetFocus
    End If

If UMED2 <> "     " Then
    miSQL2 = "SELECT  * FROM Fctabla1 WHERE CTAB = 'UM   ' AND COD = '" & UMED2 & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    Lgo2 = RgTABL!NumSD1 + RgTABL!NumSD2
    Deci2 = RgTABL!NumSD2
    CAMPO2 = String(RgTABL!NumSD1 - 1, "#") & "0"
    If RgTABL!NumSD2 > 0 Then CAMPO2 = CAMPO2 & "," & String(RgTABL!NumSD2, "0")
    Text3.DataFormat.Format = CAMPO2
    Text3.MaxLength = Len(CAMPO2)
    Label3.Caption = RgTABL!Descri
    Label3.Visible = True: Text3.Visible = True
    If UMED1 = "     " Then Text3.SetFocus
    RgTABL.Close
    End If

End Sub


