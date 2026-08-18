VERSION 5.00
Begin VB.Form Zcocina2 
   Caption         =   "Form1"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4410
   ScaleWidth      =   6915
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4035
      TabIndex        =   1
      Top             =   3735
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
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
      Height          =   420
      Left            =   5445
      TabIndex        =   0
      Top             =   3735
      Width           =   1230
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Comprobante :"
      Height          =   270
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Fecha Original que se desea corregir"
      Top             =   1005
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Height          =   270
      Left            =   3855
      TabIndex        =   3
      ToolTipText     =   "Fecha nueva que se desea colocar"
      Top             =   1005
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Este proceso modifica el nro. de ítem en la tabla FCESTAD1 y en MOVSTOCK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Fecha nueva que se desea colocar"
      Top             =   1785
      Width           =   4755
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Zcocina2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCOC As New ADODB.Connection
Dim RgCOC As New ADODB.Recordset
Dim Mensaje As String

Private Sub Command1_Click()
                
Mensaje = "Está seguro que desea modificar" & vbCrLf & "Las bases de datos ?" _
            & vbCrLf & vbCrLf & "                   ... continúa ?"
Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNo, "Modificación de Fechas")

If Respuesta = vbNo Then
    Unload Me
    Exit Sub
    End If
    
doCOC

Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

dbCOC.ConnectionString = BDatos1
dbCOC.Open

End Sub

Private Sub doCOC()
Dim Cpbte As Double

Cpbte = 2000

Do Until Cpbte > 15000

Cpbte = Cpbte + 1
Label1.Caption = Cpbte

'Me.Show
DoEvents

dbCOC.BeginTrans

'----------   Corrige Estadísticas   ------
miSQL2 = "SELECT * FROM FCESTAD1 WHERE CPBTE = " & Cpbte
If EnElArchivo(miSQL2) Then
    dbCOC.RollbackTrans
    GoTo AlLoop1
End If

dbCOC.CommitTrans

AlLoop1:

Loop

Cpbte = 2000

Do Until Cpbte > 15000

Cpbte = Cpbte + 1
Label1.Caption = Cpbte

'Me.Show
DoEvents

dbCOC.BeginTrans
    
'----------   Corrige Movim.Stock   ------
miSQL2 = "SELECT * FROM MOVSTOCK WHERE CPBTE = " & Cpbte

If EnElArchivo(miSQL2) Then
    dbCOC.RollbackTrans
    GoTo AlLoop2
End If


dbCOC.CommitTrans

AlLoop2:

Loop

End Sub

Private Function EnElArchivo(Busca As String) As Boolean
Dim item As Double

item = 0

EnElArchivo = False

RgCOC.Open Busca, dbCOC, adOpenDynamic, adLockPessimistic

If RgCOC.EOF Then
    RgCOC.Close
    EnElArchivo = True
    Exit Function
End If

Do Until RgCOC.EOF

item = item + 1

RgCOC!item = item
        
RgCOC.Update

RgCOC.MoveNext

Loop

RgCOC.Close

End Function


Private Sub Form_Unload(Cancel As Integer)

dbCOC.Close

End Sub


