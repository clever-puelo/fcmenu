VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Zcocina1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Arreglo de Fechas en Base de Datos"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker FECnva 
      Height          =   330
      Left            =   2625
      TabIndex        =   3
      ToolTipText     =   "Fecha nueva que se desea colocar"
      Top             =   1155
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777152
      CustomFormat    =   "ddd d MMM yyy"
      Format          =   57081859
      CurrentDate     =   36877
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
      Left            =   5295
      TabIndex        =   1
      Top             =   3975
      Width           =   1230
   End
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
      Left            =   3885
      TabIndex        =   0
      Top             =   3975
      Width           =   1230
   End
   Begin MSComCtl2.DTPicker FECvie 
      Height          =   330
      Left            =   2625
      TabIndex        =   2
      ToolTipText     =   "Fecha Original que se desea corregir"
      Top             =   585
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777152
      CustomFormat    =   "ddd d MMM yyy"
      Format          =   57081859
      CurrentDate     =   36877
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Este proceso modifica la fecha de grabación de las tablas de datos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1050
      TabIndex        =   6
      ToolTipText     =   "Fecha nueva que se desea colocar"
      Top             =   2025
      Width           =   4755
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha Nueva :"
      Height          =   270
      Left            =   930
      TabIndex        =   5
      ToolTipText     =   "Fecha nueva que se desea colocar"
      Top             =   1185
      Width           =   1635
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha Ant. :"
      Height          =   270
      Left            =   945
      TabIndex        =   4
      ToolTipText     =   "Fecha Original que se desea corregir"
      Top             =   630
      Width           =   1635
   End
End
Attribute VB_Name = "Zcocina1"
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

If Respuesta = vbNo Then Unload Me

doCOC

Unload Me

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

dbCOC.ConnectionString = BDatos1
dbCOC.Open

FECvie.Value = Date
FECnva.Value = Date


End Sub

Private Sub doCOC()
Dim FECHADsd As String, FECHAHst As String, ElDia As String * 2, FECmas1 As Date

FECHADsd = FECvie.Month & "/" & FECvie.Day & "/" & FECvie.Year
FECmas1 = FECvie.Value + 1
'FECHAHst = FECmas1.Month & "/" & FECmas1.Day & "/" & FECmas1.Year
FECHAHst = Month(FECmas1) & "/" & Day(FECmas1) & "/" & Year(FECmas1)

dbCOC.BeginTrans
    
'----------   Corrige Subdiario   ------
miSQL2 = "SELECT * FROM FCIVAVTA WHERE FECHA >= " & Separa & _
                        FECHADsd & Separa & " AND FECHA < " & Separa & FECHAHst & Separa
If EnElArchivo(miSQL2) Then
    dbCOC.RollbackTrans
    Exit Sub
End If
    
'----------   Corrige Cta.Cte.   ------
miSQL2 = "SELECT * FROM CTASCTES WHERE FECHA >= " & Separa & _
                        FECHADsd & Separa & " AND FECHA < " & Separa & FECHAHst & Separa

If EnElArchivo(miSQL2) Then
    dbCOC.RollbackTrans
    Exit Sub
End If
    
'----------   Corrige Estadísticas   ------
miSQL2 = "SELECT * FROM FCESTAD1 WHERE FECHA >= " & Separa & _
                        FECHADsd & Separa & " AND FECHA < " & Separa & FECHAHst & Separa

If EnElArchivo(miSQL2) Then
    dbCOC.RollbackTrans
    Exit Sub
End If
    
'----------   Corrige Mov.Stock   ------
miSQL2 = "SELECT * FROM MOVSTOCK WHERE FECHA >= " & Separa & _
                        FECHADsd & Separa & " AND FECHA < " & Separa & FECHAHst & Separa

If EnElArchivo(miSQL2) Then
    dbCOC.RollbackTrans
    Exit Sub
End If

dbCOC.CommitTrans

End Sub

Private Function EnElArchivo(Busca As String) As Boolean

EnElArchivo = False

RgCOC.Open Busca, dbCOC, adOpenDynamic, adLockPessimistic

If RgCOC.EOF Then
    RgCOC.Close
    EnElArchivo = True
    Exit Function
End If

Do Until RgCOC.EOF

RgCOC!FECHA = FECnva.Value
        
RgCOC.Update

RgCOC.MoveNext

Loop

RgCOC.Close

End Function


Private Sub Form_Unload(Cancel As Integer)

dbCOC.Close

End Sub
