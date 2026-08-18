VERSION 5.00
Begin VB.Form Inicio 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   4050
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   7080
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCCIÓN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   450
         Left            =   1260
         TabIndex        =   14
         Top             =   1305
         Width           =   4530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4620
         TabIndex        =   13
         Top             =   3300
         Width           =   2280
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   12
         Top             =   3375
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3150
         TabIndex        =   11
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E1FCE0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FCMENU"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2280
         TabIndex        =   10
         Top             =   600
         Width           =   2550
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Clever G. Colalillo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   4470
         TabIndex        =   9
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Facturación y Cuentas Corrientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1995
         TabIndex        =   8
         Top             =   2145
         Width           =   3075
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright    :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   4
         Top             =   2850
         Width           =   975
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alestel S.R.L."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4500
         TabIndex        =   3
         Top             =   2850
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Advertencia : Este programa es para uso exclusivo de ALESTEL y para todos los que ella Autorice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   390
         TabIndex        =   2
         Top             =   3750
         Width           =   6285
      End
      Begin VB.Label lblPlatform 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows 9x / NT / 2000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   5
         Top             =   3120
         Width           =   2190
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   330
         TabIndex        =   6
         Top             =   2760
         Width           =   2310
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizado al Sector Ventas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   1
         Top             =   210
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modulo de Ventas "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   7
         Top             =   1755
         Width           =   2790
      End
   End
End
Attribute VB_Name = "Inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" ( _
    ByVal hwndParent As Long, ByVal fRequest As Long, _
    ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

Private Sub Form_Deactivate()
 
 unload Me

'Call SlideWindow(Me, 50)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
'If KeyCode = vbKeyF12 Then
'   RED.Show
'   End If

If KeyCode = vbKeyF12 Then
    If FCMENU.COMPLEM.Visible Then
        FCMENU.COMPLEM.Visible = False
        FCMENU.Toolbar3.Buttons(1).Enabled = False
        FCMENU.NCINT.Enabled = False
    Else
        FCMENU.COMPLEM.Visible = True
        FCMENU.Toolbar3.Buttons(1).Enabled = True
        FCMENU.NCINT.Enabled = True
   End If
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   
' Unload RED
' Unload Me

Call SlideWindow(Me, 50)

End Sub

Private Sub Form_Load()
Dim ix As Double, I1%, Err As Boolean, VerImpre As String
Dim NomDSN As String, DescDSN As String, BDPath As String
Dim FecBack As Date

Me.width = 7300
Me.height = 4270

If HOMO Then
    Label6.Caption = "Homologación/Prueba"
    BDPath = "c:\Temp\FCMENUPr.mdb"
    ServidorSQL = BDPath
    TipoDB = "A": NomDSN = "FCMENU": DescDSN = "Base Datos Facturación"
    GoTo SaltaQuery
Else
    Label6.Caption = "PRODUCCIÓN"
End If

Me.Show
DoEvents

If TipoDB = "" Then
    TipoDB = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos", "Driver")
End If

DescDSN = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Descrip")
NomDSN = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Nombre")

BDPath = BDPath & QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Nombre")
If TipoDB = "A" Then
    If ServidorSQL = "" Then
        BDPath = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\DSN", "Carpeta") & BDPath & ".mdb"
    Else
        BDPath = ServidorSQL & BDPath & ".mdb"
    End If
End If

    
    If NomDSN = "" Or TipoDB = "" Then
        MsgBox "Falta el DSN (Driver ODBC)." & vbCrLf & " " & vbCrLf _
        & "Ejecute la Generación.", vbCritical, "FCMENU - Error Grave"
        End
    End If

 ' Si no existe la base de datos Termina el Programa
If TipoDB = "A" Then
    If Len(Dir$(BDPath)) < 1 Then
        MsgBox "Falta la Base de Datos", vbCritical, "FCMENU - Error Grave"
        End
    End If
End If

SaltaQuery:

GenSystemDSN TipoDB, NomDSN, DescDSN, BDPath
BDatos1 = "dsn=" & NomDSN
BDatos2 = "dsn=PadronARBA"

'Separador para la busqueda (SELECT) de Fecha
If TipoDB = "A" Then
    Separa = "#"
Else
    Separa = "'"
End If

'  Carga de Drivers ODBC
'For i1 = 0 To List1.ListCount - 1
'    GenSystemDSN Params(3), List1.List(i1), List3.List(i1), List2.List(i1)
'Next

VerImpre = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Impresora", "1")
If SaleImpr = "" Then SaleImpr = VerImpre

Inicio.Label3 = "Punto de Venta : " & PtoVta & "  Imprime en: " & SaleImpr
Inicio.Label4 = "Servidor SQL : " & TipoDB & " " & BDPath
Inicio.Label5 = "Imprime en : " & SaleImpr

    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
   
   AutOK = True

    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision

Realizado = QueryValue(HKEY_CURRENT_USER, "FCMENU\Datos\Email", "Realizado")

If Realizado <> "" Then
 FecBack = Mid(Realizado, 1, 10)
 If FecBack < Date Then FuerzaBack = True
 End If

End Sub

Private Sub Frame1_Click()
'     Unload RED

    unload Me

End Sub

Private Sub imgLogo_Click()

End Sub

Function GenSystemDSN(Dvr As String, NomDSN As String, DescDSN As String, BDPath As String)
Const ODBC_ADD_SYS_DSN = 4    'Agregar origen de datos

    Dim Driver As String
    Dim Attributes As String
    Dim Ret As Long

Attributes = Attributes & "DESCRIPTION=" & DescDSN & Chr$(0)
Attributes = Attributes & "DSN=" & NomDSN & Chr(0)
Attributes = Attributes & "Database=" & BDPath & Chr(0)

Select Case Dvr

    Case "S"   'Driver para SQL
         Driver = "SQL Server"
         Attributes = Attributes & "SERVER=" & ServidorSQL & Chr$(0)
         Attributes = Attributes & "Trusted_Connection=True" & Chr(0)
         
    Case "A"   'Driver para ACCESS
        Attributes = Attributes & "DBQ=" & BDPath & Chr(0)
        Driver = "Microsoft Access Driver (*.mdb)"
        
    Case "D"   'Driver para SQL
         Driver = "Microsoft dBase Driver (*.dbf)"
End Select
'Attributes = Attributes & "Uid=clever" & Chr(0) & "pwd=bajo" & Chr(0)

Ret = SQLConfigDataSource(0, ODBC_ADD_SYS_DSN, Driver, Attributes)
'Debug.Print Attributes
 'ret es igual a 1 si se realiza correctamente y 0 si hay un error
If Ret <> 1 Then
    MsgBox "Error al crear DSN"
End If

End Function
