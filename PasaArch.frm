VERSION 5.00
Begin VB.Form PasaArch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pasaje de Datos del Sistema Viejo"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6540
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
   ScaleHeight     =   3330
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command8 
      Caption         =   "Correcciones"
      Height          =   405
      Left            =   1890
      TabIndex        =   10
      Top             =   2700
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Desctos."
      Height          =   405
      Left            =   1890
      TabIndex        =   9
      Top             =   315
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   405
      Left            =   4605
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cheques"
      Height          =   405
      Left            =   360
      TabIndex        =   7
      Top             =   2700
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Subd. Cobzas"
      Height          =   405
      Left            =   360
      TabIndex        =   6
      Top             =   1740
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Subd. Vtas."
      Height          =   405
      Left            =   360
      TabIndex        =   5
      Top             =   2220
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Artículos"
      Height          =   405
      Left            =   360
      TabIndex        =   4
      Top             =   1260
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cta. Cte."
      Height          =   405
      Left            =   360
      TabIndex        =   3
      Top             =   780
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clientes"
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   405
      Left            =   3375
      TabIndex        =   2
      Top             =   1275
      Width           =   2685
   End
   Begin VB.Label Label1 
      Height          =   405
      Left            =   3375
      TabIndex        =   1
      Top             =   855
      Width           =   2685
   End
End
Attribute VB_Name = "PasaArch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Registro As String, Campo As String, I1%
Dim PClte As Integer, PCpbte As Integer, PFecha As Date, PTipo As Integer
Dim PSECC As String, PNRO As String, Secc, RgCLTE As New ADODB.Recordset


Private Sub Command1_Click()
Dim CIVA As String

Open "c:\Clientes.txt" For Input As #1

'   Vaciar la Tabla
miSQL1 = "delete from Clientes"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

Do Until EOF(1)

'Input #1, Registro
Registro = Input(296, #1)

Campo = Mid(Registro, 2, 5)

PClte = Mid(Registro, 2, 5)
CIVA = Mid(Registro, 182, 1)
 
Show
DoEvents

Label1.Caption = "Clte.=" & PClte & " Fecha=" & PFecha

miSQL1 = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

If Not RgTABL.EOF Then
    RgTABL.Close
    GoTo OtroReg
    End If

RgTABL.AddNew

RgTABL!CODIGO = PClte
RgTABL!Nomb = Mid(Registro, 10, 30)
RgTABL!Dir = Mid(Registro, 43, 25)
RgTABL!Loc = Mid(Registro, 71, 15)
RgTABL!CP = Mid(Registro, 89, 4)
RgTABL!PCIA = Mid(Registro, 96, 1)
RgTABL!tel1 = Mid(Registro, 100, 14)
RgTABL!TEL2 = Mid(Registro, 118, 1)
RgTABL!FAX = Mid(Registro, 121, 14)
RgTABL!EMAIL = Mid(Registro, 138, 1)
RgTABL!ZONA = Mid(Registro, 142, 2)
RgTABL!VEND = Mid(Registro, 147, 2)
RgTABL!CVTA = Mid(Registro, 152, 2)

RgTABL!dto1 = Mid(Registro, 157, 2)
RgTABL!dto2 = Mid(Registro, 162, 2)
RgTABL!dto3 = Mid(Registro, 167, 2)
RgTABL!dto4 = Mid(Registro, 172, 2)
RgTABL!dto5 = Mid(Registro, 177, 2)

RgTABL!CIVA = Mid(Registro, 182, 1)

RgTABL!Cuit = Mid(Registro, 186, 11)
RgTABL!NGAN = Mid(Registro, 200, 11)
RgTABL!IB = Mid(Registro, 214, 11)

RgTABL!corr1 = Mid(Registro, 228, 3)
RgTABL!CORR2 = Mid(Registro, 234, 3)
RgTABL!CANAL = Mid(Registro, 240, 2)
RgTABL!PROIND = Mid(Registro, 245, 6)
RgTABL!FALTA = CDate(Mid(Registro, 284, 10))
RgTABL!FACTUAL = Date

RgTABL!CREDIT = CCur(Mid(Registro, 254, 12))
RgTABL!DEUDA = CCur(Mid(Registro, 269, 12))

RgTABL!USUARIO = "Pasaje"

RgTABL!CIVA = 1
If CIVA = 2 Then RgTABL!CIVA = "3"
If CIVA = 3 Then RgTABL!CIVA = "4"
If CIVA = 4 Then RgTABL!CIVA = "2"

    RgTABL.Update

I1 = I1 + 1
Label2.Caption = "Grabados : " & I1
    
OtroReg:
    
    RgTABL.Close

Loop

Label1.Caption = ""
Label2.Caption = ""

Close #1
End Sub

Private Sub Command2_Click()
Dim SalAnt1 As Currency, SalAnt2 As Currency, PriFecha As Date

Open "c:\Ctasctes.txt" For Input As #1

'   Vaciar la Tabla
miSQL1 = "delete from CtasCtes"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

Do Until EOF(1)

'Input #1, Registro
Registro = Input(156, #1)

Campo = Mid(Registro, 2, 5)

PClte = Mid(Registro, 2, 5)
PFecha = Mid(Registro, 10, 10)
PCpbte = Mid(Registro, 38, 8)
PTipo = Mid(Registro, 27, 1)

Show
DoEvents

Label1.Caption = "Clte.=" & PClte & " Fecha=" & PFecha

miSQL1 = "SELECT * FROM CtasCtes WHERE CLTE = " & PClte & " AND CPBTE= " & PCpbte & _
                                " AND FECHA = " & Separa & PFecha & Separa & " AND TIPO= " & PTipo
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

If Not RgTABL.EOF Then
    RgTABL.Close
    GoTo OtroReg
    End If

RgTABL.AddNew

RgTABL!prefijo = Mid(Registro, 31, 4)
RgTABL!Cpbte = PCpbte
RgTABL!FECHA = PFecha
RgTABL!TIPO = PTipo
RgTABL!CLTE = PClte


RgTABL!imput1 = Mid(Registro, 49, 2)
RgTABL!imput2 = Mid(Registro, 54, 2)
RgTABL!imput3 = Mid(Registro, 59, 2)
RgTABL!imput4 = Mid(Registro, 64, 2)
RgTABL!imput5 = Mid(Registro, 69, 2)
RgTABL!imput6 = Mid(Registro, 74, 2)

RgTABL!DEBE = CCur(Mid(Registro, 79, 12))
RgTABL!Impte = CCur(Mid(Registro, 94, 12))
RgTABL!bon = CCur(Mid(Registro, 109, 12))

RgTABL!USUAR = "Pasaje"
RgTABL!CVTA = Mid(Registro, 137, 2)
RgTABL!tipo9 = Mid(Registro, 142, 1)
RgTABL!moti = Mid(Registro, 146, 2)
RgTABL!FECVTO = Mid(Registro, 124, 10)

RgTABL!Letra = "A"
If PTipo = 4 Then RgTABL!Letra = "X"

    RgTABL.Update

I1 = I1 + 1
Label2.Caption = "Grabados : " & I1
    
OtroReg:
    
    RgTABL.Close

Loop

Label1.Caption = ""
Label2.Caption = ""

Close #1

'-----------------------------------------------------------
'                  Arreglo de Saldos

miSQL1 = "SELECT * FROM Clientes"
RgCLTE.Open miSQL1, dbTABL, adOpenForwardOnly, adLockReadOnly

If RgCLTE.EOF Then
    RgCLTE.Close
    GoTo OtroReg3
    End If

Do Until RgCLTE.EOF

PClte = RgCLTE!CODIGO

    miSQL1 = "SELECT * FROM CtasCtes WHERE CLTE = " & PClte & _
    " AND DEBE > 0  ORDER BY FECHA, CPBTE "
    
    RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic
    
    If RgTABL.EOF Then
        GoTo AGrabar
        End If
        
    PriFecha = RgTABL!FECHA
    SalAnt1 = 0: SalAnt2 = 0
    
    Do Until RgTABL.EOF
    
    SalAnt1 = SalAnt1 + RgTABL!DEBE
    SalAnt2 = SalAnt2 + RgTABL!Impte
    
    RgTABL.MoveNext
    
    Loop

AGrabar:
    
    If SalAnt1 = SalAnt2 Then
        RgTABL.Close
        GoTo OtroReg3
        End If

    RgTABL.AddNew
    
    RgTABL!prefijo = 0
    RgTABL!Cpbte = 0
    RgTABL!FECHA = PriFecha
    RgTABL!TIPO = 0
    RgTABL!CLTE = PClte
    
    
    RgTABL!imput1 = 0
    RgTABL!imput2 = 0
    RgTABL!imput3 = 0
    RgTABL!imput4 = 0
    RgTABL!imput5 = 0
    RgTABL!imput6 = 0
    
    RgTABL!DEBE = 0
    RgTABL!Impte = SalAnt1 - SalAnt2
    RgTABL!bon = 0
    
    RgTABL!USUAR = "Pasaje"
    RgTABL!CVTA = 0
    RgTABL!tipo9 = 0
    RgTABL!moti = 0
    RgTABL!FECVTO = Date
    
    RgTABL!Letra = "X"

    RgTABL.Update
    RgTABL.Close

SalAnt1 = 0: SalAnt2 = 0

Label2.Caption = "Cliente : " & PClte

Show
DoEvents

OtroReg3:

RgCLTE.MoveNext

Loop

    RgCLTE.Close

Label1.Caption = ""
Label2.Caption = ""


End Sub

Private Sub Command3_Click()

Open "c:\Articulo.txt" For Input As #1

'   Vaciar la Tabla
miSQL1 = "delete from Articulo"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

Do Until EOF(1)

'Input #1, Registro
Registro = Input(202, #1)

Campo = Mid(Registro, 2, 10)

PSECC = Mid(Registro, 2, 5)
PNRO = Mid(Registro, 10, 10)
 
Show
DoEvents

Label1.Caption = "Sección=" & PSECC & " Código=" & PNRO

miSQL2 = "SELECT  * FROM Articulo WHERE COD1 = '" & PSECC & _
                     "' and COD2 = '" & PNRO & "'"
RgTABL.Open miSQL2, dbTABL, adOpenDynamic, adLockPessimistic

If Not RgTABL.EOF Then
    RgTABL.Close
    GoTo OtroReg
    End If

RgTABL.AddNew

RgTABL!COD1 = PSECC
RgTABL!Cod2 = PNRO
RgTABL!Descri = Mid(Registro, 23, 30)
RgTABL!Stock = Mid(Registro, 56, 10)
RgTABL!prec = Mid(Registro, 69, 10)
RgTABL!vta1 = Mid(Registro, 82, 10)
RgTABL!vta2 = Mid(Registro, 95, 10)
RgTABL!PCos = Mid(Registro, 108, 10)
RgTABL!dto1 = Mid(Registro, 121, 10)
RgTABL!Stmin = Mid(Registro, 134, 10)
RgTABL!mcalis = 0
Campo = Mid(Registro, 151, 10)
RgTABL!Stant = Mid(Registro, 151, 10)
RgTABL!gptipo = Mid(Registro, 164, 2)
RgTABL!gpdsd = Mid(Registro, 169, 3)
RgTABL!gphst = Mid(Registro, 175, 3)

RgTABL!FACTUAL = Mid(Registro, 190, 10)

RgTABL!USUARIO = "Pasaje"

    RgTABL.Update

I1 = I1 + 1
Label2.Caption = "Grabados : " & I1
    
OtroReg:
    
    RgTABL.Close

Loop

Label1.Caption = ""
Label2.Caption = ""

Close #1
End Sub

Private Sub Command4_Click()

Label1.Caption = ""
Label2.Caption = ""

'   Vaciar la Tabla
miSQL1 = "delete from FCIVAVta"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

'   Vaciar la Tabla
miSQL1 = "delete from FCEstad1"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

End Sub

Private Sub Command5_Click()

Label1.Caption = ""
Label2.Caption = ""

'   Vaciar la Tabla
miSQL1 = "delete from Imputacion"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic
    
End Sub

Private Sub Command6_Click()

Label1.Caption = ""
Label2.Caption = ""

'   Vaciar la Tabla
miSQL1 = "delete from Cheques"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic
 
End Sub

Private Sub Command7_Click()

Unload Me


End Sub

Private Sub Command8_Click()

miSQL1 = "SELECT * FROM Articulo WHERE COD1 = 'PISO'"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

If RgTABL.EOF Then
    RgTABL.Close
    Exit Sub
    End If
    
Do Until RgTABL.EOF
    
RgTABL!Cod2 = Trim(Mid(RgTABL!Cod2, 5, 1))
    
    
'RgTABL.Update
    
RgTABL.MoveNext
    
Loop

RgTABL.Close
    
End Sub

Private Sub Command9_Click()
Dim CIVA As String, i2%

Open "c:\Clientes.txt" For Input As #1

'   Vaciar la Tabla
miSQL1 = "delete from DtoxClte"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

Do Until EOF(1)

'Input #1, Registro
Registro = Input(296, #1)

Campo = Mid(Registro, 2, 5)

PClte = Mid(Registro, 2, 5)
 
'Show
'DoEvents

Label1.Caption = "Clte.=" & PClte

For i2 = 0 To 7

miSQL1 = "SELECT * FROM DtoxClte WHERE CLTE = " & PClte & " And SECCION = '" & Secc(i2) & "'"
RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockPessimistic

If Not RgTABL.EOF Then
    RgTABL.Close
    GoTo OtroReg2
    End If

RgTABL.AddNew

RgTABL!CLTE = PClte
RgTABL!SECCION = Secc(i2)


RgTABL!dto1 = Mid(Registro, 157, 2)
RgTABL!dto2 = Mid(Registro, 162, 2)
RgTABL!dto3 = Mid(Registro, 167, 2)
RgTABL!dto4 = Mid(Registro, 172, 2)
RgTABL!dto5 = Mid(Registro, 177, 2)
RgTABL!Rgo1 = 0
RgTABL!Rgo2 = 0
RgTABL!Rgo3 = 0

RgTABL!FACTUAL = Date
RgTABL!USUARIO = "Pasaje"

    RgTABL.Update
 
Show
DoEvents

I1 = I1 + 1
Label2.Caption = "Grabados : " & I1
    
OtroReg:
    
    RgTABL.Close

Next i2

Loop

OtroReg2:

Label1.Caption = ""
Label2.Caption = ""

Close #1
End Sub
Private Sub Form_Load()

Secc = Array("A    ", "B    ", "C    ", "D    ", "E    ", "O    ", "CAA  ", "CAC  ")

End Sub
