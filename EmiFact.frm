VERSION 5.00
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form EmiFact 
   BackColor       =   &H00FEFADE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                              Emisión del Comprobante"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10815
   ClipControls    =   0   'False
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
   ScaleHeight     =   6900
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1965
      Left            =   7950
      ScaleHeight     =   1905
      ScaleWidth      =   2025
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3870
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.PictureBox Picture1 
      Height          =   1080
      Left            =   165
      Picture         =   "EmiFact.frx":0000
      ScaleHeight     =   1020
      ScaleWidth      =   3345
      TabIndex        =   18
      Top             =   5520
      Visible         =   0   'False
      Width           =   3405
   End
   Begin SHDocVwCtl.WebBrowser WB1 
      Height          =   3570
      Left            =   6150
      TabIndex        =   16
      Top             =   765
      Width           =   4500
      ExtentX         =   7937
      ExtentY         =   6297
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "Emitir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Realiza conexión con AFIP y emite el comprobante"
      Top             =   4395
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ver Cpbte."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   300
      TabIndex        =   0
      ToolTipText     =   "Confecciona la FACTURA PROVISORIA "
      Top             =   4395
      Width           =   1140
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FEFADE&
      Caption         =   "Facturación Electrónica"
      Height          =   270
      Left            =   2910
      TabIndex        =   4
      Top             =   105
      Value           =   1  'Checked
      Width           =   2925
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00F9FADC&
      Caption         =   " Porcentaje de I.B."
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
      Height          =   615
      Left            =   345
      TabIndex        =   12
      Top             =   1395
      Width           =   5445
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   930
         MaxLength       =   90
         TabIndex        =   6
         Top             =   255
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   2475
         TabIndex        =   13
         Top             =   270
         Visible         =   0   'False
         Width           =   2790
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00F9FADC&
      Caption         =   "  Comprobante a Emitir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1005
      Left            =   330
      TabIndex        =   10
      Top             =   2055
      Width           =   5460
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   600
         TabIndex        =   11
         Top             =   330
         Width           =   4365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F9FADC&
      Caption         =   "  Notas al Pie  "
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
      Height          =   1185
      Left            =   330
      TabIndex        =   9
      Top             =   3105
      Width           =   5445
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   915
         MaxLength       =   90
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "EmiFact.frx":13D6
         Top             =   270
         Width           =   3795
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4575
      TabIndex        =   2
      Top             =   4395
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8355
      TabIndex        =   5
      Top             =   6180
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F9FADC&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   5460
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Verifique que esté todo en condiciones para emitir el comprobante"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   225
         TabIndex        =   8
         Top             =   270
         Width           =   5055
      End
   End
   Begin VSPDF8LibCtl.VSPDF8 VSPDF81 
      Left            =   6360
      Top             =   4395
      Author          =   ""
      Creator         =   ""
      Title           =   ""
      Subject         =   ""
      Keywords        =   ""
      Compress        =   3
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Estado de la Conexión con AFIP"
      Height          =   300
      Left            =   6540
      TabIndex        =   15
      Top             =   465
      Width           =   3630
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   240
      TabIndex        =   14
      Top             =   60
      Width           =   2655
   End
End
Attribute VB_Name = "EmiFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NROFAC As Long, NoAnduvo As Boolean, AFIPbien As Boolean, Corr As Integer
Dim LaFact As String, TIPO As String
Dim Impte As Currency, TotIVAIns As Currency, TotIVANI As Currency, TotCan As Single
Dim SumDtos As Currency, TotDtos As Currency, TotBruto As Currency, TOTNETO As Currency
Dim TotSIVA As Currency, CodFactAFIP
Dim AFIPTotal As Currency, AFIPNeto As Currency, AFIPIVA As Currency, AFIPCUIT As String
Dim FecVtoCAE, NroCAE, MotRech, PtoVtaCpbte As String * 4
Dim ValIB As Currency, TotIB As String, CUITdelCliente As String

Dim dbPADR As New ADODB.Connection, RgPADR As New ADODB.Recordset

Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset
Dim dbCCTE As New ADODB.Connection, RgCCTE As New ADODB.Recordset
Dim dbFACT As New ADODB.Connection, RgFACT As New ADODB.Recordset
Dim dbSTOK As New ADODB.Connection, RgSTOK As New ADODB.Recordset
Dim RgTOTL As New ADODB.Recordset

Dim cQrCode As ClsQrCode  '   Para código QR
Dim QRunaVez As Boolean
Dim pngClass As New LoadPNG
Dim WithEvents Bas64 As Base64
Attribute Bas64.VB_VarHelpID = -1
Dim encBuffer As String
Dim strBuffer As String
Dim flgFile As Integer
Dim byteBuffer() As Byte

Const LeyenDol1 = "LA PRESENTE FACTURA SE ABONARA EN DOLARES ESTADOUNIDENSES BILLETES O " & _
                  "EN LA CANTIDAD DE PESOS"
Const LeyenDol2 = "SUFICIENTES PARA CANCELAR LA SUMA TOTAL EN DOLARES  AL TIPO DE CAMBIO" & _
                  " VENDEDOR DEL B.N.A. DEL"
Const LeyenDol3 = "MERCADO OFICIAL QUE RIJA PARA LAS IMPORTACIONES DE BIENES (O AQUEL QU" & _
                  "E EN EL FUTURO LO "
Const LeyenDol4 = "SUSTITUYA) DEL DIA ANTERIOR AL EFECTIVO PAGO.-------------------------" & _
                  "------------------"
Const LeyenDol5 = "A LOS EFECTOS IMPOSITIVOS EL TIPO DE CAMBIO APLICADO ES u$s 1= $ "


Private Sub Form_Load()

Me.Move 1600, 1500, 6315, 5200

    Set cQrCode = New ClsQrCode
    Set Bas64 = New Base64

'BarCod.Font.Name = "Arial"
'BarCod.Font.Size = 8
    
Command1.Enabled = False
Command4.Enabled = False

If TipoFac = 4 Then
    Check1.Value = 0
Else
    Check1.Value = 1
End If

If EnDolares Then
    Label4.Caption = "es EN DÓLARES"
Else
    Label4.Caption = ""
End If

NoAnduvo = False
Text1.Text = ""

'   - - -   Para la percepcion de II.BB.  - - -
'    Frame4.Enabled = False
'If ConPercep = 2 Then                    '  si la marca de parametro es SI
'    Frame4.Enabled = False
'Else
'    BDatos2 = "DSN=PadronARBA"
'    dbPADR.ConnectionString = BDatos2
'    dbPADR.Open
'    Frame4.Enabled = True
'End If

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

TotDtos = 0: TotBruto = 0: TotCan = 0
TOTNETO = 0: TotIVAIns = 0: TotIVANI = 0

Text2.Text = 0

PtoVtaCpbte = "0000"

Select Case TipoFac
    Case 1
        Label2.Caption = "Factura Nº "
        PtoVtaCpbte = "0003"
    Case 2
        Label2.Caption = "N. de Crédito Nº "
        PtoVtaCpbte = "0004"
    Case 3
        Label2.Caption = "N. de Débito Nº "
        PtoVtaCpbte = "0004"
    Case 4
        Label2.Caption = "Cotización Nº "
End Select

    miSQL1 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
    RgTABL.Open miSQL1, dbTABL, adOpenForwardOnly, adLockReadOnly
    
'  Busca Número de Factura
Select Case TipoFac

    Case 1
        If LaLetra = "A" Then
            Label2.Caption = Label2.Caption & RgTABL!nume1 + 1
        Else
            Label2.Caption = Label2.Caption & RgTABL!nume5 + 1
        End If

    Case 2
        If LaLetra = "A" Then
            Label2.Caption = Label2.Caption & RgTABL!nume1 + 1
        Else
            Label2.Caption = Label2.Caption & RgTABL!nume5 + 1
        End If
    
    Case 3
        If LaLetra = "A" Then
            Label2.Caption = Label2.Caption & RgTABL!nume1 + 1
        Else
            Label2.Caption = Label2.Caption & RgTABL!nume5 + 1
        End If

    Case 4
        Label2.Caption = Label2.Caption & RgTABL!nume4 + 1
    
End Select

RgTABL.Close

miSQL1 = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTEf
RgCLTE.Open miSQL1, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgCLTE.EOF Then
    MsgBox "Clientes", vbCritical + vbOKOnly, "Error"
    NoAnduvo = True
    RgCLTE.Close
    Exit Sub
    End If
    
    
ClteNomb = RgCLTE!Nomb
ClteDir = RgCLTE!Dir
ClteCPos = RgCLTE!CP
ClteCIVA = RgCLTE!CIVA
CltePCIA = RgCLTE!PCIA
ClteTel1 = RgCLTE!tel1
ClteLoc = RgCLTE!Loc
ClteCUIT = RgCLTE!Cuit
ClteCVta = RgCLTE!CVTA

    RgCLTE.Close
    
Exit Sub                 '   saltea IB
    
'-------------   MINIMO A CALCULAR   ----------
If PieFact.Label304.Caption < 50 Then
            Label3.Visible = True
            Label3.Caption = "No Corresponde (< a $50)"
    Exit Sub
End If
    
'   - - -   Para la percepcion de II.BB.  - - -
'If ConPercep = 1 Then                   '  Si la marca de IB en parametro es SI
'        Dim elcuit As Double
'        elcuit = Val(ClteCUIT)
'        miSQL1 = "SELECT * FROM Padron1 WHERE cuit = '" & Val(elcuit) & "'"
'        RgPADR.Open miSQL1, dbPADR, adOpenForwardOnly, adLockReadOnly
        
'        If RgPADR.EOF Then
'                If CltePCIA = "B " Then
'                    Label3.Visible = True
'                    Label3.Caption = "Cliente no está en el Padrón"
'                    Text2.Text = "6,00"
'                End If
'            RgPADR.Close
'            Exit Sub
'            End If
        
'        Text2.Text = Format(RgPADR!percep, "0.00")
'            RgPADR.Close
'End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    dbCLTE.Close
'   - - -   Para la percepcion de II.BB.  - - -
'If ConPercep = 1 Then dbPADR.Close                   '  Si la marca de IB en parametro es SI
   
End Sub

Private Sub Check1_Click()

              Command3.Caption = "Emite"
              Command4.Visible = False
              PtoVtaCpbte = PtoVta

If Check1.Value = 0 Then           ' si quitan la marca y no es cotización avisa
  If TipoFac <> 4 Then
          vbMsgBoxTitle = "Tipo de Factura"
          vbMsgBoxText = "Si quita la marca, se emitirá  " & _
            vbCrLf & "la factura en forma convencional (papel)." & _
            vbCrLf & "Sólo debe hacerlo si no puede emitir de" & _
            vbCrLf & "forma electrónica. " & _
            vbCrLf & vbCrLf & "Está seguro que desea continuar ?  "
          vbMsgBoxResp = vbYesNo + vbExclamation + vbApplicationModal + vbDefaultButton1
        
          vbResponse = MsgBox(vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle)
        
          Select Case vbResponse
        
                 Case vbYes
                      Command3.Caption = "Emite"
                      Command4.Visible = False
                      PtoVtaCpbte = PtoVta
                      Exit Sub
        
                 Case vbNo
                      Command3.Caption = "Ver Cpbte."
                      Command4.Visible = True
                      Check1.Value = 1
                      
          End Select
    End If
  
Else
        Command3.Caption = "Ver Cpbte."
        Command4.Visible = True
'        PtoVtaCpbte = "0003"

End If

End Sub

'--------------------------------------------------------------------
'
'                    I M P R E S I O N
'
'--------------------------------------------------------------------
'Imprimir
 Sub Command1_Click()

dbTABL.BeginTrans
dbCLTE.BeginTrans

If Check1.Value = 1 Then
    If AFIPbien And TipoFac <> 4 Then Graba
Else
    If Not SinCpbte And TipoFac <> 4 Then Graba
End If

'Graba
If NoAnduvo Then
    dbTABL.RollbackTrans
    dbCLTE.RollbackTrans
    MsgBox "Hubo ERRORES. No se Grabó el Comprobante", vbExclamation, " Grabación del Comprobante "

Else
    dbTABL.CommitTrans
    dbCLTE.CommitTrans
End If

End Sub

'Salir
Private Sub Command2_Click()

'FCMENU.Toolbar3.Buttons(3).Value = tbrUnpressed
Unload Me

End Sub

' Ver Comprobante
Private Sub Command3_Click()

If Check1.Value = 1 Then
    EsProvisorio = True
Else
    EsProvisorio = False
End If

    dbTABL.BeginTrans
    IMPRIME

If Check1.Value = 1 Then                    ' Si es fact.Electr. habilita botón
    Command4.Enabled = True
    dbTABL.RollbackTrans
Else
    dbTABL.CommitTrans
    Command1_Click
End If

Command4.Enabled = True

End Sub

' Emitir
Private Sub Command4_Click()
Dim Pagina As String

Command4.Enabled = False

    ConectaAFIP

' Salva el LOG en un archivito
Pagina = WB1.Document.Body.Innertext
Pagina = Replace(Pagina, "<br><br>", vbCrLf)
'    MsgBox "QTY: " & Qty & vbCrLf & "LastId: " & Lastid & vbCrLf & "LastCBTE:" & LastCbte & vbCrLf & "CAE: " & CAE, vbInformation + vbOKOnly
Open "c:\Factu\logs\" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "\" & _
                         TIPO & "-" & LaLetra & "-" & PtoVtaCpbte & "-" & NROFAC & ".txt" For Output As #1
Write #1, Pagina
Close #1

' Graba

    EsProvisorio = False

IMPRIME_Electronica
    
    Command1_Click

Command4.Enabled = True

End Sub

Sub IMPRIME()

Dim CharsDrawn As Long, I1%, i2%, i3%, I4%
Dim Left As Long, top As Long, right As Long, bottom As Long, Hcenter As Long, Vcenter As Long
Dim Desctos As String, Simbo As String, PVta As Currency
Dim Secc As String, Descri As String, Pulg As Single, Mtrs As Single, Cant As Single
Dim s$, h$, d$, EnLetras As String, TipoMoneda As String


        If ClteCIVA < 3 Or (ClteCIVA = 4 And CltePCIA = "V ") Then
            LaLetra = "A"
        Else
            LaLetra = "B"
        End If
        
        LaFact = ""
            
    If Check1.Value = 1 Then
        LaFact = "PROVISORIA"
            NROFAC = 0

        If LaLetra = "A" Then
            Select Case TipoFac
                Case 1:   CodFactAFIP = 1
                Case 2:   CodFactAFIP = 3
                Case 3:   CodFactAFIP = 2
            End Select
        Else
            Select Case TipoFac
                Case 1:   CodFactAFIP = 6
                Case 2:   CodFactAFIP = 8
                Case 3:   CodFactAFIP = 7
            End Select
        End If
      
      GoTo HastaAlla1
      
    End If

'---------------------------------------

'  Busca Número de Factura si es factura convencional (no es electrónica)

    miSQL1 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
    RgTABL.Open miSQL1, dbTABL, adOpenDynamic, adLockOptimistic
    If RgTABL.EOF Then
    MsgBox "Parametro", vbCritical + vbOKOnly, "Error"
        NoAnduvo = True
        RgTABL.Close
        Exit Sub
    End If

Select Case TipoFac
    Case 1
        If LaLetra = "A" Then
            NROFAC = RgTABL!nume1 + 1
            RgTABL!nume1 = NROFAC
        Else
            NROFAC = RgTABL!nume5 + 1
            RgTABL!nume5 = NROFAC
        End If

    Case 2
        If LaLetra = "A" Then
            NROFAC = RgTABL!nume1 + 1
            RgTABL!nume1 = NROFAC
        Else
            NROFAC = RgTABL!nume5 + 1
            RgTABL!nume5 = NROFAC
        End If
    
    Case 3
        If LaLetra = "A" Then
            NROFAC = RgTABL!nume1 + 1
            RgTABL!nume1 = NROFAC
        Else
            NROFAC = RgTABL!nume5 + 1
            RgTABL!nume5 = NROFAC
        End If

    Case 4
            NROFAC = RgTABL!nume4 + 1
            RgTABL!nume4 = NROFAC
    
End Select

        RgTABL.Update
        RgTABL.Close
        
'**********************************************
'              Evita la Impresión
'-------------------------------------------
' Exit Sub
        
'-------------------------------------------

HastaAlla1:

Select Case TipoFac
    Case 1
        TIPO = "FACTURA"
    Case 2
        TIPO = "NOTA DE CRÉDITO"
    Case 3
         TIPO = "NOTA DE DÉBITO"
    Case 4
         TIPO = "COTIZACIÓN"
         LaLetra = "X"
End Select
  
IMPRE.Caption = " Impresión de " & TIPO & " " & LaFact & " " & LaLetra & " Nº " & Format(NROFAC, "00000000") & " (Cta.Cte.)"

IMPRE.VP1.Copies = 2

IMPRE.VP1.PaperSize = pprA4
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
IMPRE.VP1.MarginLeft = "0mm"
IMPRE.VP1.MarginRight = "0mm"
        
'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginBottom = "1mm"
IMPRE.VP1.MarginHeader = "1mm"

If Check1.Value = 1 Then
    Escribe "", 20, 1, 0, 14, 50, "P R O V I S O R I A"
End If

If HOMO Then
        Escribe "Arial Black", 10, 1, 0, 40, 95, "HOMOLOGACIÓN / Prueba"
End If

'Primer Rectángulo
If TipoFac = 4 Then
        DibujaRect 50, 7, 78, 205, 30, 30
        DibujaLine 80, 7, 80, 205
        DibujaLine 88, 7, 88, 205
        DibujaLine 250, 7, 250, 205

        Escribe "ARIAL BLACK", 24, 1, 0, 30, 100, LaLetra
        Escribe "", 20, 1, 1, 30, 140, TIPO
        Escribe "", 12, 1, 0, 37, 141, "0001-" & Format(NROFAC, "00000000")
End If

'............................................................................
'Si es N/Cred ó Déb. Tacha el formulario y escribe el tipo de cpbte.
If Check1.Value = 0 Then
    If TipoFac = 2 Or TipoFac = 3 Then
        Escribe "", 12, 1, 0, 9, 130, "XXXXXXXXXXXXX"
        Escribe "", 12, 1, 0, 61, 145, "XXXXXXX"
        Escribe "", 12, 1, 0, 22, 130, TIPO
    End If
End If
'Si es N/Cred,  escribe a que factura esta aplicada
If TipoFac = 2 Then
    Escribe "", 10, 1, 1, 66, 160, "Aplicada a:"
    laLOGIK1 = ApPtoVta: laLOGIK2 = ApCpbte
    Escribe "", 10, 0, 1, 71, 145, "Factura Nº " & String(4 - Len(laLOGIK1), "0") & _
             laLOGIK1 & "-" & String(8 - Len(laLOGIK2), "0") & laLOGIK2
    Escribe "", 10, 0, 1, 76, 150, " del  " & Format(ApFecha, "ddd. dd mmm yyyy")
End If
'-------------------------------------------------------------------------------

Escribe "", 12, 0, 0, 53, 164, Format(Date, "ddd dd MMM YYYY")
'Escribe "", 12, 0, 0, 51, 153, Format(Time, "hh:mm")

'---------------------------------------------------------------------------
Escribe "", 10, 0, 0, 54, 10, ClteNomb & "         (" & CodCLTEf & ")     "
Escribe "", 10, 0, 0, 59, 10, ClteDir
'Cod.Post., Loc.,Pcia.
I4 = Asc(CltePCIA)
Escribe "", 10, 0, 0, 64, 10, "(" & Mid(ClteCPos, 1, 4) & ")  " & ClteLoc & Space(2) & _
                                Mid(PCIA(I4 - 65), 2, 10)

'Cond.IVA, CUIT
Escribe "", 10, 0, 0, 69, 10, "IVA " & CabFact.Combo1.List(ClteCIVA - 1) & _
                              "     CUIT : " & ClteCUIT

'Cond.Vta,Fma.Envio,Remito
Escribe "", 10, 0, 0, 74, 10, "Condición de Venta : " & _
                                Mid(CabFact.Label104, 27, 40)
'----------------------------------------------------------------------------

'  Ttulo
'If TipoFac = 4 Then
    Escribe "", 10, 1, 1, 81, 12, "Cant."
    Escribe "", 10, 1, 1, 81, 56, "D e t a l l e"
    Escribe "", 10, 1, 1, 81, 83, "Despacho"
    Escribe "", 10, 1, 1, 81, 109, "% Descuento"
    Escribe "", 10, 1, 1, 81, 145, "Precio Unit."
    Escribe "", 10, 1, 1, 81, 178, "Importe "
'End If

    IMPRE.VP1.TableBorder = tbNone

s$ = "^+22mm|>+1mm|<+88mm|>+27mm|>+23mm|>+33mm"
h$ = "Cant.|       DETALLE      | % Descuento  | Precio  | Importe"

With IMPRE.VP1
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Lucida Console"
    .CurrentY = "92mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "4mm"
 End With

'IMPRE.VP1.EndDoc
 
'IMPRE.Show 1

TotDtos = 0: TotBruto = 0: TotSIVA = 0

If Motivo = 1 Then GoTo ConArticulos
'***********************************************************************************
'............................................................
' Para N/Credito o Debito sin detalle - Motivo mayor a 1
'............................................................

If DetNC.Text1(0).Text <> "" Then
    d$ = "" & "||" & DetNC.Text1(0).Text & "|" & "" & "|" _
          & "" & "|" & Format(DetNC.Text2(0).Text, " ###,##0.00")
    IMPRE.VP1.FontSize = 9
    IMPRE.VP1.AddTable s$, h$, d$, , , True
    IMPRE.VP1.FontSize = 3
    IMPRE.VP1.AddTable s$, h$, "|||||", , , True
        If Val(DetNC.Text2(0).Text) > 0 Then
            If DetNC.Check1(0).Value = 1 Then
               TotBruto = TotBruto + DetNC.Text2(0).Text
            Else
               TotSIVA = TotSIVA + DetNC.Text2(0).Text
            End If
        End If
End If

If DetNC.Text1(1).Text <> "" Then
    d$ = "" & "||" & DetNC.Text1(1).Text & "|" & "" & "|" _
          & "" & "|" & Format(DetNC.Text2(1).Text, " ###,##0.00")
    IMPRE.VP1.FontSize = 9
    IMPRE.VP1.AddTable s$, h$, d$, , , True
    IMPRE.VP1.FontSize = 3
    IMPRE.VP1.AddTable s$, h$, "|||||", , , True
        If Val(DetNC.Text2(1).Text) > 0 Then
            If DetNC.Check1(1).Value = 1 Then
               TotBruto = TotBruto + DetNC.Text2(1).Text
            Else
               TotSIVA = TotSIVA + DetNC.Text2(1).Text
            End If
        End If
End If

If DetNC.Text1(2).Text <> "" Then
    d$ = "" & "||" & DetNC.Text1(1).Text & "|" & "" & "|" _
          & "" & "|" & Format(DetNC.Text2(1).Text, " ###,##0.00")
    IMPRE.VP1.FontSize = 9
    IMPRE.VP1.AddTable s$, h$, d$, , , True
    IMPRE.VP1.FontSize = 3
    IMPRE.VP1.AddTable s$, h$, "|||||", , , True
        If Val(DetNC.Text2(2).Text) > 0 Then
            If DetNC.Check1(2).Value = 1 Then
               TotBruto = TotBruto + DetNC.Text2(2).Text
            Else
               TotSIVA = TotSIVA + DetNC.Text2(2).Text
            End If
        End If
End If

TotDtos = 0



GoTo SaleTotal
'***********************************************************************************
'............................................................
' Detalle con Artículos
'............................................................

ConArticulos:

For I1 = 1 To DetFact.FG1.Rows

If DetFact.FG1.TextMatrix(I1, 8) = "" Then GoTo SaleTotal

Secc = DetFact.FG1.TextMatrix(I1, 0)
'Pulg = CCur(DetFact.FG1.TextMatrix(I1, 1))
'Mtrs = CCur(DetFact.FG1.TextMatrix(I1, 2))
Cant = Val(DetFact.FG1.TextMatrix(I1, 6))
PVta = CCur(DetFact.FG1.TextMatrix(I1, 7))
Impte = DetFact.FG1.TextMatrix(I1, 8)

'-----------------------------------------------------------------------
' Si la Sección tiene marca usa la descr. de la Sección, sino usa la del Artículo
If DetFact.FG1.TextMatrix(I1, 11) = True Then
        'Arma la Descripción
        'Descri = DetFact.FG1.TextMatrix(i1, 5) & " " & MatrItems(i1, 1, 2) & " " & MatrItems(i1, 1, 1) & _
        '         " " & MatrItems(i1, 2, 2) & " " & MatrItems(i1, 2, 1) & _
        '         " " & MatrItems(i1, 3, 2) & " " & MatrItems(i1, 3, 1)
        
        Descri = DetFact.FG1.TextMatrix(I1, 5)
        If MatrItems(I1, 1, 2) <> "" And Val(MatrItems(I1, 1, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 1, 2) & " " & MatrItems(I1, 1, 1)
        End If
                 
        If MatrItems(I1, 2, 2) <> "" And Val(MatrItems(I1, 2, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 2, 2) & " " & MatrItems(I1, 2, 1)
        End If
                 
        If MatrItems(I1, 3, 2) <> "" And Val(MatrItems(I1, 3, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 3, 2) & " " & MatrItems(I1, 3, 1)
        End If
  '----------------
                 
        If MatrItems(I1, 4, 2) <> "" And Val(MatrItems(I1, 4, 2)) > 0 Then
            Descri = Descri & " - " & MatrItems(I1, 4, 2) & " " & MatrItems(I1, 4, 1)
        End If
                 
        If MatrItems(I1, 5, 2) <> "" And Val(MatrItems(I1, 5, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 5, 2) & " " & MatrItems(I1, 5, 1)
        End If
                 
        If MatrItems(I1, 6, 2) <> "" And Val(MatrItems(I1, 6, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 6, 2) & " " & MatrItems(I1, 6, 1)
        End If
        
Else
        Descri = DetFact.FG1.TextMatrix(I1, 5)
End If

'-----------------------------------------------------------------------
'  Arma la cantidad
'laLOGIK1 = MatrItems(I1, 4, 2) & " " & MatrItems(I1, 4, 1) & "x " & MatrItems(I1, 5, 2) & " " & MatrItems(I1, 5, 1)
laLOGIK1 = MatrItems(I1, 7, 2)
'If MatrItems(i1, 5, 2) <> "" Then laLOGIK1 = laLOGIK1 & "x" & MatrItems(i1, 5, 2)
'If MatrItems(i1, 6, 2) <> "" Then laLOGIK1 = laLOGIK1 & "x" & MatrItems(i1, 6, 2)
'-----------------------------------------------------------------------------------------

' Obtiene los Descuentos del Cliente para esa seccion/rubro
DxClte CodCLTEf, Secc
Simbo = "": Desctos = ""
SumDtos = 0
For i2 = 1 To 5
    If Dtos(i2) > 0 Then
        Desctos = Desctos & Simbo & Dtos(i2)
        SumDtos = SumDtos + ((Impte - SumDtos) * (Dtos(i2) / 100))
    End If
    Simbo = "+"
Next i2

If DetFact.FG1.TextMatrix(I1, 13) <> "" Then
    If Len(Descri) < 25 Then
     Descri = Mid(Descri, 1, 25) & Space(25 - Len(Descri))
    Else
     Descri = Mid(Descri, 1, 25)
    End If
    Descri = Descri & DetFact.FG1.TextMatrix(I1, 13)
End If

'd$ = Format(Cant, "##,##0") & "| |" & Descri & "|" & Desctos & "|" _
'          & Format(PVta, " ###,##0.00") & "|" & Format(Impte, " ###,##0.00")
d$ = laLOGIK1 & "||" & Descri & "|" & Desctos & "|" _
          & Format(PVta, " ###,##0.00") & "|" & Format(Impte, " ###,##0.00")

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 3
IMPRE.VP1.AddTable s$, h$, "|||||", , , True

TotDtos = TotDtos + SumDtos
TotBruto = TotBruto + Impte

Next I1

'...............................................................................
'       Fin Detalle con Artículos
'...............................................................................
'***********************************************************************************

SaleTotal:
        
If EnDolares Then
        TipoMoneda = "Son dólares"
        Escribe "Lucida Console", 7, 0, 0, 195, 10, LeyenDol1
        Escribe "Lucida Console", 7, 0, 0, 199, 10, LeyenDol2
        Escribe "Lucida Console", 7, 0, 0, 203, 10, LeyenDol3
        Escribe "Lucida Console", 7, 0, 0, 207, 10, LeyenDol4
        Escribe "Lucida Console", 7, 0, 0, 211, 10, LeyenDol5 & Format(LaCotiz, "###,##0.00")
        Escribe "Arial", 8, 1, 0, 250.5, 22, "U$S"
        Escribe "Arial", 8, 1, 0, 250.5, 72, "U$S"
        Escribe "Arial", 8, 1, 0, 250.5, 188, "U$S"
Else
        TipoMoneda = "Son pesos"
End If

'      Leyenda al Pie
i3 = InStr(1, Text1.Text, vbCrLf)
If i3 > 0 Then
        Escribe "Lucida Console", 12, 1, 1, 221, 10, Mid(Text1.Text, 1, i3 - 1)
Else
        Escribe "Lucida Console", 12, 1, 1, 221, 10, Mid(Text1.Text, 1, 30)
        i3 = 30
End If

I4 = InStr(i3 + 2, Text1.Text, vbCrLf)
If I4 > 0 Then
        Escribe "Lucida Console", 12, 1, 1, 225, 10, Mid(Text1.Text, i3 + 2, (I4 - 1) - i3)
Else
        Escribe "Lucida Console", 12, 1, 1, 225, 10, Mid(Text1.Text, i3 + 2, 30)
        I4 = i3 + 32
End If

i3 = InStr(I4 + 2, Text1.Text, vbCrLf)
If i3 > 0 Then
        Escribe "Lucida Console", 12, 1, 1, 229, 10, Mid(Text1.Text, I4 + 2, (i3 - 1) - I4)
Else
        Escribe "Lucida Console", 12, 1, 1, 229, 10, Mid(Text1.Text, I4 + 2, 30)
End If

Impte = TotBruto + TotSIVA

If TipoFac = 4 Then

    DibujaRect 260, 7, 268, 37, 30, 30
    DibujaRect 260, 39, 268, 68, 30, 30
    DibujaRect 260, 70, 268, 100, 30, 30
    DibujaRect 260, 102, 268, 132, 30, 30
'    DibujaRect 260, 134, 268, 164, 30, 30

    IMPRE.VP1.BrushStyle = bsSolid
    IMPRE.VP1.BrushColor = &HE0E0E0
    DibujaRect 260, 166, 268, 202, 30, 30
    IMPRE.VP1.BrushStyle = 1
    
    Escribe "Arial", 10, 0, 0, 255, 11, "Subtotal"
    Escribe "Arial", 10, 0, 0, 255, 43, "Descuento"
    Escribe "Arial", 10, 1, 0, 255, 75, "Subtotal"
    Escribe "Arial", 10, 0, 0, 255, 106, "IVA Insc. 21%"
'    Escribe "Arial", 10, 0, 0, 255, 135, ""
    Escribe "Arial", 10, 1, 1, 255, 179, "TOTAL"
Else
    Escribe "", 10, 1, 0, 231, 185, Format(Impte, "###,##0.00")
    Escribe "Arial", 10, 0, 0, 250.5, 112, "21"
    
End If
   
    
 IMPRE.VP1.StartTable

'-------------------------------------------------------------------------------
'................................................................................
'  imprime Totales

ValIB = 0

TotBruto = TotBruto - TotDtos

If Motivo <> 3 Then
    If ClteCIVA < 3 Then
        ValIB = Round(TotBruto * (CCur(Text2.Text) / 100), 2)
        TotIVAIns = TotBruto * IVAIns
        If ClteCIVA = 2 Then
           TotIVANI = TotBruto * IVANI
        End If
    End If
Else
'  Es nota de crédito o deb. por IB
    If ClteCIVA < 3 Then
        ValIB = TotSIVA
        TotIVAIns = 0: TotSIVA = 0: Impte = 0
    End If
End If

If ConPercep <> 1 Then ValIB = 0                 ' sin no tiene marca en parametro pone total en 0

d$ = Format(Impte, "###,##0.00") & "|"

TOTNETO = TotBruto + TotIVAIns + TotIVANI + TotSIVA + ValIB

With IMPRE.VP1
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Arial"
    .CurrentY = "262mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "3mm"
 End With


s$ = ">+33mm|>+30mm|>+33mm|>+33mm|>+32mm|>+38mm"
h$ = "subt1|dtos|subt2|ivains|ivani|total"

d$ = d$ & Format(TotDtos, "###,###,##0.00") & "|" & Format(TotBruto, "###,###,##0.00") & _
    "|" & Format(TotIVAIns, "###,##0.00") & "|" & Format(ValIB, "###,###.##") & _
    "|" & Format(TOTNETO, "###,###,##0.00")

IMPRE.VP1.FontSize = 10
IMPRE.VP1.AddTable s$, "", d$, , , True
IMPRE.VP1.TableCell(tcFontBold, 1, 3) = True
IMPRE.VP1.TableCell(tcFontBold, 1, 6) = True
 
 IMPRE.VP1.EndTable
 
    If ValIB > 0 Then
        DibujaRect 260, 134, 268, 164, 30, 30
        Escribe "Arial", 10, 0, 0, 255, 135, "Perc. II.BB."
    End If

        Escribe "", 7, 1, 1, 269, 163, "(" & NROFAC & ")"
   
' Importe en letras
EnLetras = MontoEscrito(TOTNETO)
Escribe "Lucida Console", 7, 0, 0, 236, 30, TipoMoneda & Mid$(EnLetras, 1, 70)
Escribe "Lucida Console", 7, 0, 0, 240, 30, Mid$(EnLetras, 71, 70)
Escribe "Lucida Console", 7, 0, 0, 244, 30, Mid$(EnLetras, 141, 70)
    
 IMPRE.VP1.EndDoc
'IMPRE.Show 1
Dim pepe

'Genera PDF
VSPDF81.Title = "c:\Factu\Comprob\" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "\" & TIPO & "-" & LaLetra & "-" & PtoVtaCpbte & "-" & NROFAC & ".pdf"
VSPDF81.Title = TIPO & LaFact & "-" & LaLetra & "-" & PtoVtaCpbte & "-"

If Check1.Value = 1 Then
    VSPDF81.Title = VSPDF81.Title & "Hora " & Format(Time, "hh mm") & "-"
Else
    VSPDF81.Title = VSPDF81.Title & NROFAC
End If

VSPDF81.ConvertDocument IMPRE.VP1, "c:\Factu\Comprob\" & _
                         Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "\" & _
                         VSPDF81.Title & ".pdf"
 
'Muestra PDF
PDF.PDF1.LoadFile "c:\Factu\Comprob\" & _
                         Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "\" & _
                         VSPDF81.Title & ".pdf"
PDF.Caption = VSPDF81.Title
PDF.Show
PDF.PDF1.setShowToolbar True

End Sub

Sub IMPRIME_Electronica()

Dim CharsDrawn As Long, I1%, i2%, i3%, I4%
Dim Left As Long, top As Long, right As Long, bottom As Long, Hcenter As Long, Vcenter As Long
Dim Desctos As String, Simbo As String, PVta As Currency
Dim Secc As String, Descri As String, Pulg As Single, Mtrs As Single, Cant As Single
Dim s$, h$, d$, EnLetras As String, TipoMoneda As String
Dim Copias As Single

If NroCAE = "" Or TipoFac = 4 Then Exit Sub

        If ClteCIVA < 3 Or (ClteCIVA = 4 And CltePCIA = "V ") Then
            LaLetra = "A"
        Else
            LaLetra = "B"
        End If
        
'**********************************************
'              Evita la Impresión
'-------------------------------------------
' Exit Sub
        
'-------------------------------------------

HastaAlla1:

Copias = 0

Select Case TipoFac
    Case 1
        TIPO = " F A C T U R A"
    Case 2
        TIPO = "NOTA DE CRÉDITO"
    Case 3
        TIPO = "NOTA DE DÉBITO"
End Select
  
IMPRE.Caption = " Impresión de " & TIPO & " " & LaFact & " " & LaLetra & " Nº " & Format(NROFAC, "00000000") & " (Cta.Cte.)"

IMPRE.VP1.Copies = 2

IMPRE.VP1.PaperSize = pprA4
IMPRE.VP1.Visible = True
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
IMPRE.VP1.MarginLeft = "5mm"
        
'...........................................
'   Comienza el Documento
IMPRE.VP1.StartDoc
IMPRE.VP1.MarginRight = "5mm"
IMPRE.VP1.MarginBottom = "1mm"
IMPRE.VP1.MarginHeader = "1mm"
    
 
Comienza1:
 
 Copias = Copias + 1
    
 IMPRE.VP1.DrawPicture Picture1, "10mm", "5mm"
 
Escribe "ARIAL BLACK", 20, 1, 0, 14, 70, "S.R.L."
Escribe "ARIAL", 9, 1, 0, 23, 20, "    CINTAS Y CORREAS DE TRANSMISIÓN"
Escribe "ARIAL", 9, 1, 0, 26.5, 20, "AUTOMOTORES - AGRÍCOLAS - INDUSTRIALES"
Escribe "ARIAL", 9, 1, 0, 30, 20, "    GOMAS - ARTÍCULOS - ACCESORIOS"

Escribe "ARIAL", 10, 1, 0, 35, 20, "Av. Crovara 2948 - (1766) La Tablada (BA)"
Escribe "ARIAL", 10, 1, 0, 39, 18, "Tel/Fax (011) 4652-1040 / 2684 / 2689 / 3080"
Escribe "ARIAL", 7, 1, 0, 42, 30, "email: correas@alestel.com.ar"
Escribe "ARIAL", 11, 0, 0, 45, 18, "I.V.A. RESPONSABLE INSCRIPTO"

Select Case Copias
    Case 1:    Escribe "ARIAL", 9, 1, 0, 45, 100, "O R I G I N A L"
    Case 2:    Escribe "ARIAL", 9, 1, 0, 45, 100, "D U P L I C A D O"
    Case 3:    Escribe "ARIAL", 9, 1, 0, 45, 100, "T R I P L I C A D O"
    Case 4:    Escribe "ARIAL", 9, 1, 0, 45, 100, "CUADRUPLICADO"
End Select

If HOMO Then
        Escribe "Arial Black", 10, 1, 0, 31, 90, "HOMOLOGACIÓN / Prueba"
End If

Escribe "Lucida Console", 8, 0, 0, 31, 140, "             CUIT: 33-70346790-9"
Escribe "Lucida Console", 8, 0, 0, 34, 140, "    IB Conv.Mult.: 901 33-70346790-9"
Escribe "Lucida Console", 8, 0, 0, 37, 140, "Partida Municipal: 1985"
Escribe "Lucida Console", 8, 0, 0, 40, 140, "     Imp.Internos: No Responsable"
Escribe "Lucida Console", 8, 0, 0, 43, 140, "  Fec.Inicio Act.: 01-09-1999"

'Primer Rectángulo
        DibujaRect 5, 101, 19, 113, 30, 30
        DibujaRect 50, 7, 78, 205, 30, 30
        DibujaLine 80, 7, 80, 205
        DibujaLine 88, 7, 88, 205
        DibujaLine 250, 7, 250, 205

        Escribe "ARIAL BLACK", 30, 1, 0, 5, 103, LaLetra
        Escribe "ARIAL", 7, 0, 0, 20, 102, "Código " & Format(CodFactAFIP, "00")
        Escribe "ARIAL BLACK", 20, 1, 0, 5, 130, TIPO
        Escribe "", 14, 1, 0, 17, 140, PtoVtaCpbte & "-" & Format(NROFAC, "00000000")

Escribe "", 10, 0, 0, 24, 146, Format(Date, "ddd dd MMM YYYY")

'............................................................................

'Si es N/Cred,  escribe a que factura esta aplicada
If TipoFac = 2 Then
    Escribe "", 10, 1, 1, 63, 160, "Aplicada a:"
    laLOGIK1 = ApPtoVta: laLOGIK2 = ApCpbte
    Escribe "", 10, 0, 1, 68, 145, "Factura Nº " & String(4 - Len(laLOGIK1), "0") & _
             laLOGIK1 & "-" & String(8 - Len(laLOGIK2), "0") & laLOGIK2
    Escribe "", 10, 0, 1, 73, 150, " del  " & Format(ApFecha, "ddd. dd mmm yyyy")
End If
'-------------------------------------------------------------------------------
'Escribe "", 12, 0, 0, 51, 153, Format(Time, "hh:mm")

'---------------------------------------------------------------------------
Escribe "", 10, 0, 0, 53, 10, ClteNomb & "         (" & CodCLTEf & ")     "
Escribe "", 10, 0, 0, 58, 10, ClteDir
'Cod.Post., Loc.,Pcia.
I4 = Asc(CltePCIA)
Escribe "", 10, 0, 0, 63, 10, "(" & Mid(ClteCPos, 1, 4) & ")  " & ClteLoc & Space(2) & _
                                Mid(PCIA(I4 - 65), 2, 10)

'Cond.IVA, CUIT
Escribe "", 10, 0, 0, 68, 10, "IVA " & CabFact.Combo1.List(ClteCIVA - 1) & _
                              "     CUIT : " & ClteCUIT
CUITdelCliente = ClteCUIT

'Cond.Vta,Fma.Envio,Remito
Escribe "", 10, 0, 0, 73, 10, "Condición de Venta : " & _
                                Mid(CabFact.Label104, 27, 40)
'----------------------------------------------------------------------------

'  Ttulo
    Escribe "", 10, 1, 1, 81, 12, "Cant."
    Escribe "", 10, 1, 1, 81, 56, "D e t a l l e"
    Escribe "", 10, 1, 1, 81, 83, "Despacho"
    Escribe "", 10, 1, 1, 81, 109, "% Descuento"
    Escribe "", 10, 1, 1, 81, 145, "Precio Unit."
    Escribe "", 10, 1, 1, 81, 178, "Importe "

    IMPRE.VP1.TableBorder = tbNone

s$ = "^+22mm|>+1mm|<+88mm|>+27mm|>+23mm|>+33mm"
h$ = "Cant.|       DETALLE      | % Descuento  | Precio  | Importe"

With IMPRE.VP1
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Lucida Console"
    .CurrentY = "92mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "4mm"
 End With

'IMPRE.VP1.EndDoc
 
'IMPRE.Show 1

TotDtos = 0: TotBruto = 0: TotSIVA = 0

If Motivo = 1 Then GoTo ConArticulos
'***********************************************************************************
'............................................................
' Para N/Credito o Debito sin detalle - Motivo mayor a 1
'............................................................

If DetNC.Text1(0).Text <> "" Then
    d$ = "" & "||" & DetNC.Text1(0).Text & "|" & "" & "|" _
          & "" & "|" & Format(DetNC.Text2(0).Text, " ###,##0.00")
    IMPRE.VP1.FontSize = 9
    IMPRE.VP1.AddTable s$, h$, d$, , , True
    IMPRE.VP1.FontSize = 3
    IMPRE.VP1.AddTable s$, h$, "|||||", , , True
        If Val(DetNC.Text2(0).Text) > 0 Then
            If DetNC.Check1(0).Value = 1 Then
               TotBruto = TotBruto + DetNC.Text2(0).Text
            Else
               TotSIVA = TotSIVA + DetNC.Text2(0).Text
            End If
        End If
End If

If DetNC.Text1(1).Text <> "" Then
    d$ = "" & "||" & DetNC.Text1(1).Text & "|" & "" & "|" _
          & "" & "|" & Format(DetNC.Text2(1).Text, " ###,##0.00")
    IMPRE.VP1.FontSize = 9
    IMPRE.VP1.AddTable s$, h$, d$, , , True
    IMPRE.VP1.FontSize = 3
    IMPRE.VP1.AddTable s$, h$, "|||||", , , True
        If Val(DetNC.Text2(1).Text) > 0 Then
            If DetNC.Check1(1).Value = 1 Then
               TotBruto = TotBruto + DetNC.Text2(1).Text
            Else
               TotSIVA = TotSIVA + DetNC.Text2(1).Text
            End If
        End If
End If

If DetNC.Text1(2).Text <> "" Then
    d$ = "" & "||" & DetNC.Text1(1).Text & "|" & "" & "|" _
          & "" & "|" & Format(DetNC.Text2(1).Text, " ###,##0.00")
    IMPRE.VP1.FontSize = 9
    IMPRE.VP1.AddTable s$, h$, d$, , , True
    IMPRE.VP1.FontSize = 3
    IMPRE.VP1.AddTable s$, h$, "|||||", , , True
        If Val(DetNC.Text2(2).Text) > 0 Then
            If DetNC.Check1(2).Value = 1 Then
               TotBruto = TotBruto + DetNC.Text2(2).Text
            Else
               TotSIVA = TotSIVA + DetNC.Text2(2).Text
            End If
        End If
End If

TotDtos = 0



GoTo SaleTotal
'***********************************************************************************
'............................................................
' Detalle con Artículos
'............................................................

ConArticulos:

For I1 = 1 To DetFact.FG1.Rows

If DetFact.FG1.TextMatrix(I1, 8) = "" Then GoTo SaleTotal

Secc = DetFact.FG1.TextMatrix(I1, 0)
'Pulg = CCur(DetFact.FG1.TextMatrix(I1, 1))
'Mtrs = CCur(DetFact.FG1.TextMatrix(I1, 2))
Cant = Val(DetFact.FG1.TextMatrix(I1, 6))
PVta = CCur(DetFact.FG1.TextMatrix(I1, 7))
Impte = DetFact.FG1.TextMatrix(I1, 8)

'-----------------------------------------------------------------------
' Si la Sección tiene marca usa la descr. de la Sección, sino usa la del Artículo
If DetFact.FG1.TextMatrix(I1, 11) = True Then
        'Arma la Descripción
        'Descri = DetFact.FG1.TextMatrix(i1, 5) & " " & MatrItems(i1, 1, 2) & " " & MatrItems(i1, 1, 1) & _
        '         " " & MatrItems(i1, 2, 2) & " " & MatrItems(i1, 2, 1) & _
        '         " " & MatrItems(i1, 3, 2) & " " & MatrItems(i1, 3, 1)
        
        Descri = DetFact.FG1.TextMatrix(I1, 5)
        If MatrItems(I1, 1, 2) <> "" And Val(MatrItems(I1, 1, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 1, 2) & " " & MatrItems(I1, 1, 1)
        End If
                 
        If MatrItems(I1, 2, 2) <> "" And Val(MatrItems(I1, 2, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 2, 2) & " " & MatrItems(I1, 2, 1)
        End If
                 
        If MatrItems(I1, 3, 2) <> "" And Val(MatrItems(I1, 3, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 3, 2) & " " & MatrItems(I1, 3, 1)
        End If
  '----------------
                 
        If MatrItems(I1, 4, 2) <> "" And Val(MatrItems(I1, 4, 2)) > 0 Then
            Descri = Descri & " - " & MatrItems(I1, 4, 2) & " " & MatrItems(I1, 4, 1)
        End If
                 
        If MatrItems(I1, 5, 2) <> "" And Val(MatrItems(I1, 5, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 5, 2) & " " & MatrItems(I1, 5, 1)
        End If
                 
        If MatrItems(I1, 6, 2) <> "" And Val(MatrItems(I1, 6, 2)) > 0 Then
            Descri = Descri & " " & MatrItems(I1, 6, 2) & " " & MatrItems(I1, 6, 1)
        End If
        
Else
        Descri = DetFact.FG1.TextMatrix(I1, 5)
End If

'-----------------------------------------------------------------------
'  Arma la cantidad
'laLOGIK1 = MatrItems(I1, 4, 2) & " " & MatrItems(I1, 4, 1) & "x " & MatrItems(I1, 5, 2) & " " & MatrItems(I1, 5, 1)
laLOGIK1 = MatrItems(I1, 7, 2)
'If MatrItems(i1, 5, 2) <> "" Then laLOGIK1 = laLOGIK1 & "x" & MatrItems(i1, 5, 2)
'If MatrItems(i1, 6, 2) <> "" Then laLOGIK1 = laLOGIK1 & "x" & MatrItems(i1, 6, 2)
'-----------------------------------------------------------------------------------------

' Obtiene los Descuentos del Cliente para esa seccion/rubro
DxClte CodCLTEf, Secc
Simbo = "": Desctos = ""
SumDtos = 0
For i2 = 1 To 5
    If Dtos(i2) > 0 Then
        Desctos = Desctos & Simbo & Dtos(i2)
        SumDtos = SumDtos + ((Impte - SumDtos) * (Dtos(i2) / 100))
    End If
    Simbo = "+"
Next i2

If DetFact.FG1.TextMatrix(I1, 13) <> "" Then
    Dim j9 As Integer
    j9 = IIf(Len(Descri) > 25, 25, Len(Descri))
    Descri = Mid(Descri, 1, 25) & Space(25 - j9)
    Descri = Descri & DetFact.FG1.TextMatrix(I1, 13)
End If

'd$ = Format(Cant, "##,##0") & "| |" & Descri & "|" & Desctos & "|" _
'          & Format(PVta, " ###,##0.00") & "|" & Format(Impte, " ###,##0.00")
d$ = laLOGIK1 & "||" & Descri & "|" & Desctos & "|" _
          & Format(PVta, " ###,##0.00") & "|" & Format(Impte, " ###,##0.00")

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 3
IMPRE.VP1.AddTable s$, h$, "|||||", , , True

TotDtos = TotDtos + SumDtos
TotBruto = TotBruto + Impte

Next I1

'...............................................................................
'       Fin Detalle con Artículos
'...............................................................................
'***********************************************************************************

SaleTotal:
        
If EnDolares Then
        TipoMoneda = "Son dólares"
        Escribe "Lucida Console", 7, 0, 0, 195, 10, LeyenDol1
        Escribe "Lucida Console", 7, 0, 0, 199, 10, LeyenDol2
        Escribe "Lucida Console", 7, 0, 0, 203, 10, LeyenDol3
        Escribe "Lucida Console", 7, 0, 0, 207, 10, LeyenDol4
        Escribe "Lucida Console", 7, 0, 0, 211, 10, LeyenDol5 & Format(LaCotiz, "###,##0.00")
        Escribe "Arial", 8, 1, 0, 250.5, 22, "U$S"
        Escribe "Arial", 8, 1, 0, 250.5, 72, "U$S"
        Escribe "Arial", 8, 1, 0, 250.5, 188, "U$S"
Else
        TipoMoneda = "Son pesos"
End If

'      Leyenda al Pie
i3 = InStr(1, Text1.Text, vbCrLf)
If i3 > 0 Then
        Escribe "Lucida Console", 12, 1, 1, 221, 10, Mid(Text1.Text, 1, i3 - 1)
Else
        Escribe "Lucida Console", 12, 1, 1, 221, 10, Mid(Text1.Text, 1, 30)
        i3 = 30
End If

I4 = InStr(i3 + 2, Text1.Text, vbCrLf)
If I4 > 0 Then
        Escribe "Lucida Console", 12, 1, 1, 225, 10, Mid(Text1.Text, i3 + 2, (I4 - 1) - i3)
Else
        Escribe "Lucida Console", 12, 1, 1, 225, 10, Mid(Text1.Text, i3 + 2, 30)
        I4 = i3 + 32
End If

i3 = InStr(I4 + 2, Text1.Text, vbCrLf)
If i3 > 0 Then
        Escribe "Lucida Console", 12, 1, 1, 229, 10, Mid(Text1.Text, I4 + 2, (i3 - 1) - I4)
Else
        Escribe "Lucida Console", 12, 1, 1, 229, 10, Mid(Text1.Text, I4 + 2, 30)
End If

Impte = TotBruto + TotSIVA

    
    DibujaRect 260, 7, 268, 37, 30, 30
    DibujaRect 260, 39, 268, 68, 30, 30
    DibujaRect 260, 70, 268, 100, 30, 30
    DibujaRect 260, 102, 268, 132, 30, 30
 '   DibujaRect 260,  134, 268, 164, 30, 30

    IMPRE.VP1.BrushStyle = bsSolid
    IMPRE.VP1.BrushColor = &HE0E0E0
    DibujaRect 260, 166, 268, 202, 30, 30
    IMPRE.VP1.BrushStyle = 1
    
    Escribe "Arial", 10, 0, 0, 255, 11, "Subtotal"
    Escribe "Arial", 10, 0, 0, 255, 43, "Descuento"
    Escribe "Arial", 10, 1, 0, 255, 75, "Subtotal"
    Escribe "Arial", 10, 0, 0, 255, 106, "IVA Insc. 21%"
 '   Escribe "Arial", 10, 0, 0, 255, 138, "IVA No Insc."
    Escribe "Arial", 10, 1, 0, 255, 179, "TOTAL"
    
    Escribe "", 9, 1, 0, 231, 184, Format(Impte, "###,###,##0.00")
    
 IMPRE.VP1.StartTable

'-------------------------------------------------------------------------------
'................................................................................
'  imprime Totales

ValIB = 0

TotBruto = TotBruto - TotDtos

If Motivo <> 3 Then
    If ClteCIVA < 3 Then
        ValIB = Round(TotBruto * (CCur(Text2.Text) / 100), 2)
        TotIVAIns = TotBruto * IVAIns
        If ClteCIVA = 2 Then
           TotIVANI = TotBruto * IVANI
        End If
    End If
Else
'  Es nota de crédito o deb. por IB
    If ClteCIVA < 3 Then
        ValIB = TotSIVA
        TotIVAIns = 0: TotSIVA = 0: Impte = 0
    End If
End If

If ConPercep <> 1 Then ValIB = 0                 ' sin no tiene marca en parametro pone total en 0

d$ = Format(Impte, "###,##0.00") & "|"

TOTNETO = TotBruto + TotIVAIns + TotIVANI + TotSIVA + ValIB

With IMPRE.VP1
    .FontBold = False
    .FontItalic = False
    .Font.Name = "Arial"
    .CurrentY = "262mm"
    .CurrentX = 0
    .TextAlign = 0
    .IndentLeft = 0
    .MarginLeft = "3mm"
 End With

s$ = ">+33mm|>+30mm|>+33mm|>+33mm|>+32mm|>+38mm"
h$ = "subt1|dtos|subt2|ivains|ivani|total"

d$ = d$ & Format(TotDtos, "###,###,##0.00") & "|" & Format(TotBruto, "###,###,##0.00") & _
    "|" & Format(TotIVAIns, "###,##0.00") & "|" & Format(ValIB, "###,###.##") & _
    "|" & Format(TOTNETO, "###,###,##0.00")

IMPRE.VP1.FontSize = 9
IMPRE.VP1.AddTable s$, "", d$, , , True
IMPRE.VP1.TableCell(tcFontBold, 1, 3) = True
IMPRE.VP1.TableCell(tcFontBold, 1, 6) = True
 
 IMPRE.VP1.EndTable
 
    If ValIB > 0 Then
            DibujaRect 260, 134, 268, 164, 30, 30
            Escribe "Arial", 10, 0, 0, 255, 135, "Perc. II.BB."
    End If

'        Escribe "", 7, 1, 1, 272, 163, "(" & NROFAC & ")"
   
' Importe en letras
EnLetras = MontoEscrito(TOTNETO)
Escribe "Lucida Console", 7, 0, 0, 236, 30, TipoMoneda & Mid$(EnLetras, 1, 70)
Escribe "Lucida Console", 7, 0, 0, 240, 30, Mid$(EnLetras, 71, 70)
Escribe "Lucida Console", 7, 0, 0, 244, 30, Mid$(EnLetras, 141, 70)
    
    
 ' Código de Barras
 Escribe "Lucida Console", 12, 0, 0, 270, 132, "   C.A.E.: " & NroCAE
 Escribe "Lucida Console", 10, 0, 0, 274, 137, " Fec.Vto.: " & FecVtoCAE
   
  CalcDV
    
'BarCod.CopyImage
'IMPRE.VP1.DrawPicture BarCod.Picture, "10mm", "270mm"
    
    Obtener_QR
     IMPRE.VP1.DrawPicture Picture3, "7mm", "269mm", "25mm", "25mm"      '   QR

If MotRech <> "NULL" And MotRech <> "00" Then
    Escribe "Lucida Console", 10, 0, 0, 284, 40, "" & _
            "Se asignó CAE pero con advertencias. Motivo: " & MotRech
End If

If Copias = 1 Then
    IMPRE.VP1.NewPage
    GoTo Comienza1
End If

 IMPRE.VP1.EndDoc
'IMPRE.Show 1

'Genera PDF
VSPDF81.Title = TIPO & "-" & LaLetra & "-" & PtoVtaCpbte & "-" & NROFAC
VSPDF81.ConvertDocument IMPRE.VP1, "c:\Factu\Comprob\" & _
                         Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "\" & _
                         VSPDF81.Title & ".pdf"
 
'Muestra PDF
PDF.PDF1.LoadFile "c:\Factu\Comprob\" & _
                         Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "\" & _
                         VSPDF81.Title & ".pdf"
PDF.Caption = VSPDF81.Title
PDF.Show
PDF.PDF1.setShowToolbar True

Sleep 1000
QRunaVez = False
'Kill "QR.png"

End Sub

Sub Graba()
Dim I0%, I1%, i2%, CanItems As Integer
Dim TotPcos As Currency, TotPvta As Currency, DiasVto As Integer

' vbMsgBoxTitle = "Grabación de la Factura"
'  vbMsgBoxText = "El siguiente proceso registrará la Factura  " & vbCrLf & "Desea continuar ?  "
'  vbMsgBoxResp = vbYesNoCancel + vbExclamation + vbApplicationModal + vbDefaultButton1

' vbResponse = MsgBox(vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle)
            
'vbResponse = MiMsgBox(0, vbMsgBoxText, vbMsgBoxTitle, vbMsgBoxResp)

'  Select Case vbResponse
'         Case vbYes
'              GoTo Grabacion

'         Case vbNo
'             NoAnduvo
'              GoTo Salir2

'         Case vbCancel
'              GoTo Salir2
'  End Select
 
Grabacion:

dbCCTE.ConnectionString = BDatos1
dbCCTE.Open

dbFACT.ConnectionString = BDatos1
dbFACT.Open

dbSTOK.ConnectionString = BDatos1
dbSTOK.Open

dbCCTE.BeginTrans
dbFACT.BeginTrans
dbSTOK.BeginTrans

TotBruto = TotBruto + TotSIVA

'  si es NC por Motivo distinto a dev. de mercadería
'        o no es Factura ni NC
'    Saltea la grabación de detalle
If (TipoFac = 2 And Motivo > 1) Or TipoFac > 2 Then
  GoTo AlTotal
End If


'---------------------------------------------------------
'         Graba Detalle - Stock,Movstck y Estadística
'---------------------------------------------------------

CanItems = 0

For I1 = 1 To DetFact.FG1.Rows

If DetFact.FG1.TextMatrix(I1, 8) = "" Then GoTo AlTotal

CanItems = CanItems + 1

Correa = DetFact.FG1.TextMatrix(I1, 0)
RESTO = DetFact.FG1.TextMatrix(I1, 14)
Impte = DetFact.FG1.TextMatrix(I1, 8)
' Obtiene los Descuentos del Cliente para esa seccion/rubro

DxClte CodCLTEf, Correa
SumDtos = 0
For i2 = 1 To 5
    If Dtos(i2) > 0 Then
        SumDtos = SumDtos + ((Impte - SumDtos) * (Dtos(i2) / 100))
    End If
Next i2

Graba_1:
    
'------------------------------------
'    Stock - Stock
'------------------------------------

miSQL1 = "SELECT * FROM Stock WHERE COD1 = '" & Correa & _
        "' and COD2 = '" & RESTO & "'"
        
RgSTOK.Open miSQL1, dbSTOK, adOpenDynamic, adLockPessimistic
    
    If RgSTOK.EOF Then
        RgSTOK.AddNew
        RgSTOK!COD1 = DetFact.FG1.TextMatrix(I1, 0)
        RgSTOK!Cod2 = DetFact.FG1.TextMatrix(I1, 1)
        RgSTOK!Stmin = 0
        RgSTOK!StmAX = 0
        RgSTOK!StREP = 0
        RgSTOK!DEP1 = 0
        RgSTOK!EST1 = 0
        RgSTOK!ESTT1 = 0
        RgSTOK!DEP2 = 0
        RgSTOK!EST2 = 0
        RgSTOK!ESTT2 = 0
        RgSTOK!STUnid = 0
        RgSTOK!Pulg = 0
        RgSTOK!Mtr = 0
        RgSTOK!Stant = 0
        RgSTOK!ENTMES = 0
        RgSTOK!SALMES = 0
        RgSTOK!AJEMES = 0
        RgSTOK!AJSMES = 0
        RgSTOK!STUnid = 0
    End If
    
If Val(DetFact.FG1.TextMatrix(I1, 6)) > 0 Then
    If TipoFac = 1 Then
        RgSTOK!STUnid = RgSTOK!STUnid - DetFact.FG1.TextMatrix(I1, 6)
        RgSTOK!SALMES = RgSTOK!SALMES + DetFact.FG1.TextMatrix(I1, 6)
    Else
        RgSTOK!STUnid = RgSTOK!STUnid + DetFact.FG1.TextMatrix(I1, 6)
        RgSTOK!SALMES = RgSTOK!SALMES - DetFact.FG1.TextMatrix(I1, 6)
    End If
End If
    
If Val(DetFact.FG1.TextMatrix(I1, 2)) > 0 Then
    If TipoFac = 1 Then
        RgSTOK!STUnid = RgSTOK!Mtr - DetFact.FG1.TextMatrix(I1, 2)
        RgSTOK!SALMES = RgSTOK!SALMES + DetFact.FG1.TextMatrix(I1, 2)
    Else
        RgSTOK!STUnid = RgSTOK!Mtr + DetFact.FG1.TextMatrix(I1, 2)
        RgSTOK!SALMES = RgSTOK!SALMES - DetFact.FG1.TextMatrix(I1, 2)
    End If
End If

    RgSTOK!FACTUAL = Date
    RgSTOK!USUARIO = Mid(Red_Usuario, 1, 6)
    
    RgSTOK.Update
    RgSTOK.Close
    
'------------------------------------
'   Movimientos de Stock - MovStock
'------------------------------------

'
' Cód.Movim.: 11-N/Créd.'A'       12-N/Créd.'B'          13-Compra
'             14-Ent.Canje        15-Aj.en Más           16-Cpra.'R'
'             17-N/Créd.'R'       18-Aj.Inventario       19-
'             21-Factura 'A'      22-Fact.'B'            23-Remito
'             24-Sal.x Cje.       25-Aj.en Menos         26-
'             27-Fact.'R'         28-Aj.en menos Invent. 29-
'
If TipoFac = 1 Then
    Select Case LaLetra
        Case "A"
            CodMovs = "21"
        Case "B"
            CodMovs = "22"
        Case "C"
            CodMovs = "27"
        Case "R"
            CodMovs = "23"
    End Select
End If

If TipoFac = 2 Then
    Select Case LaLetra
        Case "A"
            CodMovs = "11"
        Case "B"
            CodMovs = "12"
        Case "C"
            CodMovs = "17"
    End Select
End If

If TipoFac = 4 Then
            CodMovs = "27"
End If

miSQL1 = "SELECT * FROM MovStock WHERE COD1= '" & Correa & "' and COD2 = '" & _
        RESTO & "' and CPBTE = " & _
        NROFAC & " AND PTOVTA = " & PtoVtaCpbte & " AND TIPO = '" & CodMovs & "'"
RgSTOK.Open miSQL1, dbSTOK, adOpenDynamic, adLockPessimistic
    
'    If Not RgSTOK.EOF Then
'        NoAnduvo = True
'        RgSTOK.Close
'        GoTo Confirmar
 '       GoTo SaltaAdd1
'    End If

    RgSTOK.AddNew
    RgSTOK!COD1 = DetFact.FG1.TextMatrix(I1, 0)
    RgSTOK!Cod2 = DetFact.FG1.TextMatrix(I1, 1)
    RgSTOK!TIPO = CodMovs
    RgSTOK!PtoVta = PtoVtaCpbte
    RgSTOK!Cpbte = NROFAC
    
SaltaAdd1:
    RgSTOK!item = CanItems
    RgSTOK!FECHA = Now
    RgSTOK!Cant = Val(DetFact.FG1.TextMatrix(I1, 6))
        
        RgSTOK!Pulg = 0
        RgSTOK!Mtr = 0
    
    If DetFact.FG1.TextMatrix(I1, 1) <> "" Then
        RgSTOK!Pulg = DetFact.FG1.TextMatrix(I1, 1)
    End If
    
    If DetFact.FG1.TextMatrix(I1, 2) <> "" Then
        RgSTOK!Mtr = DetFact.FG1.TextMatrix(I1, 2)
    End If
    
    RgSTOK!Milim = Val(DetFact.FG1.TextMatrix(I1, 3))
    RgSTOK!Telas = Val(DetFact.FG1.TextMatrix(I1, 4))

    RgSTOK!USUARIO = Mid(Red_Usuario, 1, 6)
    
    RgSTOK.Update
    RgSTOK.Close
    
'------------------------------------
'    Despachos de Aduana si corresponde
'------------------------------------

miSQL1 = "SELECT * FROM Despachos WHERE COD1 = '" & Correa & _
        "' and COD2 = '" & RESTO & "' and NRODESP = '" & DetFact.FG1.TextMatrix(I1, 13) & "'"
        
RgSTOK.Open miSQL1, dbSTOK, adOpenDynamic, adLockPessimistic
    
    If RgSTOK.EOF Then
        GoTo SinGrabar1
    End If
    
If Val(DetFact.FG1.TextMatrix(I1, 6)) > 0 Then
    If TipoFac = 1 Then
        RgSTOK!Stock = RgSTOK!Stock - DetFact.FG1.TextMatrix(I1, 6)
        RgSTOK!salidaS = RgSTOK!salidaS + DetFact.FG1.TextMatrix(I1, 6)
    Else
        RgSTOK!Stock = RgSTOK!Stock + DetFact.FG1.TextMatrix(I1, 6)
        RgSTOK!salidaS = RgSTOK!salidaS - DetFact.FG1.TextMatrix(I1, 6)
    End If
End If
    
If Val(DetFact.FG1.TextMatrix(I1, 2)) > 0 Then
    If TipoFac = 1 Then
        RgSTOK!Stock = RgSTOK!Stock - DetFact.FG1.TextMatrix(I1, 2)
        RgSTOK!salidaS = RgSTOK!salidaS + DetFact.FG1.TextMatrix(I1, 2)
    Else
        RgSTOK!Stock = RgSTOK!Stock + DetFact.FG1.TextMatrix(I1, 2)
        RgSTOK!salidaS = RgSTOK!salidaS - DetFact.FG1.TextMatrix(I1, 2)
    '    RgSTOK!ENTRADA = RgSTOK!ENTRADA + DetFact.FG1.TextMatrix(I1, 2)
    End If
End If

    RgSTOK!fecusal = Date
    RgSTOK.Update
    
SinGrabar1:
    
    RgSTOK.Close
    
'------------------------------------
'   Estadística por Artículo - FCESTAD1
'------------------------------------

miSQL1 = "SELECT * FROM FCEstad1 WHERE COD1 = '" & Correa & "' and COD2 = '" & _
        RESTO & "' and CPBTE = " & _
        NROFAC & " AND PTOVTA = " & PtoVtaCpbte & " AND TIPO = " & TipoFac & _
        " AND LETRA = '" & LaLetra & "'"
RgFACT.Open miSQL1, dbFACT, adOpenDynamic, adLockPessimistic
    
'    If Not RgFACT.EOF Then
'        NoAnduvo = True
'        RgFACT.Close
'        GoTo Confirmar
'       GoTo SaltaAdd2
'   End If

    RgFACT.AddNew
    RgFACT!COD1 = DetFact.FG1.TextMatrix(I1, 0)
    RgFACT!Cod2 = DetFact.FG1.TextMatrix(I1, 1)
    RgFACT!TIPO = TipoFac
    RgFACT!Letra = LaLetra
    RgFACT!PtoVta = PtoVtaCpbte
    RgFACT!Cpbte = NROFAC
    RgFACT!item = CanItems
    RgFACT!CLTE = CodCLTEf
    
SaltaAdd2:
    
RgFACT!FECHA = Now
RgFACT!Cant = Val(DetFact.FG1.TextMatrix(I1, 6))

If DetFact.FG1.TextMatrix(I1, 1) = "" Then
    RgFACT!Pulg = 0
Else
    RgFACT!Pulg = DetFact.FG1.TextMatrix(I1, 1)
End If

If DetFact.FG1.TextMatrix(I1, 2) = "" Then
    RgFACT!Mtr = 0
Else
    RgFACT!Mtr = DetFact.FG1.TextMatrix(I1, 2)
End If

RgFACT!Milim = Val(DetFact.FG1.TextMatrix(I1, 3))
RgFACT!Telas = Val(DetFact.FG1.TextMatrix(I1, 4))
RgFACT!PCos = 0
RgFACT!PVta = DetFact.FG1.TextMatrix(I1, 7)
RgFACT!PEsp = DetFact.FG1.TextMatrix(I1, 7)
RgFACT!bon = SumDtos
RgFACT!IVA = 0
RgFACT!Impte = DetFact.FG1.TextMatrix(I1, 8)
RgFACT!tipo9 = 0
RgFACT!gptipo = ""
RgFACT!gpdsd = ""
RgFACT!gphst = ""

'RgFACT!USUARIO = Mid(Red_Usuario, 1, 6)
RgFACT!USUARIO = "   "
    
    RgFACT.Update
    RgFACT.Close

TotCan = TotCan + Val(DetFact.FG1.TextMatrix(I1, 6))

Next I1

'
'*****************************************************
'

AlTotal:
'----------------------------------------------
'   Totales de Facturación
'
laLOGIK1 = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
miSQL2 = "SELECT * FROM Totales WHERE FECHA = " & Separa & laLOGIK1 & Separa
RgTOTL.Open miSQL2, dbFACT, adOpenDynamic, adLockPessimistic

If RgTOTL.EOF Then

    RgTOTL.AddNew
    RgTOTL!FECHA = Date
' Pone en cero los campos en caso de que sea un reg. nuevo
    For I0 = 1 To RgTOTL.Fields.Count - 2
    
'       If RGfciva.Fields.Item(i0).Name = "FACT1" Then Exit For
       RgTOTL.Fields.item(I0) = 0
    
    Next I0

End If
  
Select Case TipoFac

    Case 1
        RgTOTL!PCos = RgTOTL!PCos + TotPcos
        RgTOTL!PVta = RgTOTL!PVta + TotBruto
        RgTOTL!PEsp = RgTOTL!PEsp + TotBruto
        
        If LaLetra = "A" Then
          RgTOTL!faca = RgTOTL!faca + 1
          RgTOTL!UNIDA = RgTOTL!UNIDA + TotCan
          RgTOTL!PESPA = RgTOTL!PESPA + TotBruto
        End If
        
        If LaLetra = "B" Then
          RgTOTL!facb = RgTOTL!facb + 1
          RgTOTL!UNIDB = RgTOTL!UNIDB + TotCan
          RgTOTL!PESPb = RgTOTL!PESPb + TotBruto
        End If
    
    Case 2
      RgTOTL!PCos = RgTOTL!PCos - TotPcos
      RgTOTL!PVta = RgTOTL!PVta - TotBruto
      RgTOTL!PEsp = RgTOTL!PEsp - TotBruto
        
        If LaLetra = "A" Then
          RgTOTL!nca = RgTOTL!nca + 1
          RgTOTL!UNIDA = RgTOTL!UNIDA - TotCan
          RgTOTL!PESPA = RgTOTL!PESPA - TotBruto
        End If
        
        If LaLetra = "B" Then
          RgTOTL!ncb = RgTOTL!ncb + 1
          RgTOTL!UNIDB = RgTOTL!UNIDB - TotCan
          RgTOTL!PESPb = RgTOTL!PESPb - TotBruto
        End If
    
    Case 3
      RgTOTL!PCos = RgTOTL!PCos + TotPcos
      RgTOTL!PVta = RgTOTL!PVta + TotBruto
      RgTOTL!PEsp = RgTOTL!PEsp + TotBruto
        
        If LaLetra = "A" Then
          RgTOTL!nda = RgTOTL!nda + 1
          RgTOTL!UNIDA = RgTOTL!UNIDA + TotCan
          RgTOTL!PESPA = RgTOTL!PESPA + TotBruto
        End If
        
        If LaLetra = "B" Then
          RgTOTL!ndb = RgTOTL!ndb + 1
          RgTOTL!UNIDB = RgTOTL!UNIDB + TotCan
          RgTOTL!PESPb = RgTOTL!PESPb + TotBruto
        End If

End Select


RgTOTL!FACTUAL = Date
RgTOTL!USUARIO = Mid(Red_Usuario, 1, 6)

    RgTOTL.Update: RgTOTL.Close
'------------------------------------
'         Subdiario de Ventas
'------------------------------------

miSQL1 = "SELECT * FROM FCIVAVTA WHERE CPBTE = " & NROFAC & " AND TIPO = '" & TipoFac & _
         "' AND LETRA = '" & LaLetra & "' AND PTOVTA = " & PtoVtaCpbte

RgFACT.Open miSQL1, dbFACT, adOpenDynamic, adLockPessimistic
    
    If Not RgFACT.EOF Then
        MsgBox "FCIVAVTA" & vbCrLf & miSQL1, vbCritical + vbOKOnly, "Error"
        NoAnduvo = True
        RgFACT.Close
        GoTo Confirmar
    End If

    RgFACT.AddNew
    
RgFACT!FECHA = Now
RgFACT!PtoVta = PtoVtaCpbte
RgFACT!Cpbte = NROFAC
RgFACT!Letra = LaLetra
RgFACT!TIPO = TipoFac

RgFACT!CLTE = CodCLTEf
RgFACT!Nomb = ClteNomb
RgFACT!PCIA = CltePCIA
RgFACT!CVTA = ClteCVta

If TipoFac = 1 Or TipoFac = 4 Then
    RgFACT!moti = TipoVta
Else
    RgFACT!moti = Motivo
End If

RgFACT!CIVA = ClteCIVA
RgFACT!Cuit = Mid(ClteCUIT, 1, 14)
RgFACT!VEND = ClteVend
RgFACT!ZONA = 0
RgFACT!TotCan = TotCan
RgFACT!GRINS = TotBruto
RgFACT!GRNOINs = 0
RgFACT!IVAIns = TotIVAIns
RgFACT!IVANOINS = TotIVANI
RgFACT!PORCIB = CCur(Text2.Text)
RgFACT!TotIB = ValIB
RgFACT!EXENTO = 0
RgFACT!totcos = 0
RgFACT!ITEMS = CanItems
RgFACT!bon = TotDtos
RgFACT!NOIMPR = 0
RgFACT!COMIS = 0

' RgFACT!USUARIO = Mid(Red_Usuario, 1, 6)

    RgFACT.Update
    RgFACT.Close

aCltes:
'------------------------------------
'   Actualiza el Mtro. de Clientes
'------------------------------------
Corr = 0
miSQL1 = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTEf
RgCLTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If RgCLTE.EOF Then
    MsgBox "Clientes 2", vbCritical + vbOKOnly, "Error"
    NoAnduvo = True
    RgCLTE.Close
Else
    If TipoFac = 2 Then
      RgCLTE!DEUDA = RgCLTE!DEUDA - TOTNETO
      RgCLTE!corr1 = RgCLTE!corr1 + 1
      Corr = RgCLTE!corr1
      ImputaFact
    Else
      RgCLTE!DEUDA = RgCLTE!DEUDA + TOTNETO
    End If
    RgCLTE!USUARIO = Mid(Red_Usuario, 1, 6)
    RgCLTE!FACTUAL = Date
    
    RgCLTE.Update
    RgCLTE.Close

End If

ACtaCte:
'------------------------------------
'         Ctas. Ctes.
'------------------------------------

If CodCLTEf = 0 Then GoTo Confirmar

miSQL1 = "SELECT * FROM CtasCtes WHERE CLTE = " & CodCLTEf & " AND CPBTE = " & _
         NROFAC & " AND TIPO = " & TipoFac & _
         " AND LETRA = '" & LaLetra & "' AND PREFIJO = " & PtoVtaCpbte

RgCCTE.Open miSQL1, dbCCTE, adOpenDynamic, adLockPessimistic
    
    If Not RgCCTE.EOF Then
        MsgBox "CtasCtes", vbCritical + vbOKOnly, "Error"
        NoAnduvo = True
        RgCCTE.Close
        GoTo Confirmar
    End If
    
    ' Tabla de Cond. de Venta
miSQL2 = "SELECT *  FROM Fctabla1 WHERE CTAB = 'CV   ' and COD = '" & _
        ClteCVta & "'"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

If RgTABL.EOF Then
        DiasVto = 30
Else
        DiasVto = RgTABL!NumSD3
End If

RgTABL.Close

    RgCCTE.AddNew

RgCCTE!CLTE = CodCLTEf
RgCCTE!FECHA = Date
RgCCTE!TIPO = TipoFac
RgCCTE!prefijo = PtoVtaCpbte
RgCCTE!Cpbte = NROFAC
RgCCTE!Letra = LaLetra
RgCCTE!imput1 = 0
RgCCTE!imput2 = 0
RgCCTE!imput3 = 0
RgCCTE!imput4 = 0
RgCCTE!imput5 = 0
RgCCTE!imput6 = 0
RgCCTE!Impte = TOTNETO
RgCCTE!DEBE = TOTNETO
RgCCTE!CVTA = ClteCVta
RgCCTE!bon = TotDtos
RgCCTE!tipo9 = 0
If LaLetra = "C" Then
    RgCCTE!tipo9 = 1
End If
RgCCTE!moti = Motivo
RgCCTE!FECVTO = Date + DiasVto
RgCCTE!USUAR = Mid(Red_Usuario, 1, 6)

If TipoFac = 2 Then
    RgCCTE!imput1 = Corr
End If

    RgCCTE.Update
    RgCCTE.Close

Confirmar:

'  Si salió todo bien, Confirma las Transacciones
If NoAnduvo Then
    dbCCTE.RollbackTrans
    dbFACT.RollbackTrans
    dbSTOK.RollbackTrans
    MsgBox "Hubo Errores durante la GRABACIÓN" & vbCrLf & "VERIFIQUE !!!", vbExclamation, "Aviso"
Else
    dbCCTE.CommitTrans
    dbFACT.CommitTrans
    dbSTOK.CommitTrans
End If

dbCCTE.Close
dbSTOK.Close
dbFACT.Close

Exit Sub

'---------------------------------------------
Salir1:

Unload PDF

Unload Me

Salir2:

End Sub

Private Sub ImputaFact()
    
'   Cuentas Corrientes
laLOGIK1 = Month(ApFecha) & "/" & Day(ApFecha) & "/" & Year(ApFecha)
miSQL1 = "SELECT * FROM CtasCtes WHERE CLTE= " & CodCLTEf & " AND CPBTE= " & ApCpbte & _
                                " AND FECHA= " & Separa & laLOGIK1 & Separa & " AND TIPO= " & ApTipo
RgCCTE.Open miSQL1, dbCCTE, adOpenDynamic, adLockPessimistic

If RgCCTE.EOF Then
    RgCCTE.Close
        MsgBox "CtasCtes 2", vbCritical + vbOKOnly, "Error"
    NoAnduvo = True
    Exit Sub
    End If

If RgCCTE!imput1 = "  " Or RgCCTE!imput1 = "00" Or RgCCTE!imput1 = "0 " Then
    RgCCTE!imput1 = Corr
    GoTo Salta1
    End If

If RgCCTE!imput2 = "  " Or RgCCTE!imput2 = "00" Or RgCCTE!imput2 = "0 " Then
    RgCCTE!imput2 = Corr
    GoTo Salta1
    End If

If RgCCTE!imput3 = "  " Or RgCCTE!imput3 = "00" Or RgCCTE!imput3 = "0 " Then
    RgCCTE!imput3 = Corr
    GoTo Salta1
    End If

If RgCCTE!imput4 = "  " Or RgCCTE!imput4 = "00" Or RgCCTE!imput4 = "0 " Then
    RgCCTE!imput4 = Corr
    GoTo Salta1
    End If

If RgCCTE!imput5 = "  " Or RgCCTE!imput5 = "00" Or RgCCTE!imput5 = "0 " Then
    RgCCTE!imput5 = Corr
    GoTo Salta1
    End If

If RgCCTE!imput6 = "  " Or RgCCTE!imput6 = "00" Or RgCCTE!imput6 = "0 " Then
    RgCCTE!imput6 = Corr
    End If
    
Salta1:
    
    RgCCTE!DEBE = RgCCTE!DEBE - TOTNETO
    
RgCCTE.Update
RgCCTE.Close

Salta2:

'      Graba la Imputación

laLOGIK1 = Month(ApFecha) & "/" & Day(ApFecha) & "/" & Year(ApFecha)
miSQL1 = "SELECT * FROM Imputacion WHERE CLTE = " & CodCLTEf & " AND CPBTE = " & ApCpbte & _
                                " AND FECHA = " & Separa & laLOGIK1 & Separa & " AND TIPO = '" & ApTipo & "'"
RgCCTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If Not RgCCTE.EOF Then
    RgCCTE.Close
        MsgBox "Imputacion", vbCritical + vbOKOnly, "Error"
    NoAnduvo = True
    Exit Sub
    End If

RgCCTE.AddNew

RgCCTE!Cpbte = NROFAC
RgCCTE!FECHA = Date
RgCCTE!TIPO = 2
RgCCTE!CLTE = CodCLTEf

RgCCTE!Corr = Corr
RgCCTE!fechai = ApFecha
RgCCTE!cpbtei = ApCpbte
RgCCTE!tipoi = ApTipo
RgCCTE!fecvtoi = ApFecha

RgCCTE!Impte = TOTNETO
RgCCTE!USUARIO = Mid(Red_Usuario, 1, 6)
    
    RgCCTE.Update
    RgCCTE.Close
    
End Sub

Private Sub Text2_GotFocus()

Text2.SelStart = 0
Text2.SelLength = Len(Text2.Text)

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)


If KeyAscii = 8 Or KeyAscii = vbKeyReturn Then Exit Sub

If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
If KeyAscii = Asc(",") Then Exit Sub

If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then KeyAscii = 0

End Sub

Private Sub Text2_Change()

'If Text2.Text > 0 Then
'    Command1.Enabled = True
'else
'    Command1.Enabled = False
'End If


End Sub

Private Sub ConectaAFIP()
Dim tra, cms, Path As String, Certificado As String, ClavePrivada As String
Dim ta, Ok, Qty, Lastid, Tipo_Cbte, LastCbte, Punto_Vta
Dim Pagina As String

AFIPbien = False

WB1.Navigate2 "about:blank"
DoEvents

If Check1.Value = 1 Then                    ' Si es fact.Electr. agranda formulario
    Me.Move 1600, 1500, 10900, 5200
End If
' Ejemplo de Uso de Interface COM con Web Services AFIP (PyAfipWs)
' 2008 (C) Mariano Reingart <mariano@nsis.com.ar>

    Dim WSAA As Object, WSFE As Object
    
    On Error GoTo ManejoError
WB1.Document.Write "Conexión con AFIP" & "<br><br>"
DoEvents
    
    ' Crear objeto interface Web Service Autenticación y Autorización
    Set WSAA = CreateObject("WSAA")
    
'-------------------------------------------------------
'    prueba nueva 7/12/2021
'
'    Dim pyCache As String, pyWSAA_url As String, pyWrapper As String, pyProxy As String
'    Dim pyCacert As String
'
'    pyCache = ""
'    pyWSAA_url = ""
'    pyWrapper = ""
'    pyProxy = ""
'    pyCacert = WSAA.InstallDir + "\conf\afip_ca_info.crt"
'
'    MsgBox pyCacert
'
'If HOMO Then
'    pyWSAA_url = "https://wsaahomo.afip.gov.ar/ws/service/LoginCms?WSDL"        ' homologación
'    MsgBox "Homo"
'Else
'    pyWSAA_url = "https://wsaa.afip.gov.ar/ws/services/LoginCms"     ' producción
'    MsgBox "Prod"
'End If
    
' WSAA.Conectar pyCache, pyWSAA_url, pyProxy, pyWrapper, pyCacert
'----------------------------------------------------------------------
    
'-------------------------------------------------------

WB1.Document.Write "Crea Objeto WSAA" & "<br><br>"
DoEvents
    ' Generar un Ticket de Requerimiento de Acceso (TRA)
    tra = WSAA.CreateTRA()
'    Debug.Print tra

WB1.Document.Write "Ticket Req.Acceso (tra) WSFE:" & tra & "<br><br>"
WB1.Document.parentWindow.scrollTo 0, 9999999
DoEvents
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    Path = CurDir() + "\"
'    Path = "D:\"
   
    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
If HOMO Then
    Certificado = "empresa.crt" ' certificado de Prueba
    ClavePrivada = "Homo.key"    ' clave privada
Else
    Certificado = "ALESTEL.crt" ' certificado de PRODUCCIÓN
    ClavePrivada = "empresa.key"    ' clave privada
End If

'MsgBox "1 " & tra & " " & Path + Certificado & " " & Path + ClavePrivada

'     Generar el mensaje firmado (CMS)
'     cms = WSAA.SignTRA(tra, Certificado, ClavePrivada)
    cms = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)

'MsgBox "1 " & tra & " " & Path + Certificado & " " & Path + ClavePrivada
        
WB1.Document.Write "Mensaje Confirmado(cms):" & cms & "<br><br>"
    DoEvents


'MsgBox "1 " & ta

    ' Llamar al web service para autenticar:
If HOMO Then
    ta = WSAA.CallWSAA(cms, "https://wsaahomo.afip.gov.ar/ws/services/LoginCms") ' Hologación
Else
    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms")     ' Producción
End If
'    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/logincms?wsdl") ' Producción


    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
    WB1.Document.Write ta & "<br><br>"
    WB1.Document.Write "Token: " & WSAA.token & "<br><br>"
    WB1.Document.Write "Sign: " & WSAA.Sign & "<br><br>"
WB1.Document.parentWindow.scrollTo 0, 9999999
DoEvents

MsgBox "1"

    ' Una vez obtenido, se puede usar el mismo token y sign por 6 horas
    ' (este período se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electrónica
    Set WSFE = CreateObject("WSFE")
    ' Setear tocken y sing de autorización (pasos previos)
    WSFE.token = WSAA.token
    WSFE.Sign = WSAA.Sign
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFE.Cuit = "33703467909"
    

    ' Conectar al Servicio Web de Facturación
If HOMO Then
    Ok = WSFE.Conectar("https://wswhomo.afip.gov.ar/wsfe/service.asmx")          ' homologación
Else
    Ok = WSFE.Conectar("https://servicios1.afip.gov.ar/wsfe/service.asmx")       ' producción
End If

 '    Ok = WSFE.Conectar("https://wsw.afip.gov.ar/wsfe/service.asmx")        ' producción
 '    Ok = WSFE.Conectar("https://servicios1.afip.gov.ar/wsfe/service.asmx") ' producción
    
    ' Llama a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFE.Dummy
        
    WB1.Document.Write "appserver status:" & WSFE.AppServerStatus & "<br><br>"
    WB1.Document.Write "dbserver status:" & WSFE.DbServerStatus & "<br><br>"
    WB1.Document.Write "authserver status:" & WSFE.AuthServerStatus & "<br><br>"
WB1.Document.parentWindow.scrollTo 0, 9999999
    DoEvents
    
    ' Recupera cantidad máxima de registros (opcional)
    Qty = WSFE.RecuperarQty()
    
    ' Recupera último número de secuencia ID
    Lastid = WSFE.UltNro()
    
    ' Recupera último número de comprobante para un punto de venta y tipo (opcional)
    Tipo_Cbte = CodFactAFIP
    
    If CodFactAFIP = 1 Or CodFactAFIP = 6 Then
        Punto_Vta = 3
    Else
        Punto_Vta = 4
    End If
    
    LastCbte = WSFE.RecuperaLastCMP(Punto_Vta, Tipo_Cbte)
    
Dim FECHA, ID, Tipo_Doc, Nro_Doc, cbt_Desde, Cbt_Hasta
Dim Imp_Tot_Conc, impto_liq_rni
Dim fecha_cbte, fecha_venc_pago, Presta_Serv, fecha_serv_desde, fecha_serv_hasta
Dim CAE, Tneto, Tbruto, Tiva, TIB

    ' Establezce los valores de la factura o lote a autorizar:
    FECHA = Format(Date, "yyyymmdd")
    ID = Lastid + 1: Presta_Serv = 0
    Tipo_Doc = 80: Nro_Doc = Format(ClteCUIT, "00000000")
    cbt_Desde = LastCbte + 1: Cbt_Hasta = LastCbte + 1
     Imp_Tot_Conc = "0.00"
     impto_liq_rni = "0.00"
    fecha_cbte = FECHA: fecha_venc_pago = FECHA
    ' Fechas del período del servicio facturado (solo si presta_serv = 1)
    fecha_serv_desde = FECHA: fecha_serv_hasta = FECHA
    
    TotBruto = Round(TotBruto, 2)
    TotIVAIns = Round(TotIVAIns, 2)
    ValIB = Round(ValIB, 2)
    TOTNETO = Round(TotBruto + TotIVAIns + ValIB, 2)
    Tneto = Replace(Format(TOTNETO, "######0.00"), ",", ".")
    Tbruto = Replace(Format(TotBruto, "######0.00"), ",", ".")
    Tiva = Replace(Format(TotIVAIns, "######0.00"), ",", ".")
    TIB = Replace(Format(ValIB, "######0.00"), ",", ".")
    
    ' Llama al WebService de Autorización para obtener el CAE
    CAE = WSFE.Aut(ID, Presta_Serv, _
        Tipo_Doc, Nro_Doc, Tipo_Cbte, Punto_Vta, _
        cbt_Desde, Cbt_Hasta, Tneto, Imp_Tot_Conc, Tbruto, _
        Tiva, impto_liq_rni, TIB, fecha_cbte, fecha_venc_pago)
 '       fecha_serv_desde, fecha_serv_hasta) ' si presta_serv = 0 no pasar estas fechas

MotRech = WSFE.Motivo

    WB1.Document.Write "Resultado: " & WSFE.Resultado & "<br><br>" & _
                 "Motivo de rechazo o advertencia: " & WSFE.Motivo & "<br><br>" & _
                 "Reprocesado?: " & WSFE.Reproceso & "<br><br>"
WB1.Document.parentWindow.scrollTo 0, 9999999
    DoEvents
    
    ' Verifica que no haya rechazo o advertencia al generar el CAE
    If CAE = "" Then
        MsgBox "La página esta caida o la respuesta es inválida"
    ElseIf CAE = "NULL" Or WSFE.Resultado <> "A" Then
        MsgBox "No se asignó CAE (Rechazado). Motivo: " & WSFE.Motivo, vbInformation + vbOKOnly
    ElseIf WSFE.Motivo <> "NULL" And WSFE.Motivo <> "00" Then
        MsgBox "Se asignó CAE pero con advertencias. Motivo: " & WSFE.Motivo, vbInformation + vbOKOnly
    End If
    
    
    WB1.Document.Write "QTY: " & Qty & "<br><br>"
    WB1.Document.Write "Cpbte.: " & cbt_Desde & "<br><br>"
    WB1.Document.Write "Comprobte.:" & LastCbte & "<br><br>"
    WB1.Document.Write "CAE: " & CAE & "<br><br>"
    WB1.Document.Write "Vencimiento: " & WSFE.Vencimiento & "<br><br>"
    
   FecVtoCAE = WSFE.Vencimiento
   NroCAE = CAE
   
    ' Imprime respuesta XML para depuración (errores de formato)
    WB1.Document.Write "Respuesta: " & WSFE.XmlResponse & "<br><br>"
    WB1.Document.parentWindow.scrollTo 0, 9999999
    DoEvents
    
    NROFAC = cbt_Desde
    
    If CAE <> "" Then
        AFIPbien = True
    End If
    
    Exit Sub
    
ManejoError:
    ' Si hubo error:
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select

'  afipbien=true
'  nrofac=777777
'  FecVtoCAE = "10-10-2023"
'  NroCAE = 11111111111111111


End Sub

Private Sub CalcDV()
Dim I1%, Suma As Double, Resul As Double, DV As Single

   
Dim Numero As String * 40
 
'Mid(Numero, 1, 11) = ClteCUIT
Mid(Numero, 1, 11) = "33703467909"
Mid(Numero, 12, 2) = Format(CodFactAFIP, "00")
Mid(Numero, 14, 4) = PtoVtaCpbte
Mid(Numero, 18, 14) = NroCAE
Mid(Numero, 32, 8) = Mid(FecVtoCAE, 7, 4) & Mid(FecVtoCAE, 4, 2) & Mid(FecVtoCAE, 1, 2)

' Etapa 1
For I1 = 1 To 39 Step 2
    Suma = Suma + Val(Mid(Numero, I1, 1))
Next I1

' Etapa 2
Resul = Suma * 3

' Etapa 3
Suma = 0
For I1 = 2 To 38 Step 2
    Suma = Suma + Val(Mid(Numero, I1, 1))
Next I1

' Etapa 4
Resul = Resul + Suma

' Busca DV
DV = Resul Mod 10
If DV = 0 Then GoTo MuestraCB

DV = 0

Do Until I1 = 0

DV = DV + 1
I1 = (Resul + DV) Mod 10

Loop

MuestraCB:

Mid(Numero, 40, 1) = DV


'BarCod.Caption = Numero

'Label7.Caption = DV

'BarCod.Width = 300
'BarCod.Height = 60

End Sub

Private Sub Obtener_QR()
Dim Texto As String, TipCpb As Integer
Dim Texto1 As String, Texto2 As String

If QRunaVez Then Exit Sub

QRunaVez = True

'TipCpb = TipoFac
'If LaLetra = "B" Then TipCpb = TipCpb + 5

Texto1 = "https://www.afip.gob.ar/fe/qr/?p=eyJ2ZXIiOjEsImZlY2hhIjoiMjAyMS0wMy0xNyIsImN1aXQiOjMzNzA4NzU1MzA5LCJwdG9WdGEiOjUsInRpcG9DbXAiOjEsIm5yb0NtcCI6MTUxMSwiaW1wb3J0ZSI6Mzg0NzguMDAsIm1vbmVkYSI6IlBFUyIsImN0eiI6MS4wMDAwMDAsInRpcG9Eb2NSZWMiOjgwLCJucm9Eb2NSZWMiOjMwNTIyNzgwNjA2LCJ0aXBvQ29kQXV0IjoiRSIsImNvZEF1dCI6NzExMTkzMTgyNjk2NTZ9"
Texto2 = ""
GoTo SaltaFactura

' Formato Nuevo
Texto = "{""ver"":1,"
Texto = Texto & """fecha"":" & Format(Date, "yyyy-mm-dd") & ","
Texto = Texto & """cuit"":33703467909,"
Texto = Texto & """ptoVta"":" & PtoVtaCpbte & ","
Texto = Texto & """tipoCmp"":" & CodFactAFIP & ","
Texto = Texto & """nroCmp"":" & Format(NROFAC, "00000000") & ","
Texto = Texto & """importe"":" & Format((TOTNETO * 100), "000000000000000") & ","
Texto = Texto & """moneda"":PES,"
Texto = Texto & """cotiz"":1,"
Texto = Texto & """tipoDocRec"":80,"
Texto = Texto & """nroDocRec"":" & Val(ClteCUIT) & ","
Texto = Texto & """tipoCodAut"":E,"
Texto = Texto & """codAut"":" & NroCAE & "}"

'  Esta rutina es para codificar en BASE64 (ahora está salteada)
Texto2 = Texto
cmdEncode Texto2

SaltaFactura:

Texto = Texto1 & Texto2

GoTo SaltaViejo

'   Formato Viejo
Texto = "ALESTEL SRL" & vbCrLf
Texto = Texto & "CUIT: 33-70346790-9" & vbCrLf
Texto = Texto & "Fecha: " & Format(Date, "ddd dd MMM YYYY") & vbCrLf
Texto = Texto & TIPO & "   " & LaLetra & vbCrLf
Texto = Texto & "Nro. " & PtoVtaCpbte & "-" & Format(NROFAC, "00000000") & vbCrLf
Texto = Texto & "Importe: " & Format(TOTNETO, " ###,##0.00") & vbCrLf
Texto = Texto & "Moneda: ($) Pesos" & vbCrLf
Texto = Texto & "C.A.E.: " & NroCAE & vbCrLf
Texto = Texto & "Fec.Vto. C.A.E.: " & FecVtoCAE & vbCrLf
Texto = Texto & "Cód.Cliente: " & CodCLTE & vbCrLf
Texto = Texto & "CUIT.Cliente: " & ClteCUIT

SaltaViejo:

Dim NomQR As String
NomQR = "c:\temp\" & PtoVtaCpbte & "-" & NROFAC & ".png"
'Shell ("qrcode.exe -o QR.png -l L -s 3 """ & Texto & ""), vbHide
Shell ("qrcode.exe -o " & NomQR & " -l L -s 3 """ & Texto & """  ")

Sleep 5000
        
'Picture2 = LoadPicture("QR.png")
    Picture3.Picture = LoadPicture("")
    pngClass.PicBox = Picture3
    pngClass.SetToBkgrnd False, 0, 0   'set to Background (True or false), x and y
    pngClass.BackgroundPicture = Picture3 'same Backgroundpicture
    pngClass.SetAlpha = True 'when Alpha then alpha
    pngClass.SetTrans = True 'when transparent Color then transparent Color
    
    ' Visualiza el Archivo en Picture2
    pngClass.OpenPNG NomQR

Sleep 2000

Kill NomQR

End Sub

'  En desuso
Private Sub Obtener_QR_xInternet()
Dim Texto As String

Texto = "ALESTEL SRL" & vbCrLf
Texto = Texto & "CUIT: 33-70346790-9" & vbCrLf
Texto = Texto & "Fecha: " & Format(Date, "ddd dd MMM YYYY") & vbCrLf
Texto = Texto & TIPO & "   " & LaLetra & vbCrLf
Texto = Texto & "Nro. " & PtoVtaCpbte & "-" & Format(NROFAC, "00000000") & vbCrLf
Texto = Texto & "Importe: " & Format(TOTNETO, " ###,##0.00") & vbCrLf
Texto = Texto & "Moneda: ($) Pesos" & vbCrLf
Texto = Texto & "C.A.E.: " & NroCAE & vbCrLf
Texto = Texto & "Fec.Vto. C.A.E.: " & FecVtoCAE & vbCrLf
Texto = Texto & "Cód.Cliente: " & CodCLTE & vbCrLf
Texto = Texto & "CUIT.Cliente: " & ClteCUIT

    Picture3.Picture = cQrCode.GetPictureQrCode(Texto, Picture3.ScaleWidth, Picture3.ScaleHeight)
    If Picture3.Picture Is Nothing Then MsgBox "Error!"
    'Picture1.Picture = cQrCode.GetPictureQrCode(Text1.Text, 200, 200, "UTF-8", "L", vbRed, vbBlue, 3)
End Sub

Private Sub cmdEncode(Texto As String)
   
    Set Bas64 = New Base64

    If flgFile = 0 Then
        strBuffer = Texto
    End If
    
    Debug.Print "Length of message = " & CStr(Len(strBuffer))
    
    If flgFile > 0 Then
        Bas64.bBuffer = byteBuffer
    Else
        Bas64.sBuffer = strBuffer
    End If
    
    Call Bas64.Base64Encode
    encBuffer = Bas64.Base64Buf
    
    If flgFile > 0 Then
        flgFile = flgFile + 1
        Texto = Texto & "Length of Encoded file = " & CStr(Len(encBuffer)) & vbCrLf
    Else
        Texto = encBuffer
        Debug.Print "Length of encoded message = " & CStr(Len(encBuffer))
        Call DebugPrintString("Encoded String", encBuffer)
    End If
    
End Sub

