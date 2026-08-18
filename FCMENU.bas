Attribute VB_Name = "FCMENUMOD"
Option Explicit

'---------------------------------------------------------
'     Parámetro Externos
'     ------------------
' /PV:(Pto. de Venta)[1]
' /IM:(salida de Impresión)[2]
' /BD:(Tipo de Base de Datos:A -Access- ó S -SQL-)[3]
' /SVR:(Servidor -SQL- ó Carpeta de Trabajo -Access-) [4]
'
' Son obligatorios "/BD:" Y "/SVR"
'---------------------------------------------------------
'      Consideraciones Vs.
'      -------------------
'
' - En el parámetro "/PV:" se debe poner "crear" para inicializar
'   el sistema
' - La Base de Datos se debe llamar "FCMENU.MDB" si es Access y
'   FCMENU si es SQL
' - El DSN (ODBC) se debe llamar FCMENU y se genera Automáticamente
' - Separa : Campo conteniendo el separador para Fecha.
'            En SQL es "'" y en Access es "#"
'---------------------------------------------------------


Public Ordenar       As String * 34
Public CodTabla      As String * 5
Public CodItem       As String
Public Correa        As String
Public RESTO         As String
Public Pulg          As Single
Public Mtros         As Single
Public Milim         As String * 2
Public Telas         As String * 2
Public NroCor        As String * 10
Public la_opcion      As String
Public LaOpcion      As Integer
Public CRITERIO       As String
Public Const comilla = """"

Public DesArt       As String
Public DesSecc      As String

Public MatrItems(1 To 25, 1 To 7, 1 To 2) As String

Public miSQL         As String
Public ORDENADOx     As String
Public ORDENADOad    As String
Public Descrip       As String * 100

Public LaLetra       As String * 1
Public TipoFac       As String * 1
Public Motivo        As String
Public CodMovs       As String
Public NroCpbte      As String
Public ConPercep     As Single

Public ApLetra       As String * 1
Public ApTipo        As String * 1
Public ApFecha       As Date
Public ApPtoVta      As Integer
Public ApCpbte       As Long
Public ApImpte       As Currency
Public LaCotiz       As Currency

Public SIGFIN1       As String
Public SIGFIN2       As String
Public SIGFIN3       As String
Public SIGFIN4       As String
Public paraLIKE      As String
Public Selecta       As String
Public Clasifica     As String
Public TotalREG      As Single
Public SelecREG      As Single
Public elCAMPO       As String
Public elCAMPO1      As String
Public elCAMPO2      As String
Public elCAMPO3      As String
Public elCAMPO4      As String
Public elCAMPO6      As String
Public elCAMPO7      As String
Public elCAMPO8      As String
Public elOPERA1      As String
Public elOPERA2      As String
Public elOPERA3      As String
Public elOPERA4      As String
Public elDATO1       As String
Public elDATO2       As String
Public elDATO3       As String
Public elDATO4       As String
Public laLOGIK1      As String
Public laLOGIK2      As String
Public laLOGIK3      As String
Public laCALCU       As String
Public NroCheq       As String * 10
Public LaTabla       As String * 10
Public elTOP         As String
Public paraSTATUS    As String
Public Respuesta     As String
Public LaFECHA       As Date
Public PosTOP        As Long
Public PosLEFT       As Long

Public CodCLTE       As Long
Public CodCLTEf      As Long
Public ClteNomb      As String
Public ClteClasif    As String
Public ClteDir       As String
Public ClteLoc       As String
Public ClteTel1      As String
Public ClteCPos      As String
Public CltePCIA      As String
Public ClteCIVA      As String
Public ClteCUIT      As String
Public ClteCVta      As String * 2
Public ClteVend      As String
Public ClteDeuda     As Currency
Public ClteCredit    As Currency

Public NroRec        As Long
Public totSAL        As Long
Public totENT        As Long
Public cpbSAL        As Integer
Public cpbENT        As Integer
Public Blanco        As String
Public miSQL1        As String
Public miSQL2        As String
Public miSQL3       As String
Public i            As Integer
Public SaldoClte    As Double
Public Total        As Double
Public Precio       As Double
Public Importe      As Double
Public Cantid       As Double
Public Descto       As Double
Public MESANIO      As String
Public LaTecla      As Byte
Public CIVA         As Byte
Public Origen       As Byte
Public TopRow       As Integer
Public LaFila       As Integer
Public LaPos        As Integer
Public Lista        As String
Public HayNotas     As Boolean
Public HayOP        As Boolean
Public HayCA        As Boolean
Public Habilita     As Boolean
Public AlFin        As Boolean
Public VerDetArt    As Boolean
Public VerDetClt    As Boolean
Public VeNotaC      As Boolean
Public HOMO         As Boolean    '  Tipo de facturacion Homologación=True Producción=False
'Public EstaFactu    As Boolean
Public ParaBcos     As Boolean
Public EsProvisorio As Boolean
Public SinCpbte     As Boolean
Public SALIR        As Boolean
Public EnDolares    As Boolean
Public TIPOMov      As Byte
Public TipoMov2     As Byte
Public TipoMov3     As Byte
Public DeDonde      As Byte
Public VolverArt    As Byte
Public ParaOP       As Byte
Public CantOP       As Integer
Public Promed       As Double
Public Prom2mes     As Double
Public PCIA

Public TipoVta      As Integer
Public IVAIns       As Currency
Public IVANI        As Currency
Public NomEmpr      As String
Public MontoDeuda   As Currency
Public DiasDeuda    As Integer

Public Dtos(1 To 5) As Single
Public Rgos(1 To 3) As Single

'Salida de Impresión
Public SaleImpr      As String
Public Params(1 To 10) As String
Public TipoDB        As String
Public Separa        As String
Public ServidorSQL   As String
Public PtoVta        As String * 4
Public PtoVtaGrab    As String * 4

Public NomZIP        As String
Public UsuEmail      As String
Public PassEmail     As String
Public ArchEmail     As String
Public HoraEmail     As String
Public Realizado     As String
Public FuerzaBack    As Boolean

Public dbTABL As New ADODB.Connection
Public RgTABL As New ADODB.Recordset

Public vbMsgBoxTitle As String
Public vbMsgBoxText As String
Public vbMsgBoxResp As Integer
Public vbResponse As Integer

'Acceso a la Red
Public AutOK                As Boolean
'Public VNet                  As DatNet
Public Red_Usuario       As String
Public Red_SOPlataf     As String
Public Red_SOVersion  As String
Public Red_Maq           As String
Public Red_Grupo(7)    As String
Public Red_Dom          As String
Public Red_Server       As String
Public Red_Privil          As String

' Llamadas a SQL

Public BDatos1 As String
Public BDatos2 As String

Public Const laselec1a = "SELECT  * FROM ARTIprec"
Public Const laselec1b = "SELECT  * FROM ARTICULO WHERE ORDEN ="
Public Const laselec2 = "CREATE VIEW dbo.VERART1" & _
                                            " AS SELECT TOP 3000 * FROM VERART1"
Public Const laselec3 = " INNER JOIN " & _
    "Precios ON Articulo.PROVALT = Precios.PROV AND " & _
    "Articulo.ORDEN = Precios.ORDEN AND " & _
    "Articulo.CORR = Precios.CORR INNER JOIN " & _
    "Stock ON Articulo.ORDEN = Stock.ORDEN AND " & _
    "Articulo.CORR = Stock.CORR"

Public Const laselec1 = "SELECT CODIGO, RSOCIAL FROM Proveed "

Public Const laselec4 = "SELECT TOP "
Public Const laselec6 = " AND COD2 = "
Public Const laselec7 = " AND FECHA < "
Public Const laselec8 = "FROM FCEstad1 WHERE COD1 = "
Public Const laselec9 = " FROM VERSALI WHERE COD1 = "
Public Const laselec10 = " AND ANIOMES <= '"
Public Const laselec11 = " AND MES <= '"


Public Const laselec20 = "SELECT CODIGO, RSOCIAL, LOC, TEL1, CUIT FROM Clientes "
Public Const laselec21 = "FROM ESTART WHERE CLTE = "
Public Const laselec40 = "SELECT FECHA, TIPO, CPBTE, I1, I2, I3, I4, I5, RESTO, DEBE, HABER " _
                                    & " FROM VERCCTE WHERE CLTE = "

Public Const elWHERE = " WHERE "
Public Const elORDEN = " ORDER BY "

'   API para manejar intervalos
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub Main()
Dim ix As Double
'On Error GoTo VerError

If App.PrevInstance Then
    MsgBox "El Facturador de Cta.Cte. ya se esta ejecutando"
    End
End If

BuscaParamExt

' Pto. de Vta. Parametro 1
' PTOVTA = "0023"
PtoVta = Params(1)
    If PtoVta = "" Or PtoVta = "    " Then
        MsgBox "No Hay Punto de Venta ó" & vbCrLf & "     el Punto de Venta" & vbCrLf _
        & "        es Incorrecto" & vbCrLf & "                (" & PtoVta & ")", , "Facturador - Aviso"
        End
    End If

' SaleImpr = Environ("PRINTER")
SaleImpr = Params(2)
    'If Params(2) = "" Then
    '    SaleImpr = "LPT1:"
    'End If

TipoDB = Params(3)

    'If Params(4) = "" Then
    '    Set VNet = New DatNet
        'ServidorSQL = "(local)"
        'ServidorSQL = "SERVER-NT"
    'Else
    '    ServidorSQL = Params(4)
    'End If

ServidorSQL = Params(4)

' Si es la primera vez genera las Bases de Datos
If Params(1) = "crear" Then
    ZZGenDB.Show
    Exit Sub
End If

If Params(6) = "PROD" Then
    HOMO = False
Else
    HOMO = True
    End If

' Si es la primera vez genera las Bases de Datos
If Params(1) = "crear" Then
    ZZGenDB.Show
    Exit Sub
End If

Inicio.Show
'  mF.LicenseKey = "126L57L99ZT38D9"

'Red_SOPlataf = VNet.OSPlataf
'Red_Usuario = VNet.UserName
'Red_Maq = VNet.ComputerName
'Red_Server = VNet.Server
'Red_Dom = VNet.LogonDomain
'For i = 0 To VNet.GroupCount
'Red_Grupo(i) = VNet.Group(i)
'Next


TipoFac = 1

dbTABL.ConnectionString = BDatos1
dbTABL.Open

If Err.Number = 2333 Then Paramet.Show

miSQL2 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    
    If RgTABL.EOF Then
        MsgBox "El porcentaje de IVA " & vbCrLf & "    NO está cargado    " & vbCrLf _
        & " " & vbCrLf & "        Verifique ... ", , "FCMENU - Aviso"
        End
    End If


    IVAIns = RgTABL!IVAIns / 100
    IVANI = RgTABL!IVANI / 100
    
    NomEmpr = RgTABL!NomEmpr
    DiasDeuda = RgTABL!nume20
    MontoDeuda = RgTABL!limvta

If IsNull(RgTABL!mcaib) Then
    ConPercep = 2
Else
    ConPercep = RgTABL!mcaib
End If

RgTABL.Close

PCIA = Array("ASalta    ", "BBs. As.  ", "CCap.Fed. ", "DSan Luis ", "EEnt.Ríos ", "FLa Rioja ", "GSgo.del E", "HChaco    ", _
"I         ", "JSan Juan ", "KCatamarca", "LLa Pampa ", "MMendoza  ", "NMisiones ", "O         ", _
"PFormosa  ", "QNeuquén  ", "RRío Negro", "SSanta Fe ", "TTucumán  ", "UChubut   ", "VT.del Fgo", "WCorr'tes.", "XCórdoba  ", _
"YJujuy    ", "ZSta.Cruz ", "          ")

FCMENU.Show

Exit Sub

VerError:

If Err.Number = -2147467259 Then Paramet.Show

End Sub

Function Escribe(NomFont As String, SizeFont As Single, Negr As Boolean, Ital As Boolean, Lin As Single, Col As Single, Texto As String)
        
Texto = Replace(Texto, vbCrLf, "")
        
IMPRE.VP1.FontName = "Arial"
IMPRE.VP1.FontSize = 10
IMPRE.VP1.FontBold = False
IMPRE.VP1.FontItalic = False
        
        If NomFont <> "" Then IMPRE.VP1.FontName = NomFont
        If SizeFont > 0 Then IMPRE.VP1.FontSize = SizeFont
        If Negr Then IMPRE.VP1.FontBold = True
        If Ital Then IMPRE.VP1.FontItalic = True

IMPRE.VP1.CurrentY = str(Lin) & "mm"
IMPRE.VP1.CurrentX = str(Col) & "mm"
IMPRE.VP1.Text = Texto

End Function

Function DibujaRect(LinDsd As Single, ColDsd As Single, LinHst As Single, ColHst As Single, Rad1 As Single, Rad2 As Single)

IMPRE.VP1.DrawRectangle str(ColDsd) & "mm", str(LinDsd) & "mm", str(ColHst) & "mm", str(LinHst) & "mm", Rad1, Rad2

End Function

Function DibujaLine(LinDsd As Single, ColDsd As Single, LinHst As Single, ColHst As Single)

IMPRE.VP1.DrawLine str(ColDsd) & "mm", str(LinDsd) & "mm", str(ColHst) & "mm", str(LinHst) & "mm"

End Function

Function BuscaEnCombo(ElCombo As ComboBox, ItCal As String) As Integer
Dim I1

For I1 = 0 To ElCombo.ListCount - 1

ElCombo.ListIndex = I1

If Val(Mid(ElCombo.Text, 1, 2)) = ItCal Then
    BuscaEnCombo = I1
   Exit Function
End If

Next

End Function

Private Sub BuscaParamExt()
Dim Args As String, I1%, i2%

Args = Command()

If Args = "" Then
    MsgBox "      No Hay Parámetros o   " & vbCrLf & "         son Incorrectos      " & vbCrLf _
    & "        Imposible Continuar    ", , " Facturador - Error Fatal "
    End
End If

For I1 = 1 To 10: Params(I1) = "": Next

' Pto. de Venta (Parametro 1)
I1 = InStr(1, Args, "/PV:") + 4
i2 = InStr(I1, Args, " ") - I1
If i2 < 1 Then i2 = Len(Args) - (I1 - 1)

If I1 > 4 Then
        Params(1) = Mid(Args, I1, i2)
End If
    
' Salida de Impresión (Parametro 2)
I1 = InStr(1, Args, "/IM:") + 4
i2 = InStr(I1, Args, " ") - I1
If i2 < 1 Then i2 = Len(Args) - (I1 - 1)

If I1 > 4 Then
        Params(2) = Mid(Args, I1, i2)
End If
    
' Tipo de Base de Datos : A-Access / S-SQL  (Parametro 3)
I1 = InStr(1, Args, "/BD:") + 4
i2 = InStr(I1, Args, " ") - I1
If i2 < 1 Then i2 = Len(Args) - (I1 - 1)

If I1 > 4 Then
        Params(3) = Mid(Args, I1, i2)
End If
    
' Servidor  (Parametro 4)
I1 = InStr(1, Args, "/SVR:") + 5
i2 = InStr(I1, Args, " ") - I1
If i2 < 1 Then i2 = Len(Args) - (I1 - 1)

If I1 > 5 Then
        Params(4) = Mid(Args, I1, i2)
End If

' Carpeta de  (Parametro 5)
I1 = InStr(1, Args, "/CAR:") + 5
i2 = InStr(I1, Args, " ") - I1
If i2 < 1 Then i2 = Len(Args) - (I1 - 1)

If I1 > 5 Then
        Params(5) = Mid(Args, I1, i2)
End If

' Tipo de Facturación (Homologación o Producción)  (Parametro 6)
I1 = InStr(1, Args, "/TF:") + 4
i2 = InStr(I1, Args, " ") - I1
If i2 < 1 Then i2 = Len(Args) - (I1 - 1)

If I1 > 5 Then
        Params(6) = Mid(Args, I1, i2)
End If

End Sub

'-------------------------------------------
'    Cálculo de Digitpo Verificador del CUIT
'       Sin separadores (sin guiones)
'-------------------------------------------
Function Cuit(strCuit As String) As Boolean

Dim intBase(9) As Integer
Dim intNumero(9) As Integer
Dim intCalculo As Integer
Dim intVerificador As Integer
Dim i As Integer
Dim x As Integer

Cuit = False

x = 5
For i = 0 To 3
intBase(i) = x
x = x - 1
Next
x = 7
For i = 4 To 9
intBase(i) = x
x = x - 1
Next
    If Len(strCuit) <> 11 Then
        Exit Function
    End If

For i = 0 To 9
intNumero(i) = Mid(strCuit, i + 1, 1)
intCalculo = intCalculo + (intNumero(i) * intBase(i))
Next

intCalculo = 11 - (intCalculo - (Int(intCalculo / 11) * 11))

Select Case intCalculo
Case 11
intCalculo = 0
Case 10
intCalculo = 9
End Select

If intCalculo = Mid(strCuit, 11, 1) Then
    Cuit = True
End If

End Function

Function MontoEscrito(tyCantidad As Currency) As String
Dim lyCantidad As Currency, lyCentavos As Currency, lnDigito As Byte, lnPrimerDigito As Byte, lnSegundoDigito As Byte, lnTercerDigito As Byte, lcBloque As String, lnNumeroBloques As Byte, lnBloqueCero
Dim laUnidades As Variant, laDecenas As Variant, laCentenas As Variant
tyCantidad = Round(tyCantidad, 2)
lyCantidad = Int(tyCantidad)
lyCentavos = (tyCantidad - lyCantidad) * 100
laUnidades = Array("UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIESISEIS", "DIESISIETE", "DIESIOCHO", "DIESINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
laDecenas = Array("DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
laCentenas = Array("CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
lnNumeroBloques = 1
Do
lnPrimerDigito = 0
lnSegundoDigito = 0
lnTercerDigito = 0
lcBloque = ""
lnBloqueCero = 0
For i = 1 To 3
lnDigito = lyCantidad Mod 10
If lnDigito <> 0 Then
Select Case i
Case 1
lcBloque = " " & laUnidades(lnDigito - 1)
lnPrimerDigito = lnDigito
Case 2
If lnDigito <= 2 Then
lcBloque = " " & laUnidades((lnDigito * 10) + lnPrimerDigito - 1)
Else
lcBloque = " " & laDecenas(lnDigito - 1) & IIf(lnPrimerDigito <> 0, " Y", Null) & lcBloque
End If
lnSegundoDigito = lnDigito
Case 3
lcBloque = " " & IIf(lnDigito = 1 And lnPrimerDigito = 0 And lnSegundoDigito = 0, "CIEN", laCentenas(lnDigito - 1)) & lcBloque
lnTercerDigito = lnDigito
End Select
Else
lnBloqueCero = lnBloqueCero + 1
End If
lyCantidad = Int(lyCantidad / 10)
If lyCantidad = 0 Then
Exit For
End If
Next i
Select Case lnNumeroBloques
Case 1
MontoEscrito = lcBloque
Case 2
MontoEscrito = lcBloque & IIf(lnBloqueCero = 3, Null, " MIL") & MontoEscrito
Case 3
MontoEscrito = lcBloque & IIf(lnPrimerDigito = 1 And lnSegundoDigito = 0 And lnTercerDigito = 0, " MILLON", " MILLONES") & MontoEscrito
End Select
lnNumeroBloques = lnNumeroBloques + 1
Loop Until lyCantidad = 0
'MontoEscrito = "(" & MontoEscrito & IIf(tyCantidad > 1, " PESOS ", " PESO ") & Format(Str(lyCentavos), "00") & "/100 )"
MontoEscrito = MontoEscrito & " con " & Format(str(lyCentavos), "00") & "/100"
End Function

Function SacaDec(Numero As String) As String
Dim I1%

SacaDec = Numero
I1 = InStr(1, Numero, ",")
If I1 = 0 Then Exit Function
SacaDec = Mid(Numero, 1, I1 - 1) & Mid(Numero, I1 + 1, Len(Numero))

End Function

Sub DxClte(CLTE As Long, Secc As String)
Dim I1%

For I1 = 1 To 5

Dtos(I1) = 0
If I1 < 4 Then Rgos(I1) = 0

Next I1

miSQL = "SELECT * FROM DtoxClte WHERE CLTE = " & CLTE & _
        " and SECCION = '" & Secc & "'"
RgTABL.Open miSQL, dbTABL, adOpenForwardOnly, adLockReadOnly
   
If RgTABL.EOF Then
    RgTABL.Close
    Exit Sub
End If
    
Dtos(1) = RgTABL!dto1
Dtos(2) = RgTABL!dto2
Dtos(3) = RgTABL!dto3
Dtos(4) = RgTABL!dto4
Dtos(5) = RgTABL!dto5

Rgos(1) = RgTABL!Rgo1
Rgos(2) = RgTABL!Rgo2
Rgos(3) = RgTABL!Rgo3
    
RgTABL.Close

End Sub

Function ImprMatri(Area As String, Salto As Single)
Dim I1 As Single

I1 = 0
'Salto = Salto - 1

Print #1, Area;

Do Until I1 = Salto

I1 = I1 + 1

Print #1, ""

Loop

End Function

Sub SlideWindow2(frmSlide As Form, iSpeed As Integer)

While frmSlide.Left + frmSlide.Width < Screen.Width
    DoEvents
    frmSlide.Left = frmSlide.Left + iSpeed
Wend

While frmSlide.top - frmSlide.Height < Screen.Height
    DoEvents
    frmSlide.top = frmSlide.top + iSpeed
Wend

Unload frmSlide

End Sub

'Sub Startrek(frm As Form)
Sub SlideWindow(frm As Form, iSpeed As Integer)
Dim GOTOVAL%, GOINTO

GOTOVAL = frm.Height / 2
For GOINTO = 1 To GOTOVAL
    DoEvents
    frm.Height = frm.Height - 100
    frm.top = (Screen.Height - frm.Height) \ 2
    If frm.Height <= 500 Then Exit For
Next GOINTO

horiz:
frm.Height = 30
GOTOVAL = frm.Width / 2
    For GOINTO = 1 To GOTOVAL
    DoEvents
    frm.Width = frm.Width - 100
    frm.Left = (Screen.Width - frm.Width) \ 2
    If frm.Width <= 2000 Then Exit For
Next GOINTO
Unload frm

End Sub
'just paste this procedure and call it
Sub ExplodeForm(frm As Form)
    Dim x As Long
    Dim factor  As Double
    Dim Width As Integer
    Dim Height As Integer
    
    Height = frm.Height
    Width = frm.Width
    factor = Height / Width
    frm.Width = 0
    frm.Height = 0
    frm.Show
    
    For x = 0 To Width Step 215
        frm.Width = x
        frm.Height = x * factor
        If Not frm.MDIChild Then
            With frm
                .Left = (Screen.Width - .Width) / 2
                .top = (Screen.Height - .Height) / 2
            End With
        End If
    Next x
        
        frm.Width = Width
        frm.Height = Height
    
End Sub
