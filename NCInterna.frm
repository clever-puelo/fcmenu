VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Begin VB.Form NCInterna 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nota de Crédito Interna"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1080
      Left            =   4980
      Picture         =   "NCInterna.frx":0000
      ScaleHeight     =   1020
      ScaleWidth      =   3345
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4110
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Otro Clte."
      Height          =   330
      Left            =   360
      TabIndex        =   4
      Top             =   3750
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   510
      TabIndex        =   9
      Top             =   2400
      Width           =   7305
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Height          =   585
         Left            =   1410
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   5805
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
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   150
         Width           =   4320
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nota:"
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1110
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo:"
         Height          =   240
         Index           =   1
         Left            =   270
         TabIndex        =   10
         Top             =   180
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   300
      TabIndex        =   5
      Top             =   900
      Width           =   7695
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Deuda Total:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   510
         Left            =   1680
         TabIndex        =   8
         Top             =   750
         Width           =   5205
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   390
         Width           =   1110
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   330
      Left            =   6720
      TabIndex        =   3
      Top             =   3750
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5460
      TabIndex        =   2
      Top             =   3750
      Width           =   1140
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   2100
      Left            =   750
      TabIndex        =   13
      Top             =   3750
      Visible         =   0   'False
      Width           =   6015
      _cx             =   10610
      _cy             =   3704
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   14352126
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16576
      ForeColorSel    =   12640511
      BackColorBkg    =   8438015
      BackColorAlternate=   12640511
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"NCInterna.frx":13D6
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSPDF8LibCtl.VSPDF8 VSPDF81 
      Left            =   7890
      Top             =   3390
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
      Caption         =   "Número:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   6120
      TabIndex        =   12
      Top             =   270
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nota de Crédito Interna"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   510
      Left            =   570
      TabIndex        =   11
      Top             =   180
      Width           =   5205
   End
End
Attribute VB_Name = "NCInterna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset
Dim RgCCTE As New ADODB.Recordset
Dim NroNCI As Long, Saldo As Currency, Corr As Integer, LaNota As String

Private Sub Command3_Click()

Label1 = "": Label2 = ""
Command3.Visible = False
Command1.Enabled = False
Combo5.ListIndex = -1
BusClte.Show

End Sub

Private Sub Form_Load()

Me.Move 0, 0, 8300, 4700

DeDonde = 7
NroNCI = 0

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

    miSQL1 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
    RgTABL.Open miSQL1, dbCLTE, adOpenDynamic, adLockOptimistic
    If Not RgTABL.EOF Then
            NroNCI = RgTABL!nume10 + 1
    End If

        RgTABL.Close

Label5 = "Número: " & Format(NroNCI, "###,##0")

CargaMoti

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbCLTE.Close

End Sub
'*************************************
'   Ingresa Motivo de N.Crédito
'*************************************

Private Sub Command1_Click()

AGrabar
unload Me

End Sub

Private Sub Command2_Click()

unload Me

End Sub

Private Sub Combo5_Click()

'Command1.Enabled = True

End Sub

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

Command1.Enabled = True

End Sub

Private Sub Text1_Change()

If Len(Text1.Text) > 6 Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If

LaNota = Text1.Text

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

Combo5.ListIndex = -1

End Sub

Public Sub BuscaDeuda()
Dim I1 As Integer

I1 = 0

FG1.Clear flexClearScrollable
FG1.Rows = 1

Saldo = 0

miSQL = "SELECT * FROM CtasCtes Where CLTE=" & CodCLTE & " and (tipo=1 or tipo=3) and debe>0"
RgCCTE.Open miSQL, dbCLTE, adOpenForwardOnly, adLockReadOnly

If RgCCTE.EOF Then
    RgCCTE.Close
    Exit Sub
    End If

Do Until RgCCTE.EOF


FG1.AddItem RgCCTE!TIPO & vbTab & RgCCTE!FECHA & vbTab & RgCCTE!Cpbte & vbTab & _
            RgCCTE!DEBE & vbTab & RgCCTE!Impte, FG1.Rows

Saldo = Saldo + RgCCTE!DEBE

RgCCTE.MoveNext

Loop

RgCCTE.Close

If Saldo > 0 Then
        Label1.Caption = "Deuda Total: " & Format(Saldo, " $ ###,##0.00 ")
End If

End Sub

Private Sub AGrabar()
Dim I1%, Codi1 As String, Codi2 As String
Dim Mensaje As String

Mensaje = "Desea GRABAR la Nota de Crédito ?" & vbCrLf & _
         "(Esta acción cancelará la deuda del cliente)"
Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "Cancelación de Deuda")

If Respuesta = vbNo Then
    GoTo AlFinal
    End If
    
If Respuesta = vbCancel Then
         Exit Sub
         End If

ImprimeNCI

dbCLTE.BeginTrans

'  Lee y actualiza Correlativo en Clientes

miSQL1 = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE
RgCLTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If RgCLTE.EOF Then
    RgCLTE.Close
    Exit Sub
    End If

Corr = Val(RgCLTE!corr1) + 1
If Corr = 0 Or Corr > 99 Then Corr = 1

    RgCLTE!corr1 = Corr
    RgCLTE!DEUDA = 0
    
RgCLTE!USUARIO = Mid(Red_Usuario, 1, 6)
RgCLTE!FACTUAL = Date
    
    RgCLTE.Update

'  Graba Número de Recibo

    miSQL1 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
    RgTABL.Open miSQL1, dbCLTE, adOpenDynamic, adLockOptimistic
    If RgTABL.EOF Then
        RgTABL.Close
        Exit Sub
    End If
            
            RgTABL!nume10 = NroNCI

        RgTABL.Update
        RgTABL.Close

    miSQL2 = "SELECT  * FROM CtasCtes WHERE clte = " & CodCLTE & _
                         " and (Tipo=1 or Tipo=3) and debe>0"

RgCCTE.Open miSQL2, dbCLTE, adOpenDynamic, adLockPessimistic

    If RgCCTE.EOF = True Then
        RgCCTE.Close
        dbCLTE.RollbackTrans
        GoTo AlFinal
        End If

Do Until RgCCTE.EOF
    
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

       RgCCTE!DEBE = 0

RgCCTE!USUAR = Mid$(Red_Usuario, 1, 6)

RgCCTE.Update
RgCCTE.MoveNext

Loop

RgCCTE.Close

'-----------------------------------------------
'                Graba la Nota de Crédito
'-----------------------------------------------

laLOGIK1 = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
miSQL1 = "SELECT * FROM CtasCtes WHERE CLTE = " & CodCLTE & " AND CPBTE= " & NroNCI & _
                                " AND FECHA = " & Separa & laLOGIK1 & Separa & " AND TIPO= " & 8
RgCCTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If Not RgCCTE.EOF Then
    RgCCTE.Close
    GoTo AlFinal
    End If

RgCCTE.AddNew

RgCCTE!prefijo = 1
RgCCTE!Cpbte = NroNCI
RgCCTE!FECHA = Date
RgCCTE!TIPO = 8
RgCCTE!CLTE = CodCLTE

If Corr > 99 Then Corr = 0

RgCCTE!imput1 = Corr
RgCCTE!imput2 = 0
RgCCTE!imput3 = 0
RgCCTE!imput4 = 0
RgCCTE!imput5 = 0
RgCCTE!imput6 = 0

RgCCTE!DEBE = Saldo
RgCCTE!Impte = Saldo

RgCCTE!USUAR = Mid(Red_Usuario, 1, 6)
RgCCTE!moti = Combo5.ItemData(Combo5.ListIndex)
RgCCTE!tipo9 = 0
RgCCTE!CVTA = 0
RgCCTE!bon = 0
RgCCTE!FECVTO = Date
RgCCTE!Letra = "X"
    
    RgCCTE.Update

dbCLTE.CommitTrans

AlFinal:

End Sub

Private Sub ImprimeNCI()

Dim I1%, i2%, i3%, LaLetra As String, CodAntic As Integer, OtraVez As Boolean
Dim s$, h$, d$, ICpbte As String * 7, Total1 As Currency, Total2 As Currency
Dim IFec1 As String * 11, ITipo As String * 15, IImpte1 As String * 12
Dim IBco As String * 20, ICheq As String * 12, IImpte2 As String * 12, IFec2 As String * 11
Dim MasLin As Integer

LaLetra = "X"

IMPRE.VP1.Visible = True
IMPRE.VP1.PaperSize = pprA4
IMPRE.VP1.Orientation = orPortrait
IMPRE.VP1.BrushStyle = 1
IMPRE.VP1.Preview = True
IMPRE.VP1.AbortWindow = True
IMPRE.VP1.MarginLeft = "1mm"
IMPRE.VP1.MarginRight = "1mm"
IMPRE.VP1.MarginBottom = "1mm"
IMPRE.VP1.MarginHeader = "1mm"
'IMPRE.VP1.Collate = colTrue
'IMPRE.VP1.Copies = 2
        
 IMPRE.VP1.StartDoc

OtraVez:
    
 IMPRE.VP1.DrawPicture Picture1, "8mm", (6 + MasLin) & "mm"
Escribe "ARIAL BLACK", 18, 0, 0, 15.6 + MasLin, 66, "S.R.L."
 
IMPRE.Caption = " Impresión de la Nota de Crédito Interna : " & LaLetra & " Nº " & "" & _
                    Format(NroNCI, "000000") & " (Cta.Cte.)"

'Primer Rectángulo
DibujaRect 5 + MasLin, 7, 130 + MasLin, 200, 30, 30
DibujaLine 31 + MasLin, 7, 31 + MasLin, 200
DibujaLine 46 + MasLin, 7, 46 + MasLin, 200
DibujaLine 53 + MasLin, 7, 53 + MasLin, 200
DibujaLine 105 + MasLin, 7, 105 + MasLin, 200

DibujaLine 145, 5, 145, 200

'  Línea 1
'Escribe "ARIAL", 16, 1, 0, 6, 8, "ALESTEL S.R.L."
Escribe "ARIAL BLACK", 24, 1, 0, 6 + MasLin, 104, LaLetra
Escribe "", 10, 0, 1, 9 + MasLin, 132, "Buenos Aires, " & Format(Date, "ddd dd MMM YYYY")
'  Línea 2
Escribe "ARIAL", 9, 1, 0, 23 + MasLin, 14, "Av. Crovara 2948 - (1766) La Tablada (BA)"
Escribe "ARIAL", 7, 0, 0, 17 + MasLin, 99, "Documento no"
Escribe "", 12, 1, 1, 14 + MasLin, 130, "Nota de Crédito"
Escribe "", 12, 1, 1, 16 + MasLin, 166, "Nº " & Format(NroNCI, "000000")
'Escribe "", 9, 0, 0, 30, 153, Format(Time, "hh:mm")
'  Línea 3
'Escribe "ARIAL", 9, 1, 0, 26, 20, "(1235) La Tablada"
Escribe "ARIAL", 7, 0, 0, 20 + MasLin, 96, "válido como Factura"
Escribe "", 12, 1, 1, 18 + MasLin, 137, "Interna"
Escribe "", 8, 0, 0, 24 + MasLin, 130, "Imp.Int.: No Resp.     I.B.: 901 33-70346790-9"
'  Línea 4
Escribe "ARIAL", 9, 1, 0, 27 + MasLin, 14, "Tel. (011) 4652-1040 / 2684 / 2689 / 3080"
Escribe "", 8, 0, 0, 27 + MasLin, 130, "Part.Munic.: 1985   CUIT: 33-70346790-9"
'  Línea 5
Escribe "", 10, 1, 1, 32 + MasLin, 8, "Sr./es: (" & CodCLTE & ") " & ClteNomb

'  Línea 6
Escribe "", 10, 0, 0, 37 + MasLin, 8, ClteDir
'  Línea 7
Escribe "", 10, 0, 0, 42 + MasLin, 8, "(" & ClteCPos & ") " & ClteLoc & " " & CltePCIA
Escribe "", 10, 1, 1, 42 + MasLin, 100, "IVA: " & CabRec.Combo1.List(ClteCIVA - 1) & _
                              "         CUIT : " & ClteCUIT

'  Línea 8   -  Títulos del Detalle
Escribe "", 8, 1, 1, 46 + MasLin, 80, "Comprobantes Cancelados"
'  Línea 9
Escribe "", 8, 1, 1, 49 + MasLin, 11, "Tipo"
Escribe "", 8, 1, 1, 49 + MasLin, 31, "Nro."
Escribe "", 8, 1, 1, 49 + MasLin, 67, "Fecha"
Escribe "", 8, 1, 1, 49 + MasLin, 123, "Importe"
Escribe "", 8, 1, 1, 49 + MasLin, 162, "Deuda"

'----------------------------------------------
'           Detalle
'----------------------------------------------

' ---- Facturas ----

IMPRE.VP1.FontName = "Lucida Console"
IMPRE.VP1.FontItalic = False
IMPRE.VP1.CurrentX = "0mm"
IMPRE.VP1.FontSize = 9
IMPRE.VP1.TableBorder = tbNone

s$ = "~<+20mm|~<+30mm|~<+40mm|~>+40mm|~>+40mm"
h$ = "Tipo|Nro.|Fecha|Aplicado|Importe"

IMPRE.VP1.CurrentY = (54 + MasLin) & "mm"
IMPRE.VP1.CurrentX = 0
IMPRE.VP1.TextAlign = 0
IMPRE.VP1.MarginLeft = "9mm"

Total1 = 0

For i2 = 1 To FG1.Rows - 1

'I3 = DetRec.FG1.Row(I1)
If CCur(FG1.TextMatrix(i2, 3)) = 0 Then GoTo Saltando1
I1 = I1 + 1

Select Case FG1.TextMatrix(i2, 0)

    Case 1: ITipo = "Factura"
    Case 2: ITipo = "N.Crédito"
    Case 3: ITipo = "Nota de Débito"

End Select

ICpbte = Val(FG1.TextMatrix(i2, 2))
IFec1 = Format(FG1.TextMatrix(i2, 1), "dd mmm yyyy")
IImpte1 = Format(FG1.TextMatrix(i2, 4), "$ ###,##0.00")
IImpte2 = Format(FG1.TextMatrix(i2, 3), "$ ###,##0.00")

d$ = ITipo & "|" & ICpbte & "|" & IFec1 & "|" & IImpte1 & "|" & IImpte2

IMPRE.VP1.FontSize = 8
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 1
IMPRE.VP1.AddTable s$, h$, " || | |", , , True

Total1 = Total1 + FG1.TextMatrix(i2, 3)

Saltando1:

Next i2

Escribe "", 9, 0, 0, 101 + MasLin, 132, " Total Cancelado: " & Format(Total1, "$ ###,##0.00")

Escribe "", 10, 1, 1, 107 + MasLin, 8, " Total: " & Format(Total1, "$ ###,##0.00")
Escribe "", 10, 1, 1, 107 + MasLin, 50, " Son Pesos:" & MontoEscrito(Total1)

Escribe "", 8, 0, 1, 120 + MasLin, 150, "___________________"
Escribe "", 8, 0, 1, 125 + MasLin, 154, "   Confeccionó"

Escribe "", 8, 0, 0, 116 + MasLin, 8, "Motivo de la Nota de Crédito Interna: " & Combo5.Text
Escribe "", 8, 0, 0, 120 + MasLin, 8, "Nota: " & Mid$(LaNota, 1, 50)
Escribe "", 8, 0, 0, 124 + MasLin, 8, "          " & Mid$(LaNota, 51, 50)

If Not OtraVez Then
    Escribe "", 8, 1, 1, 132, 160, "Original"
Else
    Escribe "", 8, 1, 1, 132 + MasLin, 160, "Duplicado"
End If
 
 If Not OtraVez Then                     '  Esto estaba cuando salian 2 hojas
    MasLin = 145
    OtraVez = True
    GoTo OtraVez
 End If

 IMPRE.VP1.EndDoc
 Dim pepe

EsProvisorio = True

'Genera PDF
'VSPDF81.Title = "c:\Factu\Comprob\" & Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "\NCI-" & LaLetra & "-" & NroNCI & ".pdf"
VSPDF81.Title = "NCI-X-" & NroNCI & "-"

'    VSPDF81.Title = VSPDF81.Title & "(Hora " & Format(Time, "hh mm") & ")"

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

 'IMPRE.Show 1
 
     OtraVez = False
    MasLin = 0

End Sub
