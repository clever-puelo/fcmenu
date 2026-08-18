VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Begin VB.Form EmiRec 
   BackColor       =   &H00FEFADE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                                   Registra el Recibo"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8430
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
   ScaleHeight     =   5130
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1080
      Left            =   330
      Picture         =   "EmiRec.frx":0000
      ScaleHeight     =   1020
      ScaleWidth      =   3345
      TabIndex        =   5
      Top             =   3975
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volver"
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
      Left            =   4635
      TabIndex        =   2
      Top             =   1935
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continúa"
      Default         =   -1  'True
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
      Left            =   3375
      TabIndex        =   1
      Top             =   1935
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FCE7C7&
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
      Left            =   360
      TabIndex        =   0
      Top             =   405
      Width           =   5460
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Verifique que esté todo en condiciones para grabar el RECIBO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   225
         TabIndex        =   3
         Top             =   360
         Width           =   5055
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   2625
      Left            =   1500
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4260
      Visible         =   0   'False
      Width           =   6690
      _cx             =   11800
      _cy             =   4630
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
      BackColor       =   16710366
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12582912
      ForeColorSel    =   16711136
      BackColorBkg    =   12632256
      BackColorAlternate=   16574407
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"EmiRec.frx":13D6
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
      Left            =   150
      Top             =   3240
      Author          =   ""
      Creator         =   ""
      Title           =   ""
      Subject         =   ""
      Keywords        =   ""
      Compress        =   3
   End
End
Attribute VB_Name = "EmiRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NoAnduvo As Boolean, Mensaje As String
Dim dbCLTE As New ADODB.Connection, RgCLTE As New ADODB.Recordset
Dim RgCCTE As New ADODB.Recordset, RgFACT As New ADODB.Recordset
Dim RgCYB As New ADODB.Recordset

Private Const laselec20 = "SELECT * FROM Clientes WHERE "

Private Sub Form_Load()

Me.Move 2600, 1800, 6350, 3000

dbCLTE.ConnectionString = BDatos1
dbCLTE.Open

'miSQL1 = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE
'RgCLTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

'If RgCLTE.EOF Then
'    NoAnduvo = True
'    RgCLTE.Close
'    Exit Sub
'    End Ifa

End Sub

Private Sub Form_Unload(Cancel As Integer)

dbCLTE.Close

End Sub

'--------------------------------------------------------------------
'
'                    I M P R E S I O N
'
'--------------------------------------------------------------------
Private Sub Command1_Click()

'IMPRIME
IMPRIME2en1

Graba

End Sub

Private Sub Command2_Click()

Unload Me

End Sub
'Private Sub Command3_Click()
'fcmenu.Toolbar3.Buttons(3).Value = tbrUnpressed
'fcmenu.Toolbar3.Buttons(1).Value = tbrPressed
'Unload Me
'End Sub

Sub IMPRIME()

Dim I1%, i2%, i3%, LaLetra As String, CodAntic As Integer, OtraVez As Boolean
Dim s$, h$, d$, ICpbte As String * 7, Total1 As Currency, Total2 As Currency
Dim IFec1 As String * 11, ITipo As String * 6, IImpte1 As String * 12, IDto As String * 12
Dim IBco As String * 20, ICheq As String * 12, IImpte2 As String * 12, IFec2 As String * 11

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
    
 IMPRE.VP1.DrawPicture Picture1, "8mm", "5mm"
Escribe "ARIAL BLACK", 18, 0, 0, 14.6, 66, "S.R.L."
 
IMPRE.Caption = " Impresión del Recibo : " & LaLetra & " Nº " & "0001-" & _
                    Format(NroRec, "00000000") & " (Cta.Cte.)"

'Primer Rectángulo
DibujaRect 5, 7, 130, 200, 30, 30
DibujaLine 31, 7, 31, 200
DibujaLine 46, 7, 46, 200
DibujaLine 53, 7, 53, 200
DibujaLine 105, 7, 105, 200
DibujaLine 46, 95, 105, 95

'  Línea 1
'Escribe "ARIAL", 16, 1, 0, 6, 8, "ALESTEL S.R.L."
Escribe "ARIAL BLACK", 24, 1, 0, 6, 104, LaLetra
Escribe "", 10, 0, 1, 9, 132, "Buenos Aires, " & Format(Date, "ddd dd MMM YYYY")
'  Línea 2
Escribe "ARIAL", 9, 1, 0, 23, 14, "Av. Crovara 2948 - (1766) La Tablada (BA)"
Escribe "ARIAL", 7, 0, 0, 17, 99, "Documento no"
Escribe "", 16, 1, 1, 15, 120, "Recibo Oficial"
Escribe "", 12, 1, 1, 16, 160, "0001-" & Format(NroRec, "00000000")
'Escribe "", 9, 0, 0, 30, 153, Format(Time, "hh:mm")
'  Línea 3
'Escribe "ARIAL", 9, 1, 0, 26, 20, "(1235) La Tablada"
Escribe "ARIAL", 7, 0, 0, 20, 96, "válido como Factura"
Escribe "", 8, 0, 0, 23, 130, "Imp.Int.: No Resp.     I.B.: 901 33-70346790-9"
'  Línea 4
Escribe "ARIAL", 9, 1, 0, 27, 14, "Tel. (011) 4652-1040 / 2684 / 2689 / 3080"
Escribe "", 8, 0, 0, 27, 130, "Part.Munic.: 1985   CUIT: 33-70346790-9"
'  Línea 5
Escribe "", 10, 1, 1, 32, 8, "Sr./es: (" & CodCLTE & ") " & ClteNomb

'  Línea 6
Escribe "", 10, 0, 0, 37, 8, ClteDir
'  Línea 7
Escribe "", 10, 0, 0, 42, 8, "(" & ClteCPos & ") " & ClteLoc & " " & CltePCIA
Escribe "", 10, 1, 1, 42, 100, "IVA: " & CabRec.Combo1.List(ClteCIVA - 1) & _
                              "         CUIT : " & ClteCUIT

'  Línea 8   -  Títulos del Detalle
Escribe "", 8, 1, 1, 46, 32, "Comprobantes Cancelados"
Escribe "", 8, 1, 1, 46, 135, "Valores Recibidos"
'  Línea 9
Escribe "", 7, 0, 1, 49, 11, "Tipo             Nro.                 Fecha                        " & _
                            "   Importe                     Dto."
Escribe "", 7, 0, 1, 49, 102, "Nro./Tipo                        Bco./Detalle" & _
  "                                        Fecha                    Importe"

'----------------------------------------------
'           Detalle
'----------------------------------------------

' ---- Facturas ----

IMPRE.VP1.FontName = "Lucida Console"
IMPRE.VP1.FontItalic = False
IMPRE.VP1.CurrentX = "0mm"
IMPRE.VP1.FontSize = 8
IMPRE.VP1.TableBorder = tbNone

s$ = "~<+12mm|~<+15mm|~<+24mm|~>+20mm|~>+20mm"
h$ = "Tipo|Nro.|Fecha|Aplicado|Importe"

IMPRE.VP1.CurrentY = "54mm"
IMPRE.VP1.CurrentX = 0
IMPRE.VP1.TextAlign = 0
IMPRE.VP1.MarginLeft = "8mm"

Total1 = 0

For i2 = 1 To DetRec.FG1.Rows - 1

'I3 = DetRec.FG1.Row(I1)
If CCur(DetRec.FG1.TextMatrix(i2, 3)) = 0 Then GoTo Saltando1
I1 = I1 + 1

ITipo = DetRec.FG1.TextMatrix(i2, 1)
ICpbte = Val(DetRec.FG1.TextMatrix(i2, 0))
IFec1 = Format(DetRec.FG1.TextMatrix(i2, 2), "dd mmm yyyy")
IImpte1 = Format(DetRec.FG1.TextMatrix(i2, 3), "$ ###,##0.00")
IDto = ""
'IDto = Format(DetRec.FG1.TextMatrix(i2, 4), "$ ###,##0.00")

d$ = ITipo & "|" & ICpbte & "|" & IFec1 & "|" & IImpte1 & "|" & IDto

IMPRE.VP1.FontSize = 8
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 1
IMPRE.VP1.AddTable s$, h$, " || | |", , , True

Total1 = Total1 + DetRec.FG1.TextMatrix(i2, 3)

Saltando1:

Next i2

' ---- Cheques ----

PasaGrilla DetPago.FG1

IMPRE.VP1.FontName = "Lucida Console"
IMPRE.VP1.FontItalic = False
IMPRE.VP1.CurrentX = "0mm"
IMPRE.VP1.FontSize = 8
IMPRE.VP1.TableBorder = tbNone

s$ = "~<+24mm|~<+37mm|~<+23mm|~>+18mm"
h$ = "Nro./Cpbte.|Bco./Det.|Fecha|Importe"

IMPRE.VP1.CurrentY = "54mm"
IMPRE.VP1.CurrentX = 0
IMPRE.VP1.TextAlign = 0
IMPRE.VP1.MarginLeft = "97mm"

Total2 = 0

For i2 = 1 To DetPago.FG1.Rows - 1

'I3 = DetRec.FG1.Row(I1)
If Val(DetPago.FG1.TextMatrix(i2, 3)) = 0 Then GoTo Saltando2
I1 = I1 + 1

ICheq = DetPago.FG1.TextMatrix(i2, 0)
IBco = Mid(DetPago.FG1.TextMatrix(i2, 1), 1, 20)
IFec2 = Format(DetPago.FG1.TextMatrix(i2, 2), "dd mmm yyyy")
IImpte2 = Format(DetPago.FG1.TextMatrix(i2, 3), "$ ###,##0.00")

d$ = ICheq & "|" & IBco & "|" & IFec2 & "|" & IImpte2

IMPRE.VP1.FontSize = 8
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 1
IMPRE.VP1.AddTable s$, h$, " || | |", , , True

Total2 = Total2 + DetPago.FG1.TextMatrix(i2, 3)

Saltando2:

Next i2

' ---------------------------
'      Efectivo
' ---------------------------
If Val(DetPago.Text1.Text) > 0 Then
        ICheq = "* Efectivo *"
        IBco = ""
        IFec2 = Format(Date, "dd mmm yyyy")
        IImpte2 = Format(CCur(DetPago.Text1.Text), "$ ###,##0.00")
        
        d$ = ICheq & "|" & IBco & "|" & IFec2 & "|" & IImpte2
        
        IMPRE.VP1.FontSize = 8
        IMPRE.VP1.AddTable s$, h$, d$, , , True
        IMPRE.VP1.FontSize = 1
        IMPRE.VP1.AddTable s$, h$, " || | |", , , True
        
        Total2 = Total2 + CCur(DetPago.Text1.Text)
End If

'IMPRE.VP1.TableCell(37, 0, 0, i1, 4) = 50

Escribe "", 9, 0, 0, 101, 35, " Total Cancelado: " & Format(Total1, "$ ###,##0.00")
Escribe "", 9, 0, 0, 101, 160, " Total Valores: " & Format(Total2, "$ ###,##0.00")

Escribe "", 10, 1, 1, 107, 8, " Total: " & Format(Total1, "$ ###,##0.00")
Escribe "", 10, 1, 1, 107, 50, " Son Pesos:" & MontoEscrito(Total1)

Escribe "", 8, 0, 1, 120, 150, "___________________"
Escribe "", 8, 0, 1, 125, 152, "  Por Alestel SRL"

Escribe "", 8, 0, 0, 125, 8, "Este pago no implica cancelación de intereses por pago de facturas fuera de término"

If Not OtraVez Then
    Escribe "", 8, 1, 1, 132, 160, "Original"
Else
    Escribe "", 8, 1, 1, 132, 160, "Duplicado"
End If
 
 If Not OtraVez Then                     '  Esto estaba cuando salian 2 hojas
    IMPRE.VP1.NewPage
    OtraVez = True
    GoTo OtraVez
 End If

 IMPRE.VP1.EndDoc
 
 IMPRE.Show 1
 
OtraVez = False

End Sub

Sub IMPRIME2en1()

Dim I1%, i2%, i3%, LaLetra As String, CodAntic As Integer, OtraVez As Boolean
Dim s$, h$, d$, ICpbte As String * 7, Total1 As Currency, Total2 As Currency
Dim IFec1 As String * 11, ITipo As String * 6, IImpte1 As String * 16, IDto As String * 16
Dim IBco As String * 20, ICheq As String * 12, IImpte2 As String * 16, IFec2 As String * 11
Dim MasLin As Integer, ITCh As String * 3

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
 
IMPRE.Caption = " Impresión del Recibo : " & LaLetra & " Nº " & "0001-" & _
                    Format(NroRec, "00000000") & " (Cta.Cte.)"

'Primer Rectángulo
DibujaRect 5 + MasLin, 7, 130 + MasLin, 200, 30, 30
DibujaLine 31 + MasLin, 7, 31 + MasLin, 200
DibujaLine 46 + MasLin, 7, 46 + MasLin, 200
DibujaLine 53 + MasLin, 7, 53 + MasLin, 200
DibujaLine 105 + MasLin, 7, 105 + MasLin, 200
DibujaLine 46 + MasLin, 95, 105 + MasLin, 95

DibujaLine 145, 5, 145, 200

'  Línea 1
'Escribe "ARIAL", 16, 1, 0, 6, 8, "ALESTEL S.R.L."
Escribe "ARIAL BLACK", 24, 1, 0, 6 + MasLin, 104, LaLetra
Escribe "", 10, 0, 1, 9 + MasLin, 132, "Buenos Aires, " & Format(Date, "ddd dd MMM YYYY")
'  Línea 2
Escribe "ARIAL", 9, 1, 0, 23 + MasLin, 14, "Av. Crovara 2948 - (1766) La Tablada (BA)"
Escribe "ARIAL", 7, 0, 0, 17 + MasLin, 99, "Documento no"
Escribe "", 16, 1, 1, 15 + MasLin, 120, "Recibo Oficial"
Escribe "", 12, 1, 1, 16 + MasLin, 160, "0001-" & Format(NroRec, "00000000")
'Escribe "", 9, 0, 0, 30, 153, Format(Time, "hh:mm")
'  Línea 3
'Escribe "ARIAL", 9, 1, 0, 26, 20, "(1235) La Tablada"
Escribe "ARIAL", 7, 0, 0, 20 + MasLin, 96, "válido como Factura"
Escribe "", 8, 0, 0, 23 + MasLin, 130, "Imp.Int.: No Resp.     I.B.: 901 33-70346790-9"
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
Escribe "", 8, 1, 1, 46 + MasLin, 32, "Comprobantes Cancelados"
Escribe "", 8, 1, 1, 46 + MasLin, 135, "Valores Recibidos"
'  Línea 9
Escribe "", 7, 0, 1, 49 + MasLin, 11, "Tipo             Nro.                 Fecha                        " & _
                            "                        Importe    "
Escribe "", 7, 0, 1, 49 + MasLin, 102, "Nro./Tipo                        Bco./Detalle" & _
  "             Fecha                                              Importe"

'----------------------------------------------
'           Detalle
'----------------------------------------------

' ---- Facturas ----

IMPRE.VP1.FontName = "Lucida Console"
IMPRE.VP1.FontItalic = False
IMPRE.VP1.CurrentX = "0mm"
IMPRE.VP1.FontSize = 7
IMPRE.VP1.TableBorder = tbNone

s$ = "~<+12mm|~<+15mm|~<+24mm|~>+35mm|~>+20mm"
h$ = "Tipo|Nro.|Fecha|Aplicado|Importe"

IMPRE.VP1.CurrentY = (54 + MasLin) & "mm"
IMPRE.VP1.CurrentX = 0
IMPRE.VP1.TextAlign = 0
IMPRE.VP1.MarginLeft = "8mm"

Total1 = 0

For i2 = 1 To DetRec.FG1.Rows - 1

'I3 = DetRec.FG1.Row(I1)
If CCur(DetRec.FG1.TextMatrix(i2, 3)) = 0 Then GoTo Saltando1
I1 = I1 + 1

ITipo = DetRec.FG1.TextMatrix(i2, 1)
ICpbte = Val(DetRec.FG1.TextMatrix(i2, 0))
IFec1 = Format(DetRec.FG1.TextMatrix(i2, 2), "dd mmm yyyy")
RSet IImpte1 = Format(DetRec.FG1.TextMatrix(i2, 3), "$ ###,###,##0.00")
IDto = ""
'IDto = Format(DetRec.FG1.TextMatrix(i2, 4), "$ ###,##0.00")

d$ = ITipo & "|" & ICpbte & "|" & IFec1 & "|" & IImpte1 & "|"

IMPRE.VP1.FontSize = 7
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 1
IMPRE.VP1.AddTable s$, h$, " || | |", , , True

Total1 = Total1 + DetRec.FG1.TextMatrix(i2, 3)

Saltando1:

Next i2

' ---- Cheques ----

PasaGrilla DetPago.FG1

IMPRE.VP1.FontName = "Lucida Console"
IMPRE.VP1.FontItalic = False
IMPRE.VP1.CurrentX = "0mm"
IMPRE.VP1.FontSize = 7
IMPRE.VP1.TableBorder = tbNone

s$ = "~<+7mm|~<+20mm|~<+26mm|~<+20mm|~>+32mm"
h$ = "Tipo Ch.|Nro./Cpbte.|Bco./Det.|Fecha|Importe"

IMPRE.VP1.CurrentY = (54 + MasLin) & "mm"
IMPRE.VP1.CurrentX = 0
IMPRE.VP1.TextAlign = 0
IMPRE.VP1.MarginLeft = "95mm"

Total2 = 0

For i2 = 1 To DetPago.FG1.Rows - 1

'I3 = DetRec.FG1.Row(I1)
If Val(DetPago.FG1.TextMatrix(i2, 3)) = 0 Then GoTo Saltando2
I1 = I1 + 1


ITCh = ""
If DetPago.FG1.TextMatrix(i2, 12) = 1 Then
     ITCh = "ECh"
End If

ICheq = DetPago.FG1.TextMatrix(i2, 0)
IBco = Mid(DetPago.FG1.TextMatrix(i2, 1), 1, 20)
IFec2 = Format(DetPago.FG1.TextMatrix(i2, 2), "dd mmm yyyy")
RSet IImpte2 = Format(DetPago.FG1.TextMatrix(i2, 3), "$ ###,###,##0.00")


d$ = ITCh & "|" & ICheq & "|" & IBco & "|" & IFec2 & "|" & IImpte2

IMPRE.VP1.FontSize = 7
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 5
d$ = "  " & "|" & "(" & DetPago.FG1.TextMatrix(i2, 7) & "|" & _
        DetPago.FG1.TextMatrix(i2, 8) & "|" & DetPago.FG1.TextMatrix(i2, 9) & ")" & _
        "|" & ""
IMPRE.VP1.AddTable s$, h$, d$, , , True

IMPRE.VP1.FontSize = 1
IMPRE.VP1.AddTable s$, h$, " ||| | |", , , True

Total2 = Total2 + DetPago.FG1.TextMatrix(i2, 3)

Saltando2:

Next i2

' ---------------------------
'      Efectivo
' ---------------------------
If Val(DetPago.Text1.Text) > 0 Then
        ICheq = "* Efectivo *"
        IBco = ""
        IFec2 = Format(Date, "dd mmm yyyy")
        RSet IImpte2 = Format(CCur(DetPago.Text1.Text), "$ ###,###,##0.00")
        
        d$ = " |" & ICheq & "|" & IBco & "|" & IFec2 & "|" & IImpte2
        
        IMPRE.VP1.FontSize = 7
        IMPRE.VP1.AddTable s$, h$, d$, , , True
        IMPRE.VP1.FontSize = 1
        IMPRE.VP1.AddTable s$, h$, " || | |", , , True
        
        Total2 = Total2 + CCur(DetPago.Text1.Text)
End If

'IMPRE.VP1.TableCell(37, 0, 0, i1, 4) = 50

Escribe "", 8, 0, 0, 101 + MasLin, 45, " Total Cancelado: " & Format(Total1, "$ ###,###,##0.00")
Escribe "", 8, 0, 0, 101 + MasLin, 155, " Total Valores: " & Format(Total2, "$ ###,###,##0.00")

Escribe "", 9, 1, 1, 107 + MasLin, 8, " Total: " & Format(Total1, "$ ###,###,##0.00")
Escribe "", 9, 1, 1, 107 + MasLin, 50, " Son Pesos:" & MontoEscrito(Total1)

Escribe "", 8, 0, 1, 120 + MasLin, 150, "___________________"
Escribe "", 8, 0, 1, 125 + MasLin, 152, "  Por Alestel SRL"

Escribe "", 8, 0, 0, 125 + MasLin, 8, "Este pago no implica cancelación de intereses por pago de facturas fuera de término"

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
 
 'IMPRE.Show 1

'Genera PDF
VSPDF81.Title = "Recibo-" & LaLetra & "-0001-" & NroRec
VSPDF81.ConvertDocument IMPRE.VP1, "c:\Factu\Comprob\" & _
                         Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "\" & _
                         VSPDF81.Title & ".pdf"

EsProvisorio = True

'Muestra PDF
PDF.PDF1.LoadFile "c:\Factu\Comprob\" & _
                         Year(Date) & "-" & Format(Month(Date), "00") & "-" & Format(Day(Date), "00") & "\" & _
                         VSPDF81.Title & ".pdf"
PDF.Caption = VSPDF81.Title
PDF.Show
PDF.PDF1.setShowToolbar True

     OtraVez = False
    MasLin = 0

End Sub


Sub IMPRIME2en1_Vie()

Dim I1%, i2%, i3%, LaLetra As String, CodAntic As Integer, OtraVez As Boolean
Dim s$, h$, d$, ICpbte As String * 7, Total1 As Currency, Total2 As Currency
Dim IFec1 As String * 11, ITipo As String * 6, IImpte1 As String * 12, IDto As String * 12
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
 
IMPRE.Caption = " Impresión del Recibo : " & LaLetra & " Nº " & "0001-" & _
                    Format(NroRec, "00000000") & " (Cta.Cte.)"

'Primer Rectángulo
DibujaRect 5 + MasLin, 7, 130 + MasLin, 200, 30, 30
DibujaLine 31 + MasLin, 7, 31 + MasLin, 200
DibujaLine 46 + MasLin, 7, 46 + MasLin, 200
DibujaLine 53 + MasLin, 7, 53 + MasLin, 200
DibujaLine 105 + MasLin, 7, 105 + MasLin, 200
DibujaLine 46 + MasLin, 95, 105 + MasLin, 95

DibujaLine 145, 5, 145, 200

'  Línea 1
'Escribe "ARIAL", 16, 1, 0, 6, 8, "ALESTEL S.R.L."
Escribe "ARIAL BLACK", 24, 1, 0, 6 + MasLin, 104, LaLetra
Escribe "", 10, 0, 1, 9 + MasLin, 132, "Buenos Aires, " & Format(Date, "ddd dd MMM YYYY")
'  Línea 2
Escribe "ARIAL", 9, 1, 0, 23 + MasLin, 14, "Av. Crovara 2948 - (1766) La Tablada (BA)"
Escribe "ARIAL", 7, 0, 0, 17 + MasLin, 99, "Documento no"
Escribe "", 16, 1, 1, 15 + MasLin, 120, "Recibo Oficial"
Escribe "", 12, 1, 1, 16 + MasLin, 160, "0001-" & Format(NroRec, "00000000")
'Escribe "", 9, 0, 0, 30, 153, Format(Time, "hh:mm")
'  Línea 3
'Escribe "ARIAL", 9, 1, 0, 26, 20, "(1235) La Tablada"
Escribe "ARIAL", 7, 0, 0, 20 + MasLin, 96, "válido como Factura"
Escribe "", 8, 0, 0, 23 + MasLin, 130, "Imp.Int.: No Resp.     I.B.: 901 33-70346790-9"
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
Escribe "", 8, 1, 1, 46 + MasLin, 32, "Comprobantes Cancelados"
Escribe "", 8, 1, 1, 46 + MasLin, 135, "Valores Recibidos"
'  Línea 9
Escribe "", 7, 0, 1, 49 + MasLin, 11, "Tipo             Nro.                 Fecha                        " & _
                            "   Importe                     Dto."
Escribe "", 7, 0, 1, 49 + MasLin, 102, "Nro./Tipo                        Bco./Detalle" & _
  "                                        Fecha                    Importe"

'----------------------------------------------
'           Detalle
'----------------------------------------------

' ---- Facturas ----

IMPRE.VP1.FontName = "Lucida Console"
IMPRE.VP1.FontItalic = False
IMPRE.VP1.CurrentX = "0mm"
IMPRE.VP1.FontSize = 8
IMPRE.VP1.TableBorder = tbNone

s$ = "~<+12mm|~<+15mm|~<+24mm|~>+20mm|~>+20mm"
h$ = "Tipo|Nro.|Fecha|Aplicado|Importe"

IMPRE.VP1.CurrentY = (54 + MasLin) & "mm"
IMPRE.VP1.CurrentX = 0
IMPRE.VP1.TextAlign = 0
IMPRE.VP1.MarginLeft = "8mm"

Total1 = 0

For i2 = 1 To DetRec.FG1.Rows - 1

'I3 = DetRec.FG1.Row(I1)
If CCur(DetRec.FG1.TextMatrix(i2, 3)) = 0 Then GoTo Saltando1
I1 = I1 + 1

ITipo = DetRec.FG1.TextMatrix(i2, 1)
ICpbte = Val(DetRec.FG1.TextMatrix(i2, 0))
IFec1 = Format(DetRec.FG1.TextMatrix(i2, 2), "dd mmm yyyy")
IImpte1 = Format(DetRec.FG1.TextMatrix(i2, 3), "$ ###,##0.00")
IDto = ""
'IDto = Format(DetRec.FG1.TextMatrix(i2, 4), "$ ###,##0.00")

d$ = ITipo & "|" & ICpbte & "|" & IFec1 & "|" & IImpte1 & "|" & IDto

IMPRE.VP1.FontSize = 8
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 1
IMPRE.VP1.AddTable s$, h$, " || | |", , , True

Total1 = Total1 + DetRec.FG1.TextMatrix(i2, 3)

Saltando1:

Next i2

' ---- Cheques ----

PasaGrilla DetPago.FG1

IMPRE.VP1.FontName = "Lucida Console"
IMPRE.VP1.FontItalic = False
IMPRE.VP1.CurrentX = "0mm"
IMPRE.VP1.FontSize = 8
IMPRE.VP1.TableBorder = tbNone

s$ = "~<+24mm|~<+37mm|~<+23mm|~>+18mm"
h$ = "Nro./Cpbte.|Bco./Det.|Fecha|Importe"

IMPRE.VP1.CurrentY = (54 + MasLin) & "mm"
IMPRE.VP1.CurrentX = 0
IMPRE.VP1.TextAlign = 0
IMPRE.VP1.MarginLeft = "97mm"

Total2 = 0

For i2 = 1 To DetPago.FG1.Rows - 1

'I3 = DetRec.FG1.Row(I1)
If Val(DetPago.FG1.TextMatrix(i2, 3)) = 0 Then GoTo Saltando2
I1 = I1 + 1

ICheq = DetPago.FG1.TextMatrix(i2, 0)
IBco = Mid(DetPago.FG1.TextMatrix(i2, 1), 1, 20)
IFec2 = Format(DetPago.FG1.TextMatrix(i2, 2), "dd mmm yyyy")
IImpte2 = Format(DetPago.FG1.TextMatrix(i2, 3), "$ ###,##0.00")

d$ = ICheq & "|" & IBco & "|" & IFec2 & "|" & IImpte2

IMPRE.VP1.FontSize = 8
IMPRE.VP1.AddTable s$, h$, d$, , , True
IMPRE.VP1.FontSize = 1
IMPRE.VP1.AddTable s$, h$, " || | |", , , True

Total2 = Total2 + DetPago.FG1.TextMatrix(i2, 3)

Saltando2:

Next i2

' ---------------------------
'      Efectivo
' ---------------------------
If Val(DetPago.Text1.Text) > 0 Then
        ICheq = "* Efectivo *"
        IBco = ""
        IFec2 = Format(Date, "dd mmm yyyy")
        IImpte2 = Format(CCur(DetPago.Text1.Text), "$ ###,##0.00")
        
        d$ = ICheq & "|" & IBco & "|" & IFec2 & "|" & IImpte2
        
        IMPRE.VP1.FontSize = 8
        IMPRE.VP1.AddTable s$, h$, d$, , , True
        IMPRE.VP1.FontSize = 1
        IMPRE.VP1.AddTable s$, h$, " || | |", , , True
        
        Total2 = Total2 + CCur(DetPago.Text1.Text)
End If

'IMPRE.VP1.TableCell(37, 0, 0, i1, 4) = 50

Escribe "", 9, 0, 0, 101 + MasLin, 35, " Total Cancelado: " & Format(Total1, "$ ###,##0.00")
Escribe "", 9, 0, 0, 101 + MasLin, 160, " Total Valores: " & Format(Total2, "$ ###,##0.00")

Escribe "", 10, 1, 1, 107 + MasLin, 8, " Total: " & Format(Total1, "$ ###,##0.00")
Escribe "", 10, 1, 1, 107 + MasLin, 50, " Son Pesos:" & MontoEscrito(Total1)

Escribe "", 8, 0, 1, 120 + MasLin, 150, "___________________"
Escribe "", 8, 0, 1, 125 + MasLin, 152, "  Por Alestel SRL"

Escribe "", 8, 0, 0, 125 + MasLin, 8, "Este pago no implica cancelación de intereses por pago de facturas fuera de término"

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
 
 IMPRE.Show 1
 
     OtraVez = False
    MasLin = 0


End Sub

Sub Graba()
Dim I1%, i2%, item As Single, LaLetra As String, LaFact As Double, LaFech As Date
Dim ElTipo As String, Corr As Integer, LaFecVt As Date
     
  vbMsgBoxTitle = "Grabación del Recibo"
  vbMsgBoxText = "El siguiente proceso registrará el Recibo  " & vbCrLf & "Desea continuar ?  "
  vbMsgBoxResp = vbYesNoCancel + vbExclamation + vbApplicationModal + vbDefaultButton1
 
  vbResponse = MsgBox(vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle)
 
  Select Case vbResponse
         Case vbYes
              GoTo Grabacion
 
         Case vbNo
              GoTo Salir1
 
         Case vbCancel
              GoTo Salir2
  End Select
 
Grabacion:

dbCLTE.BeginTrans

'  Graba Número de Recibo

    miSQL1 = "SELECT top 1 * FROM Parametro WHERE CLAVE = '1'"
    RgTABL.Open miSQL1, dbCLTE, adOpenDynamic, adLockOptimistic
    If RgTABL.EOF Then
        NoAnduvo = True
        RgTABL.Close
        Exit Sub
    End If
            
            RgTABL!nume8 = NroRec

        RgTABL.Update
        RgTABL.Close

'  Lee y actualiza Correlativo en Clientes

miSQL1 = "SELECT * FROM Clientes WHERE CODIGO = " & CodCLTE
RgCLTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If RgCLTE.EOF Then
    NoAnduvo = True
    RgCLTE.Close
    Exit Sub
    End If

Corr = Val(RgCLTE!corr1) + 1
If Corr = 0 Or Corr > 99 Then Corr = 1

For I1 = 1 To DetRec.FG1.Rows - 1

If DetRec.FG1.TextMatrix(I1, 3) = 0 Then GoTo Salta2

LaFact = Val(DetRec.FG1.TextMatrix(I1, 0))
LaFech = DetRec.FG1.TextMatrix(I1, 2)
ElTipo = Val(DetRec.FG1.TextMatrix(I1, 7))
'RgCLTE.Close

Graba_1:
    
'   Cuentas Corrientes
laLOGIK1 = Month(LaFech) & "/" & Day(LaFech) & "/" & Year(LaFech)

miSQL1 = "SELECT * FROM CtasCtes WHERE CLTE= " & CodCLTE & " AND CPBTE= " & LaFact & _
                                " AND FECHA= " & Separa & laLOGIK1 & Separa & " AND TIPO= " & ElTipo
RgCCTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If RgCCTE.EOF Then
    RgCCTE.Close
    GoTo Salta2
    End If
   
LaFecVt = RgCCTE!FECVTO

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
    
If ElTipo = 2 Or ElTipo = 5 Or ElTipo = 8 Then
    RgCCTE!DEBE = RgCCTE!DEBE + DetRec.FG1.TextMatrix(I1, 3)
Else
    RgCCTE!DEBE = RgCCTE!DEBE - DetRec.FG1.TextMatrix(I1, 3)
End If
    
RgCCTE.Update
RgCCTE.Close

'      Graba la Imputación

laLOGIK1 = Month(LaFech) & "/" & Day(LaFech) & "/" & Year(LaFech)
miSQL1 = "SELECT * FROM Imputacion WHERE CLTE = " & CodCLTE & " AND CPBTE = " & LaFact & _
                                " AND FECHA = " & Separa & laLOGIK1 & Separa & " AND TIPO = '" & ElTipo & "'"
RgCCTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If Not RgCCTE.EOF Then
    RgCCTE.Close
    GoTo Salta2
    End If

RgCCTE.AddNew

RgCCTE!Cpbte = NroRec
RgCCTE!FECHA = Date
RgCCTE!TIPO = 4
RgCCTE!CLTE = CodCLTE

RgCCTE!Corr = Corr
RgCCTE!fechai = LaFech
RgCCTE!cpbtei = LaFact
RgCCTE!tipoi = ElTipo
RgCCTE!fecvtoi = LaFecVt

RgCCTE!Impte = CCur(DetRec.FG1.TextMatrix(I1, 3))

RgCCTE!USUARIO = Mid(Red_Usuario, 1, 6)
    
    RgCCTE.Update
    
If DetRec.FG1.TextMatrix(I1, 4) = 0 Then GoTo SaltaDesc

RgCCTE.AddNew

RgCCTE!Cpbte = NroRec
RgCCTE!FECHA = Date
RgCCTE!TIPO = 6
RgCCTE!CLTE = CodCLTE

RgCCTE!Corr = Corr
RgCCTE!fechai = LaFech
RgCCTE!cpbtei = LaFact
RgCCTE!tipoi = ElTipo
RgCCTE!fecvtoi = LaFecVt

RgCCTE!Impte = CCur(DetRec.FG1.TextMatrix(I1, 4))

RgCCTE!USUARIO = Mid(Red_Usuario, 1, 6)
    
    RgCCTE.Update
    
SaltaDesc:

    RgCCTE.Close
    
Salta2:

Next I1

'-----------------------------------------------
'                Graba el Recibo
'-----------------------------------------------

laLOGIK1 = Month(Date) & "/" & Day(Date) & "/" & Year(Date)
miSQL1 = "SELECT * FROM CtasCtes WHERE CLTE = " & CodCLTE & " AND CPBTE= " & NroRec & _
                                " AND FECHA = " & Separa & laLOGIK1 & Separa & " AND TIPO= " & 4
RgCCTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If Not RgCCTE.EOF Then
    RgCCTE.Close
    GoTo aCltes
    End If

RgCCTE.AddNew

RgCCTE!prefijo = 1
RgCCTE!Cpbte = NroRec
RgCCTE!FECHA = Date
RgCCTE!TIPO = 4
RgCCTE!CLTE = CodCLTE

If Corr > 99 Then Corr = 0

RgCCTE!imput1 = Corr
RgCCTE!imput2 = 0
RgCCTE!imput3 = 0
RgCCTE!imput4 = 0
RgCCTE!imput5 = 0
RgCCTE!imput6 = 0

RgCCTE!DEBE = CCur(PieRec.Label303)
RgCCTE!Impte = CCur(PieRec.Label303)

RgCCTE!USUAR = Mid(Red_Usuario, 1, 6)
RgCCTE!moti = 0
RgCCTE!tipo9 = 0
RgCCTE!CVTA = 0
RgCCTE!bon = 0
RgCCTE!FECVTO = Date
RgCCTE!Letra = "X"
    
    RgCCTE.Update

'-----------------------------------------------
'                Graba el Descuento
'-----------------------------------------------
If PieRec.Label305.Caption = 0 Then GoTo aCltes

RgCCTE.AddNew

RgCCTE!prefijo = 1
RgCCTE!Cpbte = NroRec
RgCCTE!FECHA = Date
RgCCTE!TIPO = 6
RgCCTE!CLTE = CodCLTE

RgCCTE!imput1 = Corr
RgCCTE!imput2 = 0
RgCCTE!imput3 = 0
RgCCTE!imput4 = 0
RgCCTE!imput5 = 0
RgCCTE!imput6 = 0

RgCCTE!DEBE = CCur(PieRec.Label305)
RgCCTE!Impte = CCur(PieRec.Label305)

RgCCTE!USUAR = Mid(Red_Usuario, 1, 6)
RgCCTE!moti = 0
RgCCTE!tipo9 = 0
RgCCTE!CVTA = 0
RgCCTE!bon = 0
RgCCTE!FECVTO = Date
RgCCTE!Letra = "X"
    
    RgCCTE.Update
    
aCltes:
    
    RgCCTE.Close
    
'
'   Actualiza el Mtro. de Clientes
'
    RgCLTE!corr1 = Corr
    RgCLTE!DEUDA = RgCLTE!DEUDA - CCur(PieRec.Label303.Caption)
    

RgCLTE!USUARIO = Mid(Red_Usuario, 1, 6)
RgCLTE!FACTUAL = Date
    
    RgCLTE.Update

'----------------------------------
'   Actualiza el Arch. de Cheques
'----------------------------------


For I1 = 1 To DetPago.FG1.Rows - 1

If Val(DetPago.FG1.TextMatrix(I1, 10)) > 0 Then GoTo aAnticipos

If Val(DetPago.FG1.TextMatrix(I1, 3)) = 0 Then GoTo AlNextI1

miSQL1 = "SELECT * FROM Cheques WHERE NROCHEQ = " & Val(DetPago.FG1.TextMatrix(I1, 0))
RgCYB.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If Not RgCYB.EOF Then
    RgCYB.Close
    GoTo aAnticipos
    End If

RgCYB.AddNew

RgCYB!NroCheq = Val(DetPago.FG1.TextMatrix(I1, 0))
RgCYB!BcoSuc = DetPago.FG1.TextMatrix(I1, 4)
RgCYB!cpBCO = ""
RgCYB!datosad = DetPago.FG1.TextMatrix(I1, 7)
RgCYB!FECVTO = DetPago.FG1.TextMatrix(I1, 2)
RgCYB!fecemi = DetPago.FG1.TextMatrix(I1, 5)
RgCYB!CLTE = CodCLTE
RgCYB!tiping = 1                                ' 1-Recibo CC
RgCYB!cpbing = NroRec
RgCYB!concep = DetPago.FG1.TextMatrix(I1, 8)
RgCYB!observ = DetPago.FG1.TextMatrix(I1, 9)
If DetPago.FG1.TextMatrix(I1, 6) Then
    RgCYB!orden = 1
Else
    RgCYB!orden = 0
End If
RgCYB!propio = 0                                ' 0-no
RgCYB!Estado = 1                                ' 1-En Cartera
RgCYB!nroopE = 0

RgCYB!Importe = CCur(DetPago.FG1.TextMatrix(I1, 3))

RgCYB!tipegr = 0
RgCYB!cpbegr = 0
RgCYB!fecegr = "01/01/1900"
RgCYB!fecrech = "01/01/1900"
RgCYB!MotRech = 0
RgCYB!destino = 0
RgCYB!moneda = 0
RgCYB!cotiz = 0
RgCYB!NROOPS = 0


RgCYB!FECING = Date
RgCYB!USUARIO = Mid(Red_Usuario, 1, 6)
    
    RgCYB.Update
    RgCYB.Close
    
    GoTo AlNextI1
    
aAnticipos:

miSQL1 = "SELECT * FROM MovimVS WHERE CPBTE = " & NroRec & " AND TIPREG = '" & DetPago.FG1.TextMatrix(I1, 10) & "'"
RgCCTE.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

If Not RgCCTE.EOF Then
    RgCCTE.Close
    GoTo AlNextI1
    End If

RgCCTE.AddNew

RgCCTE!CLTE = CodCLTE
RgCCTE!FECHA = Now
RgCCTE!TIPO = 4                                ' 1-Recibo CC
RgCCTE!Cpbte = NroRec
RgCCTE!TIPREG = DetPago.FG1.TextMatrix(I1, 10)
RgCCTE!Impte = DetPago.FG1.TextMatrix(I1, 3)
RgCCTE!BcoSuc = DetPago.FG1.TextMatrix(I1, 4)
RgCCTE!datosad = DetPago.FG1.TextMatrix(I1, 7)
RgCCTE!concep = DetPago.FG1.TextMatrix(I1, 8)
RgCCTE!observ = DetPago.FG1.TextMatrix(I1, 9)

RgCCTE.Update
RgCCTE.Close

AlNextI1:

Next I1
' ---------------------------
'      Efectivo
' ---------------------------
If Val(DetPago.Text1.Text) > 0 Then
    miSQL1 = "SELECT * FROM Efectivo WHERE Cpbte = " & NroRec & " AND Tipo = " & 4
    RgCYB.Open miSQL1, dbCLTE, adOpenDynamic, adLockPessimistic

    If RgCYB.EOF Then
'        RgCYB.Close
'        dbCLTE.RollbackTrans
'        GoTo Salir1
        RgCYB.AddNew
        End If
        
        
        RgCYB!FECHA = Now
        RgCYB!TIPO = 4
        RgCYB!Cpbte = NroRec
        RgCYB!Impte = CCur(DetPago.Text1.Text)
        
            RgCYB.Update
            RgCYB.Close

End If
'-----------------------------------------
    
'   Cliente  -   Aqui debería grabar el total de valores pendientes
        
'    rgCLTE!CORR1 = Corr
'    rgCLTE!SALDO = rgCLTE!SALDO - CCur(PieRec.Label303)
    
'    rgCLTE.Update
'---------------------------------------
        
dbCLTE.CommitTrans

Salir1:

Unload Me

CodCLTE = 0
FCMENU.ReiniciaRec

Salir2:

End Sub

Function PasaGrilla(Grilla As VSFlexGrid)
Dim I1

If FG1.Rows > 0 Then LimpiaGrilla

For I1 = 1 To Grilla.Rows - 1

If Val(Grilla.TextMatrix(I1, 10)) = 0 Then
    Grilla.TextMatrix(I1, 10) = 99
End If

If Val(Grilla.TextMatrix(I1, 3)) > 0 Then
    FG1.AddItem Grilla.TextMatrix(I1, 0) & vbTab & Grilla.TextMatrix(I1, 1) & vbTab & _
                Grilla.TextMatrix(I1, 2) & vbTab & Grilla.TextMatrix(I1, 3) & vbTab & _
                Grilla.TextMatrix(I1, 4) & vbTab & Grilla.TextMatrix(I1, 5) & vbTab & _
                Grilla.TextMatrix(I1, 6) & vbTab & Grilla.TextMatrix(I1, 7) & vbTab & _
                Grilla.TextMatrix(I1, 8) & vbTab & Grilla.TextMatrix(I1, 9) & vbTab & _
                Grilla.TextMatrix(I1, 10), FG1.Rows
End If

Next

If Val(DetPago.Text1.Text) > 0 Then
    FG1.AddItem "- Efectivo -" & vbTab & " " & vbTab & Date & vbTab & Val(DetPago.Text1.Text), FG1.Rows
End If

'FG1.FixedRows = 1

FG1.Col = 10
FG1.Sort = flexSortGenericAscending
' EmiRec.BorderStyle = 5: FG1.Visible = True

End Function

Sub LimpiaGrilla()
Dim I1%
   
FG1.Clear

For i = (FG1.Rows - 1) To 0 Step -1
FG1.RemoveItem (i)
Next i

End Sub
