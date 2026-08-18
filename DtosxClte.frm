VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form DtosxClte 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Descuentos del cliente"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10305
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
   ScaleHeight     =   4335
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Modificar"
      Height          =   345
      Left            =   300
      TabIndex        =   10
      Top             =   3090
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      Height          =   2655
      Left            =   4050
      ScaleHeight     =   2595
      ScaleWidth      =   5865
      TabIndex        =   5
      Top             =   780
      Visible         =   0   'False
      Width           =   5925
      Begin VB.CommandButton Command6 
         Caption         =   "Cerrar"
         Height          =   285
         Left            =   4710
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00F9FADC&
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   2250
         Width           =   2685
      End
      Begin VSFlex8Ctl.VSFlexGrid FG2 
         Height          =   1875
         Left            =   90
         TabIndex        =   8
         Top             =   360
         Width           =   5745
         _cx             =   10134
         _cy             =   3307
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Lucida Console"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"DtosxClte.frx":0000
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Selecciona Sección"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   60
         Width           =   5625
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grabar"
      Height          =   345
      Left            =   7230
      TabIndex        =   4
      Top             =   3090
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   345
      Left            =   8610
      TabIndex        =   3
      Top             =   3090
      Width           =   1275
   End
   Begin VSFlex8Ctl.VSFlexGrid FG1 
      Height          =   2175
      Left            =   150
      TabIndex        =   0
      Top             =   840
      Width           =   9855
      _cx             =   17383
      _cy             =   3836
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"DtosxClte.frx":0050
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
      ExplorerBar     =   0
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
   Begin VB.Label Label2 
      Height          =   285
      Left            =   180
      TabIndex        =   2
      Top             =   510
      Width           =   6225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Descuentos por Tipo de Artículo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2070
      TabIndex        =   1
      Top             =   60
      Width           =   6015
   End
End
Attribute VB_Name = "DtosxClte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LaRow As Integer, EsEnter As Boolean
Dim Descri As String, CodPaso As String, ALFAB As String

Public dbDTOS As New ADODB.Connection
Public RgTAB2 As New ADODB.Recordset

Private Sub Command3_Click()

FG1.Enabled = True
Command3.Visible = False

End Sub

Private Sub Form_Load()

Me.Move 0, 1610, 10400, 4000

LaRow = 1

Label2.Caption = "  " & CodCLTE & " - " & ClteNomb

CargaGrilla

End Sub

Private Sub Form_Activate()

If Not FG1.Enabled Then
    Command3.Visible = True
End If
          
If FG1.Rows <= 2 Then Exit Sub
If FG1.Row < 1 Then FG1.Row = 1

End Sub

Private Sub Form_Deactivate()

FG1.Enabled = False

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

    Grabar

End Sub

Private Sub Command6_Click()

Picture2.Visible = False

End Sub

'--------------------------------------------------------------
'           G R I L L A
'--------------------------------------------------------------

'------------------------------------------------------------
'           0-Sección         1-Descrip.         2-Dto.1
'           3-Dto.2           4-Dto.3            5-Dto.4
'           6-Dto.5           7-Rgo.1            8-Rgo.2
'           9-Rgo.3
'------------------------------------------------------------

Private Sub FG1_GotFocus()
            
    FG1.EditCell
    If FG1.EditWindow = 0 Then FG1.EditCell

End Sub

Private Sub FG1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Dim ShiftTest As Integer

If Col = 1 Then Exit Sub

LaTecla = KeyCode


'   ALT-DEL Borra el Renglón (si no es el único)
ShiftTest = Shift And 7
If (KeyCode = 46 And (ShiftTest = 4 Or ShiftTest = 6)) And FG1.Rows > 2 Then
    BorraRenglon (Row)
'    VerSiAgrega
    FG1.Select (FG1.Rows - 1), 0
    FG1.EditCell
    If FG1.EditWindow = 0 Then FG1.EditCell
    Exit Sub
End If

End Sub

Private Sub FG1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

' If Col = 3 Then Exit Sub

If KeyAscii = 13 Then
    EsEnter = True
End If

If KeyAscii = 8 Or KeyAscii = 13 Or KeyAscii = 9 Then Exit Sub

If Col = 1 Then Exit Sub
                
If Col > 1 Then
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
        KeyAscii = 0
    End If
End If

If Col > 1 Then
            If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
            If KeyAscii = Asc(",") Then Exit Sub
End If

End Sub

Private Sub FG1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim Control As String

If Col > 1 And FG1.TextMatrix(Row, 0) = "" Then
'------------------------------------------
'           Verifica si ya se ingresó el código
'
        Control = FG1.TextMatrix(Row, 0)
            
        If FG1.FindRow(Control, , 0) > -1 And FG1.FindRow(Control, , 0) <> Row Then
           FG1.Col = 0
           FG1.EditCell
           Exit Sub
        Else
            FG1.TextMatrix(Row, 0) = Control
        End If
'-----------------------------------------

End If

End Sub

Private Sub FG1_AfterEdit(ByVal Row As Long, ByVal Col As Long)

If Col > 1 Then
    If FG1.TextMatrix(Row, Col) = "" Then FG1.TextMatrix(Row, Col) = "0"
End If

If Not EsEnter Then Exit Sub

Select Case Col
 
    Case 0
'        FG1.TextMatrix(FG1.Row, 0) = 0
        FG1.Cell(6, Row, 0, Row, 6) = FG1.Cell(flexcpBackColor, Row, 9)
        
        Correa = FG1.TextMatrix(FG1.Row, 0)
        If Correa = "" Then
            EsEnter = False
            Text1.Text = ""
            BuscaEnTabla Correa
        Else
            Descri = LaDescri(FG1.TextMatrix(FG1.Row, 0))
            If Descri = "" Then
                MsgBox "La Sección " & Correa & " No Existe", vbExclamation, "Secciones"
                FG1.Col = 0
                Exit Sub
                
            Else
                FG1.TextMatrix(FG1.Row, 1) = Descri
            End If
        End If
            
                FG1.Col = 2
        
    Case 2
                
        If FG1.TextMatrix(Row, 2) = "" Then FG1.TextMatrix(Row, 2) = 0
        FG1.Col = 3
            
    Case 3
             
        If FG1.TextMatrix(Row, 3) = "" Then FG1.TextMatrix(Row, 3) = 0
        FG1.Col = 4
  
    Case 4
             
        If FG1.TextMatrix(Row, 4) = "" Then FG1.TextMatrix(Row, 4) = 0
        FG1.Col = 5
  
    Case 5
        
        If FG1.TextMatrix(Row, 5) = "" Then FG1.TextMatrix(Row, 5) = 0
        FG1.Col = 6
  
    Case 6
        
        If FG1.TextMatrix(Row, 6) = "" Then FG1.TextMatrix(Row, 6) = 0
        FG1.Col = 7
  
    Case 7
        
        If FG1.TextMatrix(Row, 7) = "" Then FG1.TextMatrix(Row, 7) = 0
        FG1.Col = 8
  
    Case 8
        
        If FG1.TextMatrix(Row, 8) = "" Then FG1.TextMatrix(Row, 8) = 0
        FG1.Col = 9
  
    Case 9
        
        If FG1.TextMatrix(Row, 9) = "" Then FG1.TextMatrix(Row, 9) = 0
'        FG1.Col = 9

End Select

If EsEnter Then
    EsEnter = False
    
    If Col < 9 Then
       FG1.EditCell
       If FG1.EditWindow = 0 Then FG1.EditCell
    Else
        VerSiAgrega
    End If

End If

End Sub

Private Sub FG1_Click()
            
If FG1.Col = 1 Then FG1.Col = 0

    FG1.EditCell
    If FG1.EditWindow = 0 Then FG1.EditCell

End Sub

Private Sub FG1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Salir

If FG1.MouseRow = -1 Then Exit Sub
If LaRow > (FG1.Rows - 1) Then LaRow = (FG1.Rows - 1)

If FG1.MouseCol > 0 Then
        FG1.ToolTipText = " " & FG1.TextMatrix(FG1.MouseRow, 1) & " "
Else
        FG1.ToolTipText = " " & "Blanco ó '0' para buscar en tabla" & " "
    End If

If FG1.MouseRow = FG1.Row Then
   FG1.Cell(7, LaRow, 0, LaRow, 5) = &H80000012
   FG1.BackColorSel = &HFF0000
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 5) = &HFCE7C7
   LaRow = FG1.MouseRow
   Else
   FG1.BackColorSel = &H800000
   End If

If FG1.MouseRow > 0 And FG1.MouseRow < FG1.Rows Then
   FG1.Cell(7, LaRow, 0, LaRow, 5) = &H80000012
   FG1.Cell(7, FG1.MouseRow, 0, FG1.MouseRow, 5) = &HFF0000
   LaRow = FG1.MouseRow
   End If

Salir:

End Sub

Private Sub FG2_Click()

        Correa = FG2.TextMatrix(FG2.Row, 0)
        Descri = FG2.TextMatrix(FG2.Row, 1)
        FG1.TextMatrix(FG1.Row, 0) = FG2.TextMatrix(FG2.Row, 0)
        FG1.TextMatrix(FG1.Row, 1) = FG2.TextMatrix(FG2.Row, 1)
    
Picture2.Visible = False
FG1.Enabled = True
FG1.SetFocus

End Sub

Private Sub text1_Change()

ALFAB = Text1.Text
BuscaEnTabla CodPaso

End Sub

Private Sub text1_GotFocus()

Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub text1_KeyPress(KeyAscii As Integer)

If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    KeyAscii = KeyAscii - 32
    End If

If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub

End Sub

Private Function BuscaEnTabla(CODIGO As String)
Dim I1%

CodPaso = CODIGO


Picture2.Visible = True
FG1.Enabled = False

Picture2.Move 4080, 810, 5900, 2700
FG2.Rows = 1
FG2.Row = 0

If ApuntaTabla(CODIGO) Then Exit Function

Do Until RgTABL.EOF

FG2.AddItem RgTABL!COD & vbTab & RgTABL!Descri, FG2.Row + 1

RgTABL.MoveNext

Loop
 
FG2.AutoSize 0, 1
Text1.SetFocus

RgTABL.Close

End Function

Private Function ApuntaTabla(CODIGO As String) As Boolean
        
ApuntaTabla = True
        
    miSQL2 = "SELECT cod, descri FROM Fctabla1 WHERE CTAB = 'SC   '" & _
    " and DESCRI LIKE '" & ALFAB & "%' ORDER BY DESCRI DESC"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    
    If RgTABL.EOF Then
        MsgBox "No hay registros para esa selección", vbExclamation, " Tablas "
        RgTABL.Close
        Exit Function
    End If
        
ApuntaTabla = False

End Function

Private Function LaDescri(CODIGO As String) As String
Dim I1%
        
    miSQL2 = "SELECT cod, descri FROM Fctabla1 WHERE CTAB = 'SC   ' AND COD = '" & _
             CODIGO & "'"
    RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly
    
    If RgTABL.EOF Then
        LaDescri = ""
        RgTABL.Close
        Exit Function
    End If

LaDescri = RgTABL!Descri
 
RgTABL.Close

End Function

Sub VerSiAgrega()
Dim c1%

If FG1.TextMatrix((FG1.Rows - 1), 1) <> "" Then
                    FG1.Rows = FG1.Rows + 1
                    FG1.Row = FG1.Rows - 1
Else
    FG1.Row = FG1.Rows - 1
                    End If
                    
                    FG1.Col = 0
                    FG1.EditCell

End Sub

Sub VerSiBusca()
        
miSQL2 = "SELECT  * FROM FCTabla1 WHERE COD1 = 'SC   ' and COD2 = '" & _
                         Correa & "'"
RgTAB2.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

    If RgTAB2.EOF = True Then
'        Me.Enabled = False
        Exit Sub
    End If
    
FG1.TextMatrix(FG1.Row, 1) = RgTAB2!Descri

End Sub

Private Function BorraRenglon(Row As Integer)
Dim ColCell As Long

ColCell = FG1.Cell(6, Row, 0, Row, 5)

FG1.Cell(6, FG1.Row, 0, FG1.Row, 5) = vbGreen

vbMsgBoxText = "Desea Eliminar el renglón ?"
vbMsgBoxResp = MsgBox(vbMsgBoxText, vbQuestion + vbYesNo, "Carga Descuentos")

If vbMsgBoxResp = vbNo Then
    FG1.Cell(6, Row, 0, Row, 5) = ColCell
    Exit Function
    End If

FG1.RemoveItem (Row)

End Function

Private Sub CargaGrilla()
        
miSQL2 = "SELECT  * FROM DtoxClte WHERE clte = " & CodCLTE
RgTABL.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

    If RgTABL.EOF = True Then
        RgTABL.Close
        Exit Sub
    End If
    
FG1.Rows = 1
    
Do Until RgTABL.EOF
        
miSQL2 = "SELECT descri FROM FCTabla1 WHERE CTAB = 'SC   ' and COD = '" & _
                         RgTABL!SECCION & "'"
RgTAB2.Open miSQL2, dbTABL, adOpenForwardOnly, adLockReadOnly

If Not RgTAB2.EOF Then
    FG1.AddItem RgTABL!SECCION & vbTab & RgTAB2!Descri & vbTab & RgTABL!dto1 & vbTab & _
                RgTABL!dto2 & vbTab & RgTABL!dto3 & vbTab & RgTABL!dto4 & vbTab & _
                RgTABL!dto5 & vbTab & RgTABL!Rgo1 & vbTab & RgTABL!Rgo2 & vbTab & _
                RgTABL!Rgo3, FG1.Rows
End If

RgTAB2.Close

RgTABL.MoveNext

Loop

RgTABL.Close

FG1.Rows = FG1.Rows + 1
FG1.Row = FG1.Rows - 1

End Sub

Private Sub Grabar()
Dim I1%, i2%, i3%
Dim Mensaje As String

                Mensaje = "Desea Grabar ?"

Respuesta = MsgBox(Mensaje, vbQuestion + vbYesNoCancel, "Descuentos del Cliente")

If Respuesta = vbNo Then Exit Sub
If Respuesta = vbCancel Then
         Exit Sub
         End If

dbDTOS.ConnectionString = BDatos1
dbDTOS.Open

dbDTOS.BeginTrans
   
'*********************************************
'    Elimina todos los descuentos anteriores
'*********************************************
        
miSQL1 = "SELECT * FROM DtoxClte WHERE CLTE = " & CodCLTE
        
RgTABL.Open miSQL1, dbDTOS, adOpenDynamic, adLockPessimistic
    
    If RgTABL.EOF Then
        GoTo AGrabar
        End If
        
Do Until RgTABL.EOF

RgTABL.Delete

RgTABL.MoveNext

Loop

AGrabar:

RgTABL.Close

'******************************************
'      Graba en Descuets para el Cliente
'******************************************

For I1 = 1 To FG1.Rows - 2

Correa = Mid(FG1.TextMatrix(I1, 0), 1, 5)

miSQL1 = "SELECT * FROM DtoxClte WHERE CLTE = " & CodCLTE & _
        " and SECCION = '" & Correa & "'"
RgTABL.Open miSQL1, dbDTOS, adOpenDynamic, adLockPessimistic
    
    If Not RgTABL.EOF Then Exit Sub
    
        RgTABL.AddNew
        
        RgTABL!CLTE = CodCLTE
        RgTABL!SECCION = Correa
        
        RgTABL!dto1 = CCur(FG1.TextMatrix(I1, 2))
        RgTABL!dto2 = CCur(FG1.TextMatrix(I1, 3))
        RgTABL!dto3 = CCur(FG1.TextMatrix(I1, 4))
        RgTABL!dto4 = CCur(FG1.TextMatrix(I1, 5))
        RgTABL!dto5 = CCur(FG1.TextMatrix(I1, 6))
        RgTABL!Rgo1 = CCur(FG1.TextMatrix(I1, 7))
        RgTABL!Rgo2 = CCur(FG1.TextMatrix(I1, 8))
        RgTABL!Rgo3 = CCur(FG1.TextMatrix(I1, 9))

RgTABL!FACTUAL = Date
RgTABL!USUARIO = Mid$(Red_Usuario, 1, 6)

RgTABL.Update

RgTABL.Close

Next I1

dbDTOS.CommitTrans

dbDTOS.Close

Unload Me

End Sub
