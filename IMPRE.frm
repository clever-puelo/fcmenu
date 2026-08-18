VERSION 5.00
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form IMPRE 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Impresión"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5940
      TabIndex        =   4
      Top             =   5400
      Width           =   990
   End
   Begin VSPrinter8LibCtl.VSPrinter VP1 
      Height          =   5175
      Left            =   300
      TabIndex        =   3
      Top             =   105
      Width           =   6675
      _cx             =   11774
      _cy             =   9128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   21.6517857142857
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Configurar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2505
      TabIndex        =   2
      Top             =   5400
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1395
      TabIndex        =   1
      Top             =   5400
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   0
      Top             =   5400
      Width           =   990
   End
End
Attribute VB_Name = "IMPRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Unload Me

End Sub

Private Sub Command4_Click()

VP1.PrintDoc
   

End Sub

Private Sub Command2_Click()
     
  vbMsgBoxTitle = "Cancelación del Proceso"
  vbMsgBoxText = "El siguiente proceso Cancelará" & vbCrLf & " la impresión del Comprobante  " & vbCrLf & "Desea continuar ?  "
  vbMsgBoxResp = vbYesNo + vbExclamation + vbApplicationModal + vbDefaultButton1

  vbResponse = MsgBox(vbMsgBoxText, vbMsgBoxResp, vbMsgBoxTitle)

  Select Case vbResponse

         Case vbYes
              SinCpbte = True

         Case vbNo
              Exit Sub
  End Select

VP1.KillDoc
Unload Me

End Sub

Private Sub Command3_Click()

VP1.PrintDialog pdPrinterSetup

End Sub

Private Sub Form_Load()

Me.Move 200, 1260, 8500, 8500
VP1.top = 100
VP1.left = 100

SinCpbte = False

End Sub

Private Sub Form_Resize()

If IMPRE.width < 3800 Then IMPRE.width = 3800

VP1.width = IMPRE.width - 1000
VP1.height = IMPRE.height - 2000

Command1.top = IMPRE.height - 1800
Command2.top = IMPRE.height - 1800
Command3.top = IMPRE.height - 1800
Command4.top = IMPRE.height - 1800

End Sub

Private Sub VP1_ReadyStateChange(ByVal ReadyState As Long)

Select Case ReadyState

    Case vpstEmpty
        VP1.Enabled = False
        VP1.Visible = False
        
    Case vpstLoading
        
    Case vpstReady
        VP1.Enabled = True
        VP1.Visible = True
        
    Case vpstOpen
        
    Case vpstSaving
        Me.Enabled = False
        Me.Visible = False
       
    Case vpstPrinting
        Unload Me

End Select

End Sub


