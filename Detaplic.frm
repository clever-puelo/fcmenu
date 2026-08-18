VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsFlex8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form DetAplic 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aplicado"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   30
      TabIndex        =   2
      Top             =   3240
      Width           =   6405
      Begin VB.CommandButton Command1 
         Caption         =   "Anticipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   5
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pagos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1530
         TabIndex        =   4
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5010
         TabIndex        =   3
         Top             =   180
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6450
      Begin VSFlex8Ctl.VSFlexGrid FG1 
         Height          =   3045
         Left            =   60
         TabIndex        =   1
         Top             =   135
         Width           =   6330
         _cx             =   11165
         _cy             =   5371
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
         SelectionMode   =   3
         GridLines       =   3
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"DetAplic.frx":0000
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
         Editable        =   2
         ShowComboButton =   -1  'True
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
      End
   End
   Begin MSAdodcLib.Adodc CCTE 
      Height          =   330
      Left            =   240
      Top             =   4560
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=CTACTE"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "CTACTE"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "CCTE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "DetAplic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()

DetPago.Show

End Sub

Private Sub Command3_Click()

Unload Me

End Sub

Private Sub FG1_Click()

If FG1.Row > 1 Then
    If FG1.TextMatrix(FG1.Row, 3) > 0 Then
        FG1.TextMatrix(FG1.Row, 3) = 0
    Else
        FG1.TextMatrix(FG1.Row, 3) = FG1.TextMatrix(FG1.Row, 4)
    End If
End If

CalcAplic

End Sub


Private Sub Form_Load()

Me.Move 0, 4000, 7000, 4500
CtaCte.Label2.Caption = CodClte & "  " & ClteNomb


End Sub

Private Sub Form_Unload(Cancel As Integer)

Recibo.Toolbar2.Buttons(2).Value = tbrUnpressed

End Sub


Private Sub CalcAplic()
Dim A1 As Integer, ImpAplic As Currency, ImpDebe As Currency

For A1 = 1 To FG1.Rows - 1

    ImpAplic = ImpAplic + FG1.TextMatrix(A1, 3)
    
    If FG1.TextMatrix(A1, 3) > 0 Then FG1.IsSelected(A1) = True
    
Next A1

PieRec.Label302.Caption = Format(ImpAplic, "$###,##0.00")


End Sub



