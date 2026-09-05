VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptfrmer1p1 
   Caption         =   "Form E.R 1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdvExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   9090
      TabIndex        =   5
      Top             =   7455
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   345
      Left            =   7455
      TabIndex        =   4
      Top             =   7455
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton Cmdview 
      Caption         =   "View"
      Height          =   375
      Left            =   4425
      TabIndex        =   1
      Top             =   2340
      Width           =   975
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   2340
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTmonth 
      Height          =   345
      Left            =   4770
      TabIndex        =   2
      Top             =   1905
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   609
      _Version        =   393216
      Format          =   22347779
      CurrentDate     =   40208
   End
Begin AceSpread fpfrmer1p1
      Height          =   7350
      Left            =   165
      TabIndex        =   3
      Top             =   60
      Visible         =   0   'False
      Width           =   11580
      _Version        =   458752
      _ExtentX        =   20426
      _ExtentY        =   12965
      _StockProps     =   64
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowHoriz   =   0   'False
      GridShowVert    =   0   'False
      MaxCols         =   12
      MaxRows         =   34
   End
End
Attribute VB_Name = "rptfrmer1p1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
' On Error Resume Next
  
    fpfrmer1p1.PrintBorder = True
    fpfrmer1p1.PrintCenterOnPageH = True
    fpfrmer1p1.PrintCenterOnPageV = False
    fpfrmer1p1.PrintOrientation = PrintOrientationLandscape
    Printer.PaperSize = vbPRPSA4
    fpfrmer1p1.Action = ActionPrint
    sShowMessage "Print Job Over"
End Sub

Private Sub cmdvExit_Click()
  fpfrmer1p1.Visible = False
  DTmonth.Visible = True
  cmdView.Visible = True
  cmdexit.Visible = True

End Sub

Private Sub Cmdview_Click()
  Try
  Me.MousePointer = vbHourglass

fpfrmer1p1.Visible = True
fpfrmer1p1.SetText 9, 12, Format(DTmonth.Value, "MMMM") & "-" & Format(DTmonth, "yyyy")

DTmonth.Visible = False
cmdView.Visible = False
cmdexit.Visible = False
cmdvExit.Visible = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  Try
  
  DTmonth.Value = Format$(ServerDate, "dd/mm/yyyy")
 
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Design_fpfrmer1p1(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fpfrmer1p1

End Sub

Private Sub InitializeSpreads()

	Call Design_fpfrmer1p1("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptfrmer1p1.frm", fpfrmer1p1, "")

End Sub
