VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptfrmpreface 
   Caption         =   "RT-12 Preface"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   345
      Left            =   7140
      TabIndex        =   6
      Top             =   7740
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton cmdvExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   8775
      TabIndex        =   5
      Top             =   7740
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4845
      TabIndex        =   3
      Top             =   3300
      Width           =   975
   End
   Begin VB.CommandButton Cmdview 
      Caption         =   "View"
      Height          =   375
      Left            =   3870
      TabIndex        =   2
      Top             =   3300
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTmonth 
      Height          =   345
      Left            =   4215
      TabIndex        =   1
      Top             =   2865
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   609
      _Version        =   393216
      Format          =   56754179
      CurrentDate     =   40208
   End
   Begin VB.ComboBox cmptype 
      Height          =   315
      ItemData        =   "rptfrmpreface.frx":0000
      Left            =   4245
      List            =   "rptfrmpreface.frx":0010
      TabIndex        =   0
      Top             =   2520
      Width           =   1395
   End
Begin AceSpread fppreface
      Height          =   7485
      Left            =   195
      TabIndex        =   4
      Top             =   180
      Visible         =   0   'False
      Width           =   11310
      _Version        =   458752
      _ExtentX        =   19950
      _ExtentY        =   13203
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
      MaxCols         =   10
      MaxRows         =   54
   End
End
Attribute VB_Name = "rptfrmpreface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fpSpread1_Advance(ByVal AdvanceNext As Boolean)

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()

' On Error Resume Next
  
    fppreface.PrintBorder = True
    fppreface.PrintCenterOnPageH = True
    fppreface.PrintCenterOnPageV = False
    fppreface.PrintOrientation = PrintOrientationPortrait
    Printer.PaperSize = vbPRPSA4
    fppreface.Action = ActionPrint
    sShowMessage "Print Job Over"
End Sub

Private Sub cmdvExit_Click()

  fppreface.Visible = False
  cmptype.Visible = True
  DTmonth.Visible = True
  Cmdview.Visible = True
  cmdexit.Visible = True

End Sub

Private Sub Cmdview_Click()
 Try
  Me.MousePointer = vbHourglass

fppreface.Visible = True
fppreface.SetText 1, 14, Format(ServerDate, "dd/MMMM/YYYY")
fppreface.SetText 8, 7, cmptype.Text
fppreface.SetText 6, 23, Format(DTmonth.Value, "MMMM") & "-" & Format(DTmonth, "yyyy")
cmptype.Visible = False
DTmonth.Visible = False
Cmdview.Visible = False
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

Private Sub Design_fppreface(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'fppreface

End Sub

Private Sub InitializeSpreads()

	Call Design_fppreface("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptfrmpreface.frm", fppreface, "")

End Sub
