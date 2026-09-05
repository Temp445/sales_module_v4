VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptfrmer1p5 
   Caption         =   "Form E.R 1 P5"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   345
      Left            =   7140
      TabIndex        =   5
      Top             =   7020
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton cmdvExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   8790
      TabIndex        =   4
      Top             =   7020
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5625
      TabIndex        =   2
      Top             =   3165
      Width           =   975
   End
   Begin VB.CommandButton Cmdview 
      Caption         =   "View"
      Height          =   375
      Left            =   4650
      TabIndex        =   1
      Top             =   3165
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTmonth 
      Height          =   345
      Left            =   4995
      TabIndex        =   3
      Top             =   2730
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   609
      _Version        =   393216
      Format          =   56950787
      CurrentDate     =   40210
   End
Begin AceSpread rptfrmer1p5
      Height          =   6795
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   10995
      _Version        =   458752
      _ExtentX        =   19394
      _ExtentY        =   11986
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
      MaxCols         =   10
      MaxRows         =   48
   End
End
Attribute VB_Name = "rptfrmer1p5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
' On Error Resume Next
  
    rptfrmer1p5.PrintBorder = True
    rptfrmer1p5.PrintCenterOnPageH = True
    rptfrmer1p5.PrintCenterOnPageV = False
    rptfrmer1p5.PrintOrientation = PrintOrientationPortrait
    Printer.PaperSize = vbPRPSA4
    rptfrmer1p5.Action = ActionPrint
    sShowMessage "Print Job Over"
End Sub

Private Sub cmdvExit_Click()
  rptfrmer1p5.Visible = False
  DTmonth.Visible = True
  cmdView.Visible = True
  cmdexit.Visible = True
End Sub

Private Sub Cmdview_Click()
 Dim rec As New AceADODB.Recordset
 Dim rsob As New AceADODB.Recordset
 Dim k As Long
 Dim j As Long
 Dim Value As Variant
 Dim duty As Variant
 Dim cess As Variant
 Dim scess As Variant
 Dim totduty As Variant
 Dim totcess As Variant
 Dim totscess As Variant
 
 Try
 Me.MousePointer = vbHourglass

rptfrmer1p5.Visible = True

Sql = " select a.Plaedcredit,a.Placesscredit  from SAL_Revenue_Particulars a "
Sql = Sql + " where   month(a.Particularsmonth)='" & Format$(DTmonth, "MM") & "' and  year(a.Particularsmonth)='" & Format$(DTmonth, "yyyy") & "'  "

If RSOpen(rec, Sql) = False Then Exit Sub
If rec.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " select min(a.invoiceno)fromi,max(a.invoiceno)toi from salesinvoice a "
Sql = Sql + " where month(a.invoicedate)='" & Format$(DTmonth, "MM") & "' and  year(a.invoicedate)='" & Format$(DTmonth, "yyyy") & "'"

If RSOpen(rsob, Sql) = False Then Exit Sub
If rsob.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

j = 10
For k = 1 To rec.RecordCount

    rptfrmer1p5.SetText 4, 19, Val(rec!Plaedcredit) + Val(rec!Placesscredit)
rptfrmer1p5.FontBold = True

    rptfrmer1p5.SetText 4, 21, rsob!fromi & "To" & rsob!toi
rptfrmer1p5.FontBold = True

    rptfrmer1p5.SetText 3, 30, Format$(ServerDate, "dd/mm/yyyy")
rptfrmer1p5.FontBold = True

    rptfrmer1p5.SetText 8, 40, Format(DTmonth.Value, "MMMM") & "-" & Format(DTmonth, "yyyy")
rptfrmer1p5.FontBold = True

    rptfrmer1p5.SetText 5, 44, Format(DTmonth.Value, "dd")
rptfrmer1p5.FontBold = True

    rptfrmer1p5.SetText 6, 44, Format(DTmonth.Value, "MM")
rptfrmer1p5.FontBold = True

    rptfrmer1p5.SetText 7, 44, Format(DTmonth.Value, "yyyy")
rptfrmer1p5.FontBold = True
     j = j + 1
     
    rec.MoveNext

Next k

DTmonth.Visible = False
cmdView.Visible = False
cmdexit.Visible = False
cmdvExit.Visible = True
cmdprint.Visible = True

  Screen.MousePointer = 1
  totalpages = pageno
  Set rec = Nothing
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

Private Sub Design_rptfrmer1p5(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'rptfrmer1p5

End Sub

Private Sub InitializeSpreads()

	Call Design_rptfrmer1p5("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptfrmer1p5.frm", rptfrmer1p5, "")

End Sub
