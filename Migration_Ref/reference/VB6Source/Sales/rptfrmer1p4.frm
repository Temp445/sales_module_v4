VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptfrmer1p4 
   Caption         =   "Form E.R 1 P4"
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
      Left            =   7545
      TabIndex        =   5
      Top             =   6870
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton cmdvExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   9180
      TabIndex        =   4
      Top             =   6870
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5190
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Cmdview 
      Caption         =   "View"
      Height          =   375
      Left            =   4215
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTmonth 
      Height          =   345
      Left            =   4635
      TabIndex        =   3
      Top             =   2370
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   609
      _Version        =   393216
      Format          =   56885251
      CurrentDate     =   40210
   End
Begin AceSpread rptfrmer1p4
      Height          =   6705
      Left            =   165
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   11220
      _Version        =   458752
      _ExtentX        =   19791
      _ExtentY        =   11827
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
      MaxRows         =   24
   End
End
Attribute VB_Name = "rptfrmer1p4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
' On Error Resume Next
  
    rptfrmer1p4.PrintBorder = True
    rptfrmer1p4.PrintCenterOnPageH = True
    rptfrmer1p4.PrintCenterOnPageV = False
    rptfrmer1p4.PrintOrientation = PrintOrientationPortrait
    Printer.PaperSize = vbPRPSA4
    rptfrmer1p4.Action = ActionPrint
    sShowMessage "Print Job Over"
End Sub

Private Sub cmdvExit_Click()
rptfrmer1p4.Visible = False
  DTmonth.Visible = True
  cmdView.Visible = True
  cmdexit.Visible = True
End Sub

Private Sub Cmdview_Click()

 Dim rec As New AceADODB.Recordset
 Dim rec1 As New AceADODB.Recordset
 Dim rsob As New AceADODB.Recordset
 Dim rsob1 As New AceADODB.Recordset
 Dim k As Long
 Dim j As Long
 Dim cb As Variant
 Dim cb1 As Variant
 Dim cb2 As Variant
 Dim cb3 As Variant
 Dim Value As Variant
 Dim duty As Variant
 Dim cess As Variant
 Dim scess As Variant
 Dim totduty As Variant
 Dim totcess As Variant
 Dim totscess As Variant
 
 Try
 Me.MousePointer = vbHourglass

rptfrmer1p4.Visible = True

Sql = " select * from SAL_Revenue_Particulars a "
Sql = Sql + " where   month(a.Particularsmonth)='" & Format$(DTmonth, "MM") & "' and  year(a.Particularsmonth)='" & Format$(DTmonth, "yyyy") & "'  "

If RSOpen(rec, Sql) = False Then Exit Sub
If rec.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " select * from SAL_Revenue_Abstract a "
Sql = Sql + " where   month(a.Particularsmonth)='" & Format$(DTmonth, "MM") & "' and  year(a.Particularsmonth)='" & Format$(DTmonth, "yyyy") & "'  "

If RSOpen(rec1, Sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " select * from SAL_Revenue_Particulars a "
Sql = Sql + " where   month(a.Particularsmonth)<convert(datetime,'" & Format(DTmonth, "dd-MM-yyyy") & "',111)  order by a.Particularsmonth desc "

If RSOpen(rsob, Sql) = False Then Exit Sub
If rsob.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " select * from SAL_Revenue_Abstract a "
Sql = Sql + " where   month(a.Particularsmonth)<convert(datetime,'" & Format(DTmonth, "dd-MM-yyyy") & "',111)  order by a.Particularsmonth desc "

If RSOpen(rsob1, Sql) = False Then Exit Sub
If rsob1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

j = 10
For k = 1 To rec.RecordCount

    rptfrmer1p4.SetText 2, 7, Val(rsob!partcedcb) + Val(rsob1!inputedcb)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 7, 7, Val(rsob!partccesscb) + Val(rsob1!inputcesscb)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 8, 7, rsob1!Servicetaxedcb
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 9, 7, rsob1!Servicetaxcesscb
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 2, 9, Val(rec1!inputedcredit)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 7, 9, (Val(rec1!inputcesscredit))
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 2, 11, Val(rec!Partcedcredit)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 7, 11, (Val(rec!Partccesscredit))
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 8, 13, Val(rec1!Servicetaxedcredit)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 9, 13, (Val(rec1!Servicetaxcesscredit))
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 2, 15, Val(Val(rsob!partcedcb) + Val(rsob1!inputedcb)) + Val(rec1!inputedcredit) + Val(rec!Partcedcredit)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 7, 15, Val(rsob!partccesscb) + Val(rsob1!inputcesscb) + (Val(rec1!inputcesscredit)) + (Val(rec!Partccesscredit))
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 8, 15, Val(rsob1!Servicetaxedcb) + Val(rec1!Servicetaxedcredit)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 9, 15, Val(rsob1!Servicetaxcesscb) + Val(rec1!Servicetaxcesscredit)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 2, 16, (Val(rec1!inputedsaldebit))
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 7, 16, (Val(rec1!inputcesssaldebit))
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 8, 16, (Val(rec1!Servicetaxedsaldebit))
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 9, 16, (Val(rec1!Servicetaxcesssaldebit))
rptfrmer1p4.FontBold = True
    cb = Val(Val(Val(rsob!partcedcb) + Val(rsob1!inputedcb)) + Val(rec1!inputedcredit) + Val(rec!Partcedcredit)) - (Val(rec1!inputedsaldebit))
    cb1 = Val(Val(rsob!partccesscb) + Val(rsob1!inputcesscb) + (Val(rec1!inputcesscredit)) + (Val(rec!Partccesscredit))) - (Val(rec1!inputcesssaldebit))
    cb2 = Val(Val(rsob1!Servicetaxedcb) + Val(rec1!Servicetaxedcredit)) - (Val(rec1!Servicetaxedsaldebit))
    cb3 = Val(Val(rsob1!Servicetaxcesscb) + Val(rec1!Servicetaxcesscredit)) - (Val(rec1!Servicetaxcesssaldebit))
    rptfrmer1p4.SetText 2, 22, Val(cb)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 7, 22, Val(cb1)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 8, 22, Val(cb2)
rptfrmer1p4.FontBold = True

    rptfrmer1p4.SetText 9, 22, Val(cb3)
rptfrmer1p4.FontBold = True
    
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

Private Sub Design_rptfrmer1p4(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'rptfrmer1p4

End Sub

Private Sub InitializeSpreads()

	Call Design_rptfrmer1p4("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptfrmer1p4.frm", rptfrmer1p4, "")

End Sub
