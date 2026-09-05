VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptfrmer1p3 
   Caption         =   "Form E.R.1 P3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdvExit 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   9120
      TabIndex        =   5
      Top             =   7380
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   345
      Left            =   7485
      TabIndex        =   4
      Top             =   7380
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton Cmdview 
      Caption         =   "View"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3435
      Width           =   975
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5310
      TabIndex        =   1
      Top             =   3435
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DTmonth 
      Height          =   345
      Left            =   4680
      TabIndex        =   3
      Top             =   3000
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   609
      _Version        =   393216
      Format          =   22282243
      CurrentDate     =   40210
   End
Begin AceSpread rptfrmer1p3
      Height          =   7185
      Left            =   195
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   11205
      _Version        =   458752
      _ExtentX        =   19764
      _ExtentY        =   12674
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
      MaxRows         =   40
   End
End
Attribute VB_Name = "rptfrmer1p3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsob As New AceADODB.Recordset

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
' On Error Resume Next
  
    rptfrmer1p3.PrintBorder = True
    rptfrmer1p3.PrintCenterOnPageH = True
    rptfrmer1p3.PrintCenterOnPageV = False
    rptfrmer1p3.PrintOrientation = PrintOrientationLandscape
    Printer.PaperSize = vbPRPSA4
    rptfrmer1p3.Action = ActionPrint
    sShowMessage "Print Job Over"
End Sub

Private Sub cmdvExit_Click()
rptfrmer1p3.Visible = False
  DTmonth.Visible = True
  cmdView.Visible = True
  cmdexit.Visible = True
End Sub

Private Sub Cmdview_Click()

 Dim rec As New AceADODB.Recordset
 Dim rec1 As New AceADODB.Recordset
 Dim Rec2 As New AceADODB.Recordset
 Dim rec3 As New AceADODB.Recordset
 Dim rsob As New AceADODB.Recordset
 Dim k As Long
 Dim j As Long
 Dim Value As Variant
 Dim duty As Variant
 Dim cess1 As Variant
 Dim cess2 As Variant
 Dim pacess1 As Variant
 Dim pacess2 As Variant
 Dim cess3 As Variant
 Dim scess As Variant
 Dim totduty As Variant
 Dim totcess As Variant
 Dim totscess As Variant
 
 Try
 Me.MousePointer = vbHourglass

rptfrmer1p3.Visible = True

Sql = " select * from SAL_Revenue_Particulars a "
Sql = Sql + " where   month(a.Particularsmonth)='" & Format$(DTmonth, "MM") & "' and  year(a.Particularsmonth)='" & Format$(DTmonth, "yyyy") & "'  "

If RSOpen(rec, Sql) = False Then Exit Sub
If rec.EOF Then
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

Sql = "  Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
Sql = Sql + " Where month(a.RefDate) ='" & Format$(DTmonth, "MM") & "' and year(a.RefDate) = '" & Format$(DTmonth, "yyyy") & "'  and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='inv' group by c.LedgerName,c.ledgercode "

If RSOpen(Rec2, Sql) = False Then Exit Sub
If Rec2.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
Sql = Sql + " Where  A.refno BETWEEN '" & rec!placessinvno & "' AND '" & rec!Placessinvno1 & "' AND month(a.RefDate) ='" & Format$(DTmonth, "MM") & "' and year(a.RefDate) = '" & Format$(DTmonth, "yyyy") & "'  and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='inv' group by c.LedgerName,c.ledgercode "

Sql = Sql + " union Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
Sql = Sql + " Where  A.refno in ( '" & rec!Placessinvno2 & "','" & rec!Placessinvno3 & "','" & rec!Placessinvno4 & "','" & rec!Placessinvno5 & "','" & rec!Placessinvno6 & "') AND month(a.RefDate) ='" & Format$(DTmonth, "MM") & "' and year(a.RefDate) = '" & Format$(DTmonth, "yyyy") & "'  and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='inv' group by c.LedgerName,c.ledgercode "

If RSOpen(rec1, Sql) = False Then Exit Sub
If rec1.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

Sql = " Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
Sql = Sql + " Where  A.refno BETWEEN '" & rec!Partacessinvno & "' AND '" & rec!Partacessinvno1 & "' AND month(a.RefDate) ='" & Format$(DTmonth, "MM") & "' and year(a.RefDate) = '" & Format$(DTmonth, "yyyy") & "'  and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='inv' group by c.LedgerName,c.ledgercode "

Sql = Sql + " union Select distinct  c.ledgercode,c.LedgerName, sum(b.DrAmt)DrAmt,sum( b.CrAmt)CrAmt From FAS_VoucherTranRef a Left Outer Join FAS_VoucherTranSub b  On a.DocType = b.DocType and a.DocNo = b.DocNo and a.DocDate = b.DocDate  and a.VrRoute = b.VrRoute and a.CompCode = b.CompCode Left Outer Join FAS_LedgerMaster c On b.LrCode = c.LedgerCode and b.CompCode = c.CompCode "
Sql = Sql + " Where  A.refno in ( '" & rec!Partacessinvno2 & "','" & rec!Partacessinvno3 & "','" & rec!Partacessinvno4 & "','" & rec!Partacessinvno5 & "','" & rec!Partacessinvno6 & "') AND month(a.RefDate) ='" & Format$(DTmonth, "MM") & "' and year(a.RefDate) = '" & Format$(DTmonth, "yyyy") & "'  and a.CompCode = '0200' and b.VTSNo <> 1 and  a.RefType='inv' group by c.LedgerName,c.ledgercode "

If RSOpen(rec3, Sql) = False Then Exit Sub
If rec3.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

For k = 1 To Rec2.RecordCount
  
    Select Case Rec2("ledgercode")
    Case "8"
    duty = Round(Trim$(CheckNull((Rec2!CrAmt))))
    Case "9"
     cess3 = cess3 + Rec2!CrAmt
    Case "340"
     cess3 = cess3 + Rec2!CrAmt
    
   
  End Select
    Rec2.MoveNext
  
  Next k

For k = 1 To rec1.RecordCount
  
    Select Case rec1("ledgercode")
    
    Case "9"
     cess2 = cess2 + rec1!CrAmt
    Case "340"
     cess1 = cess1 + rec1!CrAmt
    
   
  End Select
    rec1.MoveNext
  
  Next k
 
 For k = 1 To rec3.RecordCount
  
    Select Case rec3("ledgercode")
    
    Case "9"
     pacess2 = pacess2 + rec3!CrAmt
    Case "340"
     pacess1 = pacess1 + rec3!CrAmt
    
   
  End Select
    rec3.MoveNext
  
  Next k
  

For k = 1 To rec.RecordCount

    rptfrmer1p3.SetText 8, 22, rsob!plaedcb
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 9, 22, rsob!placesscb
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 8, 24, rec!Plaedcredit
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 9, 24, rec!Placesscredit
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 8, 29, (Val(rsob!plaedcb) + Val(rec!Plaedcredit))
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 9, 29, (Val(rsob!placesscb) + Val(rec!Placesscredit))
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 8, 14, duty
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 8, 15, cess3
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 3, 14, Val(rec!Plaedsaldebit)
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 4, 14, duty - Val(rec!Plaedsaldebit)
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 3, 15, cess2
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 3, 16, cess1
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 4, 15, pacess2
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 4, 16, pacess1
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 8, 31, rec!Plaedsaldebit
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 9, 31, rec!Placesssaldebit
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 8, 37, ((Val(rsob!plaedcb) + Val(rec!Plaedcredit)) - Val(rec!Plaedsaldebit))
rptfrmer1p3.FontBold = True

    rptfrmer1p3.SetText 9, 37, ((Val(rsob!placesscb) + Val(rec!Placesscredit)) - Val(rec!Placesssaldebit))
rptfrmer1p3.FontBold = True
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
  Set rec1 = Nothing
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

Private Sub Design_rptfrmer1p3(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'rptfrmer1p3

End Sub

Private Sub InitializeSpreads()

	Call Design_rptfrmer1p3("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptfrmer1p3.frm", rptfrmer1p3, "")

End Sub
