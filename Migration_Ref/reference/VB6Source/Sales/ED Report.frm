VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form rptfrmed 
   Caption         =   "ED Report"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Customer"
      Height          =   555
      Left            =   3765
      TabIndex        =   7
      Top             =   3180
      Width           =   2835
      Begin VB.OptionButton optRandom 
         Caption         =   "Random"
         Height          =   195
         Left            =   1620
         TabIndex        =   9
         Top             =   240
         Width           =   900
      End
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   255
         Value           =   -1  'True
         Width           =   570
      End
   End
   Begin VB.CommandButton cbView 
      Caption         =   "&View"
      Height          =   400
      Left            =   3855
      TabIndex        =   6
      Top             =   3735
      Width           =   1300
   End
   Begin VB.Frame Frame2 
      Height          =   765
      Left            =   3750
      TabIndex        =   1
      Top             =   2400
      Width           =   2865
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   110166017
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   75
         TabIndex        =   3
         Top             =   360
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   110166017
         CurrentDate     =   37273
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   210
         Left            =   1380
         TabIndex        =   5
         Top             =   150
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   210
         Left            =   75
         TabIndex        =   4
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   5160
      TabIndex        =   0
      Top             =   3735
      Width           =   1300
   End
Begin AceSpread spCust
      Height          =   1305
      Left            =   6660
      TabIndex        =   10
      Top             =   2475
      Visible         =   0   'False
      Width           =   3915
      _Version        =   458752
      _ExtentX        =   6906
      _ExtentY        =   2302
      _StockProps     =   64
      DAutoSizeCols   =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   2
      OperationMode   =   4
      RowHeaderDisplay=   0
      ScrollBars      =   2
      SelectBlockOptions=   0
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "rptfrmed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public MenuName As String

Public j As Integer
Private Copies As Integer

Public pageno As Integer
Public totalpages As Integer
Private Sub cbExit_Click()

' On Error Resume Next
  Unload Me

End Sub

Private Sub cbView_Click()
Dim frmOutput As frmReportViewer
Try
  
  Me.MousePointer = vbHourglass

  Set frmCurrent = Me
  DoEvents
  iLineNo = 1
  iPageNo = 1
  If sPrintCurrent = False Then
 GoTo ExitHere
End If
 
  sClosePort
  Set frmOutput = New frmReportViewer
  frmOutput.Caption = gStrReportHeading
  frmOutput.txtTotalPages.Text = CStr(iPageNo)
  frmOutput.wbrView.Navigate strFilename
  frmOutput.Tag = strFilename
  gStrReportHeading = ""
  frmOutput.Show vbModal
  DoEvents
  Screen.MousePointer = vbNormal
  Exit Sub
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

''Private Sub cbView1_Click()
''Dim rsDetails As ADODB.Recordset
''  Dim frmOutput As frmReportViewer
''  Dim i As Integer
''  Dim NewPg As Integer
''  Dim sflag As Boolean
''  Dim j As Integer
''  Dim k As Integer
''
''
''  Dim SRow As Long
''  Dim tamt As Double
''  Dim EDAmt As Double
''  Dim CessAmt As Double
''  Dim Cess1Amt As Double
''  Dim cess2amt As Double
''  Dim invamt As Double
''  Dim GV As Double
''  Dim sql As String
''  'Dim sflag As Boolean
''  'Dim i As Integer
''  Dim rCnt As Integer
''  On Error GoTo ErrHandler
''
''  Set rsDetails = New ADODB.Recordset
''
''
''  sql = " select x.InvoiceNo, x.InvoiceDate, x.GV, "
''  sql = sql + " x.ED, x.Cess,x.Cess1 from (select a.InvoiceNo, "
''  sql = sql + " a.InvoiceDate, a.CustomerCode, (select d.PartyName from FAS_Party_Master d"
''  sql = sql + " where d.PartyCode = a.CustomerCode and d.Branchid = a.Branchid "
''  sql = sql + " and d.Compid = a.Compid) CustomerName, a.ItemCode, a.InvoiceQty, "
''  sql = sql + " a.Rate, (a.InvoiceQty * a.Rate) GV, "
''
''  sql = sql + " isNull((select sum(c.ResultValue) from salesCalculationDetails c where "
''  sql = sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
''  sql = sql + " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear "
''  sql = sql + " and c.ExprName in ('EXCISE DUTY','CENVAT CLAIMS')),0) ed,"
''
''  sql = sql + " isNull((select c.ResultValue from salesCalculationDetails c where "
''  sql = sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
''  sql = sql + " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear "
''  sql = sql + " and c.ExprName = 'CESS 2%'),0) CESS,"
''
''  sql = sql + " isNull((select c.ResultValue from salesCalculationDetails c where "
''  sql = sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
''  sql = sql + " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear "
''  sql = sql + " and c.ExprName = 'S&H EDU.CESS 1%'),0) CESS1,"
''
''
''  sql = sql + " a.TotalAmount from SalesInvoice a where a.branchid = '" & gBranchID & "'"
''  sql = sql + " and a.compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
''  sql = sql + " and a.invoiceDate between '" + Format(dpfromdt, "yyyy-MM-dd") + "'"
''  sql = sql + " and '" + Format(dptodt, "yyyy-MM-dd") + "'"
''
''
''  If optRandom Then
''  If spCust.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": GoTo ExitHere
''    sql = sql & " and a.CustomerCode in("
''    With spCust
''      SRow = 0
''      For i = 1 To .SelectionCount
''        SRow = .GetMultiSelItem(SRow)
''        .Row = SRow: .Col = 1
''        sql = sql & "'" & Trim$(.Text) & "',"
''      Next i
''      sql = Left$(sql, Len(sql) - 1)
''    End With
''    sql = sql & ")"
''  End If
''  sql = sql & ") x"
''  sql = sql + " where (x.ED > 0)"
''  If RSOpen(rsDetails, sql) = False Then Exit Sub
''  'If rsDetails.RecordCount = 0 Then
''   ' MsgBox "No records exist"
''    'sPrintcurrent = False
''    'GoTo ExitHere
''  'End If
'''''''''  fOpenPort Me
'''''''''  iLineNo = 1
'''''''''  pheight = 75
'''''''''  pPaper = vbPRPSFanfoldStdGerman
'''''''''  Orient = 1
'''''''''  iPageWidth = 80
'''''''''  tamt = 0
'''''''''  sSendToPort ""
'''''''''  'sSendToPort CAlign(iPageWidth, gCompName)
'''''''''  sSendToPort ""
'''''''''  'gStrReportHeading = "Sales Excise Duty Register between " & dpfromdt & " and " & dptodt
'''''''''  'sSendToPort CAlign(iPageWidth, gStrReportHeading)
'''''''''  sSendToPort ""
''''''''' ' sPageHeader
'''''''''  i = 1
'''''''''  tamt = 0
''''''''''  Do While Not rsDetails.EOF
''''''''''
''''''''''    sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(15, Format(rsDetails("GV"), "0.00")) & Space(1) & RAlign(15, Format(rsDetails("ED"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("Cess"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("Cess1"), "0.00"))
''''''''''    GV = GV + Format(rsDetails("GV"), "0.00")
''''''''''    EDAmt = EDAmt + Format(rsDetails("ed"), "0.00")
''''''''''    CessAmt = CessAmt + Format(rsDetails("cess"), "0.00")
''''''''''    Cess1Amt = Cess1Amt + Format(rsDetails("cess1"), "0.00")
''''''''''     i = i + 1
''''''''''    rsDetails.MoveNext
''''''''''
''''''''''  Loop
'' ' If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
''
''
''  pageno = 1
''  pheight = 10
''  NewPg = 21:
''  sflag = False
''
''
''  If Not rsDetails.EOF Then
''
''   ' With rptfrmed.fpSpread1
''    'rptfrmed.Show
''    '.MaxRows = rsDetails.RecordCount
''    'For rCnt = 2 To rsDetails.RecordCount
''    'sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(15, Format(rsDetails("GV"), "0.00")) & Space(1) & RAlign(15, Format(rsDetails("ED"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("Cess"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("Cess1"), "0.00"))
''    GV = GV + Format(rsDetails("GV"), "0.00")
''    EDAmt = EDAmt + Format(rsDetails("ed"), "0.00")
''    CessAmt = Format(rsDetails("cess"), "0.00")
''    Cess1Amt = Format(rsDetails("cess1"), "0.00")
''    'cess2amt = CessAmt + Cess1Amt
''    .SetText 1, rCnt, rsDetails("invoiceno")
''    .SetText 2, rCnt, rsDetails("invoicedate")
''    .SetText 3, rCnt, rsDetails("gv")
''    .SetText 4, rCnt, rsDetails("ed")
''    .SetText 5, rCnt, rsDetails("cess")
''    .SetText 6, rCnt, rsDetails("cess1")
''    '.SetText 6, rCnt, &cess2
''     rsDetails.MoveNext
''    ' Next rCnt
''
''    .MaxRows = .DataRowCnt
''    'End With
''  End If
''
''  'Loop
'''''  ' For i = 1 To rsDetails.RecordCount
'''''  If Not rsDetails.EOF Then
'''''
'''''    With fpSpread1
'''''      .Row = 3: .Col = 1
'''''      '.TypePictPicture = LoadPicture(App.Path & "\Drawing\" & txtItemCode.Text & ".Bmp")
'''''      '.TypePicMask = 1
'''''     ' .TypePictCenter = True
'''''      '.TypePictMaintainScale = True
'''''     ' .TypePictStretch = True
'''''      '.SetText 9, 33, IIf(Trim$(sNAture) = "R", rs("ChildCode"), rs("StageCode"))
'''''      '.ColWidth(8) = 9
'''''      '.SetText 1, 1, "Doc No : INS -" & rs("RefNo")
'''''      '.SetText 3, 2, UCase(cmbInspectionType.Text)
'''''      .SetText 2, 2, rsDetails("invoiceno")
'''''      '.SetText 9, 35, IIf(Trim$(sNAture) <> "R", rs("OperationName"), "")
'''''      '.SetText 9, 36, IIf(Trim$(sNAture) <> "R", rs("CustomerName"), "")
'''''      '.SetText 7, 38, ""
'''''      '.SetText 9, 38, IIf(Trim$(sNAture) <> "R", rs("PartNo"), "")
'''''      '.SetText 9, 40, rs("RevisionNo")
'''''      '.SetText 11, 39, IIf(Trim$(sNAture) <> "R", rs("StageCode"), rs("ChildCode"))
'''''      '.SetText 13, 36, IIf(Trim$(sNAture) <> "R", rs("CustomerCode"), "")
'''''     'Call cmdPPrint_Click
'''''
'''''
'''''
'''''
'''''    End With
'''''
'''''
'''''  End If
'''''
'''''   If Not rsDetails.EOF Then
'''''        sflag = False
'''''        j = j + 1
'''''        Call cmdPPrint_Click
'''''        rsDetails.MoveNext
'''''        'GoTo sPrint
'''''      Else
'''''        If sflag = False Then
'''''          TotalPages = 1
'''''          GoTo ExitHere
'''''          Call cmdPPrint_Click
'''''          '.ClearRange 1, 6, .MaxCols, 27, False
'''''        End If
'''''      End If
'''''
''
''''''''''  sSendToPort ""
''''''''''  sSendToPort LAlign(18, "Total") & Space(1) & RAlign(20, Format(GV, "0.00")) & Space(1) & RAlign(15, Format(EDAmt, "0.00")) & Space(1) & RAlign(10, Format(CessAmt, "0.00")) & Space(1) & RAlign(10, Format(Cess1Amt, "0.00"))
'''rptfrmed.Show
'' fpSpread1.Visible = True
''  rsDetails.Close
''''''''''  'sSendToPort LAlign(10, "Total") & Space(1) & RAlign(28, Format(GV, "0.00")) & Space(1) & RAlign(15, Format(EDAmt, "0.00")) & Space(1) & RAlign(15, Format(CessAmt, "0.00"))
''''''''''  Call EndofReport(iPageWidth, i - 1)
''  'sPrintcurrent = True
''
''ExitHere:
''  On Error GoTo 0
''  Me.MousePointer = vbDefault
''  Set frmOutput = Nothing
''  'Set rs = Nothing
''  'Set rsCriteria = Nothing
''  'Set rsMachine = Nothing
''  Set rsDetails = Nothing
''  'Set rsEmp = Nothing
''  Exit Sub
''ErrHandler:
''  Call ErrMsgShow
''  Resume ExitHere
''End Sub
''
'Private Sub Cmdexcel_Click()
'Dim dlgSave As CommonDialog
'  Dim retVal As String
'
'
'
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  Set dlgSave = Me.dlgMain
'  dlgSave.CancelError = True
'  dlgSave.FileName = "Report-" & Format(ServerDate, "yyyyMMdd") & ".xls"
'  dlgSave.Filter = "html File (*.xls)|*.xls"
'  dlgSave.ShowSave
'  retVal = Dir$(dlgSave.FileName)
'  If retVal = dlgSave.FileTitle Then
'    If MsgBox("File already Exists. Do you want to overwrite it?", vbYesNo) = vbNo Then
'      GoTo ExitHere
'    Else
'      Kill dlgSave.FileName
'    End If
'  End If
'
'  If spView1.ExportToHTML("Report-" & Format(ServerDate, "yyyyMMdd") & ".xls", False, "") = True Then
'    frm
'    sShowMessage "Export To Excel Succeeded."
'  End If
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub
'
Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  
  Try
  Me.MousePointer = vbHourglass

  CurrDate Me
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Public Sub CurrDate(FormName)
  Dim MyCtrl As Control
  Dim CurrDate As Date
  
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  
  For Each MyCtrl In FormName.Controls
    If TypeOf MyCtrl Is DTPicker Then
      MyCtrl.Value = CurrDate
    End If
  Next
End Sub

Private Sub optAll_Click()

' On Error Resume Next
  spCust.Visible = False
End Sub

Private Sub optIAll_Click()

End Sub

Private Sub optIRandom_Click()

End Sub

Private Sub optRandom_Click()
  Dim rs As AceADODB.Recordset
  Dim Sql As String
 
  Try
  Me.MousePointer = vbHourglass

  Sql = " select distinct a.CustomerCode [Customer Code],(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql + " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) [Customer Name]"
  Sql = Sql + " from SAL_Invoice_Hdr a where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and a.finyear = '" & gFinyear & "' and InvoiceDate between "
  Sql = Sql + " convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111)"
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spCust.MaxRows = spCust.DataRowCnt
'  spCust.Visible = True

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    optAll.Value = True
    GoTo ExitHere
  Else
    spCust.Enabled = True
 spCust.Visible = True
    spCust.ClearSelection
    Call ADORS_Spread_DataFill(rs, spCust, True)
    
    spCust.MaxRows = spCust.DataRowCnt
    Call Spread_Set_Focus(spCust, 1, 1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
  Set rs = Nothing
End Try
End Sub

Private Function sPrintCurrent() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim sRow As Long
  Dim tamt As Double
  Dim EDAmt As Double
  Dim CessAmt As Double
  Dim Cess1Amt As Double
  Dim cess2amt As Double
  Dim cess3amt As Double
  Dim invamt As Double
  Dim GV As Double
  Dim Sql As String
  Dim sflag As Boolean
  Dim i As Integer
  Dim rCnt As Integer
  Try
  
  Set rsDetails = New AceADODB.Recordset
  
  Sql = " select x.InvoiceNo, x.InvoiceDate, x.GV, "
  Sql = Sql + " x.ED, x.Cess,x.Cess1 from (select a.InvoiceNo, "
  Sql = Sql + " a.InvoiceDate, a.CustomerCode, (select d.PartyName from FAS_Party_Master d"
  Sql = Sql + " where d.PartyCode = a.CustomerCode and d.Branchid = a.Branchid "
  Sql = Sql + " and d.Compid = a.Compid) CustomerName, a.ItemCode, a.InvoiceQty, "
  Sql = Sql + " a.Rate, (a.InvoiceQty * a.Rate) GV, "
  
  Sql = Sql + " isNull((select sum(c.ResultValue) from salesCalculationDetails c"
  Sql = Sql + " Where A.invoicetype = c.DocType And A.invoiceno = c.DocNo And A.invoicedate = c.DocDate"
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid and a.finyear = c.finyear"
  Sql = Sql + " and c.ExprGroupId in"
  Sql = Sql + " (select ExprGroupId from SAL_LSTCSTGrouping d where [Description] ='ED' and d.Branchid=c.Branchid"
  Sql = Sql + " and d.Compid=c.Compid)),0) ed,"
  
  Sql = Sql + " isNull((select c.ResultValue from salesCalculationDetails c where  a.invoicetype = c.doctype"
  Sql = Sql + " and a.invoiceno = c.docno and a.invoicedate = c.docdate  and a.branchid = c.branchid"
  Sql = Sql + " and a.compid = c.compid and a.finyear = c.finyear"
  Sql = Sql + " and c.ExprGroupId in"
  Sql = Sql + " (select ExprGroupId from SAL_LSTCSTGrouping d where [Description] ='CESS' and d.Branchid=c.Branchid"
  Sql = Sql + " and d.Compid=c.Compid)),0) CESS,"
  
  Sql = Sql + " isNull((select c.ResultValue from salesCalculationDetails c where a.invoicetype = c.doctype"
  Sql = Sql + " and a.invoiceno = c.docno and a.invoicedate = c.docdate  and a.branchid = c.branchid"
  Sql = Sql + " and a.compid = c.compid and a.finyear = c.finyear"
  Sql = Sql + " and c.ExprGroupId in"
  Sql = Sql + " (select ExprGroupId from SAL_LSTCSTGrouping d where [Description] ='SHE' and d.Branchid=c.Branchid"
  Sql = Sql + " and d.Compid=c.Compid)),0) CESS1,"
  
  Sql = Sql + " a.TotalAmount,a.type from SalesInvoice a where  a.branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  Sql = Sql + " and a.invoiceDate between '" + Format(dpfromdt, "yyyy-MM-dd") + "'"
  Sql = Sql + " and '" + Format(dptodt, "yyyy-MM-dd") + "'"
  
  
  If optRandom Then
  If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    Sql = Sql & " and a.CustomerCode in("
      sRow = 0
      For i = 1 To spCust.SelectionCount
        sRow = spCust.GetMultiSelItem(sRow)
        spCust.Row = sRow
spCust.Col = 1
        Sql = Sql & "'" & Trim$(spCust.Text) & "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql & ")"
  End If
  Sql = Sql & ") x"
  Sql = Sql + " where (x.ED > 0)"
  If RSOpen(rsDetails, Sql) = False Then Exit Function
  If rsDetails.RecordCount = 0 Then
    MsgBox "No records exist"
    sPrintCurrent = False
    GoTo ExitHere
  End If
  fOpenPort Me
  iLineNo = 1
  pheight = 75
  pPaper = vbPRPSFanfoldStdGerman
  Orient = 1
  iPageWidth = 80
  tamt = 0
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  gStrReportHeading = "Sales Excise Duty Register between " & dpfromdt & " and " & dptodt
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort ""
  sPageHeader
  i = 1
  tamt = 0
  Do While Not rsDetails.EOF

    sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(15, Format(rsDetails("GV"), "0.00")) & Space(1) & RAlign(15, Format(rsDetails("ED"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("Cess"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("Cess1"), "0.00"))
    GV = GV + Format(rsDetails("GV"), "0.00")
    EDAmt = EDAmt + Format(rsDetails("ed"), "0.00")
    CessAmt = CessAmt + Format(rsDetails("cess"), "0.00")
    Cess1Amt = Cess1Amt + Format(rsDetails("cess1"), "0.00")
    
     i = i + 1
    rsDetails.MoveNext

  Loop
 
  'If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
 
  
'''''''  If Not rsDetails.EOF Then
'''''''    With fpSpread1
'''''''    .MaxRows = rsDetails.RecordCount
'''''''    For rCnt = 2 To rsDetails.RecordCount
'''''''    'sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(12, rsDetails("InvoiceDate")) & Space(1) & RAlign(15, Format(rsDetails("GV"), "0.00")) & Space(1) & RAlign(15, Format(rsDetails("ED"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("Cess"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("Cess1"), "0.00"))
'''''''    GV = GV + Format(rsDetails("GV"), "0.00")
'''''''    EDAmt = EDAmt + Format(rsDetails("ed"), "0.00")
'''''''    CessAmt = Format(rsDetails("cess"), "0.00")
'''''''    Cess1Amt = Format(rsDetails("cess1"), "0.00")
'''''''    cess2amt = CessAmt + Cess1Amt
'''''''    '.SetText 1, rCnt, rsDetails("invoiceno")
'''''''    '.SetText 2, rCnt, rsDetails("invoicedate")
'''''''    '.SetText 3, rCnt, rsDetails("gv")
'''''''    '.SetText 4, rCnt, rsDetails("cess")
'''''''    '.SetText 5, rCnt, rsDetails("cess1")
'''''''    '.SetText 6, rCnt, &cess2
'''''''     rsDetails.MoveNext
'''''''     Next rCnt
'''''''    End With
'''''''  End If
  
  'Loop
'''  ' For i = 1 To rsDetails.RecordCount
'''  If Not rsDetails.EOF Then
'''
'''    With fpSpread1
'''      .Row = 3: .Col = 1
'''      '.TypePictPicture = LoadPicture(App.Path & "\Drawing\" & txtItemCode.Text & ".Bmp")
'''      '.TypePicMask = 1
'''     ' .TypePictCenter = True
'''      '.TypePictMaintainScale = True
'''     ' .TypePictStretch = True
'''      '.SetText 9, 33, IIf(Trim$(sNAture) = "R", rs("ChildCode"), rs("StageCode"))
'''      '.ColWidth(8) = 9
'''      '.SetText 1, 1, "Doc No : INS -" & rs("RefNo")
'''      '.SetText 3, 2, UCase(cmbInspectionType.Text)
'''      .SetText 2, 2, rsDetails("invoiceno")
'''      '.SetText 9, 35, IIf(Trim$(sNAture) <> "R", rs("OperationName"), "")
'''      '.SetText 9, 36, IIf(Trim$(sNAture) <> "R", rs("CustomerName"), "")
'''      '.SetText 7, 38, ""
'''      '.SetText 9, 38, IIf(Trim$(sNAture) <> "R", rs("PartNo"), "")
'''      '.SetText 9, 40, rs("RevisionNo")
'''      '.SetText 11, 39, IIf(Trim$(sNAture) <> "R", rs("StageCode"), rs("ChildCode"))
'''      '.SetText 13, 36, IIf(Trim$(sNAture) <> "R", rs("CustomerCode"), "")
'''     'Call cmdPPrint_Click
'''
'''
'''
'''
'''    End With
'''
'''
'''  End If
'''
'''   If Not rsDetails.EOF Then
'''        sflag = False
'''        j = j + 1
'''        Call cmdPPrint_Click
'''        rsDetails.MoveNext
'''        'GoTo sPrint
'''      Else
'''        If sflag = False Then
'''          TotalPages = 1
'''          GoTo ExitHere
'''          Call cmdPPrint_Click
'''          '.ClearRange 1, 6, .MaxCols, 27, False
'''        End If
'''      End If
'''
  
  sSendToPort ""
  sSendToPort LAlign(18, "Total") & Space(1) & RAlign(20, Format(GV, "0.00")) & Space(1) & RAlign(15, Format(EDAmt, "0.00")) & Space(1) & RAlign(10, Format(CessAmt, "0.00")) & Space(1) & RAlign(10, Format(Cess1Amt, "0.00"))
  rsDetails.Close
  'sSendToPort LAlign(10, "Total") & Space(1) & RAlign(28, Format(GV, "0.00")) & Space(1) & RAlign(15, Format(EDAmt, "0.00")) & Space(1) & RAlign(15, Format(CessAmt, "0.00"))
  Call EndofReport(iPageWidth, i - 1)
  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

'Public Sub cmdPPrint_Click()
'
'  On Error GoTo ErrHand
'  dlgMain.Flags = &H2&
'
'  dlgMain.Min = 1
'  dlgMain.Max = TotalPages
'  dlgMain.CancelError = True
'  dlgMain.FromPage = 1
'  dlgMain.ToPage = TotalPages
'  dlgMain.Orientation = cdlPortrait
'  dlgMain.PrinterDefault = True
'  If dlgMain.Flags = &H0& Then
'    fpSpread1.PrintPageStart = dlgMain.FromPage
'    fpSpread1.PrintPageEnd = dlgMain.ToPage
'    dlgMain.Flags = &H4&
'  Else
'    fpSpread1.PrintPageStart = dlgMain.FromPage
'    fpSpread1.PrintPageEnd = dlgMain.ToPage
'  End If
'  fpSpread1.BorderStyle = BorderStyleFixedSingle
''  sprPrint.PrintFirstPageNumber = 1
'  fpSpread1.PrintBorder = True
''  sprPrint.PrintType = PrintTypeAll
'  fpSpread1.PrintCenterOnPageV = True
''  sprPrint.FontSize = 7
'  fpSpread1.PrintCenterOnPageH = True
'  fpSpread1.PrintMarginLeft = 1
'  fpSpread1.PrintMarginTop = 600
'  fpSpread1.PrintMarginBottom = 1
'  fpSpread1.PrintMarginRight = 1.25
'  fpSpread1.PrintBorder = True
'  fpSpread1.PrintType = 3
'  fpSpread1.Col = 1: fpSpread1.Col2 = fpSpread1.MaxCols

'  fpSpread1.Row = 1: fpSpread1.Row2 = fpSpread1.MaxRows
'  fpSpread1.BlockMode = True
'  fpSpread1.PrintColor = False
'  For Copies = 1 To dlgMain.Copies
'    fpSpread1.Action = 13
'  Next
'  fpSpread1.BlockMode = True
'  '  sShowMessage "Print Job Over "
'  Exit Sub
'
'ErrHand:
'  If Err.Number = 32755 Then: Exit Sub
'  MsgBox Err.Number & "  " & Err.Description
'
'End Sub
'
'
Public Sub sPageHeader()
  
  Try
  Me.MousePointer = vbHourglass
  sSendToPort String(iPageWidth, "-")
  sSendToPort LAlign(10, "Inv. No.") & Space(1) & RAlign(10, "Inv. Date") & Space(1) & RAlign(17, "Goods Value") & Space(1) & RAlign(14, "ED Value") & Space(2) & RAlign(10, "Cess") & Space(1) & RAlign(10, "Cess1")
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_spCust(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spCust

End Sub

Private Sub InitializeSpreads()

	Call Design_spCust("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\ED Report.frm", spCust, "")

End Sub
