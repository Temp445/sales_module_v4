VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptfrmdaily 
   Caption         =   "Total Component Despatch"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Daily Component Despatch.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame FraItem 
      Caption         =   "Product"
      Height          =   615
      Left            =   3015
      TabIndex        =   8
      Top             =   2415
      Width           =   2655
      Begin VB.OptionButton OptPrdrandom 
         Caption         =   "Random"
         Height          =   195
         Left            =   1425
         TabIndex        =   10
         Top             =   240
         Width           =   900
      End
      Begin VB.OptionButton OptPAll 
         Caption         =   "All"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Height          =   750
      Left            =   2985
      TabIndex        =   15
      Top             =   3015
      Width           =   2655
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "Daily Component Despatch.frx":0C42
         Left            =   90
         List            =   "Daily Component Despatch.frx":0C64
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   330
         Width           =   2340
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Type of Invoice"
         Height          =   210
         Left            =   90
         TabIndex        =   17
         Top             =   150
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   435
      Left            =   2985
      TabIndex        =   12
      Top             =   3855
      Width           =   1365
   End
   Begin VB.Frame FraAccount 
      Caption         =   "Customer"
      Height          =   585
      Left            =   3015
      TabIndex        =   5
      Top             =   1815
      Width           =   2640
      Begin VB.OptionButton OptAccRandom 
         Caption         =   "Random"
         Height          =   195
         Left            =   1425
         TabIndex        =   7
         Top             =   255
         Width           =   900
      End
      Begin VB.OptionButton OptAAll 
         Caption         =   "All"
         Height          =   195
         Left            =   195
         TabIndex        =   6
         Top             =   255
         Value           =   -1  'True
         Width           =   570
      End
   End
   Begin VB.Frame fradate 
      Height          =   735
      Left            =   3030
      TabIndex        =   0
      Top             =   1065
      Width           =   2640
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   345
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   110166019
         CurrentDate     =   37267
      End
      Begin MSComCtl2.DTPicker dpTodt 
         Height          =   345
         Left            =   1305
         TabIndex        =   2
         Top             =   315
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   110166019
         CurrentDate     =   37267
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1305
         TabIndex        =   4
         Top             =   135
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   135
         Width           =   735
      End
   End
Begin AceSpread spCust
      Height          =   1305
      Left            =   5805
      TabIndex        =   13
      Top             =   1140
      Visible         =   0   'False
      Width           =   4575
      _Version        =   458752
      _ExtentX        =   8070
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
Begin AceSpread spItem
      Height          =   1515
      Left            =   5790
      TabIndex        =   14
      Top             =   2460
      Visible         =   0   'False
      Width           =   4710
      _Version        =   458752
      _ExtentX        =   8308
      _ExtentY        =   2672
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
      SelectBlockOptions=   0
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   4275
      TabIndex        =   11
      Top             =   3855
      Width           =   1365
   End
End
Attribute VB_Name = "rptfrmdaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    'For Print Reports
   '    Public TimeCnt As Integer 'For AutoClose Function

    
    
    

'    Public TimeCnt As Integer 'For AutoClose Function

    
    
    

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass

  If DateValue(dpTodt) < DateValue(dpfromdt) Then
    sShowMessage "To Date should be greater than From Date"
    dpTodt.SetFocus
    GoTo ExitHere
  End If
  Set frmCurrent = Me
  DoEvents
  iLineNo = 1
  iPageNo = 1
  
  If Sales = False Then
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
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
  Me.MousePointer = 0
End Try
  End Sub

Private Function Sales() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim ItemCode As String
  Dim Custcode As Integer
  
  Dim sRow As Long
  Dim Sql As String
  Dim i As Long
  Dim tqty As Double
  Dim TotQty As Double
  Dim totqu As Double
  Dim tGV As Double
  Dim Gtot As Double
  Dim mstr As String
  Dim tsqty As Double
  
  Try
  
  Sales = False
  
  Set rsDetails = New AceADODB.Recordset

'  Sql = "Select a.InvoiceType,a.InvoiceNo,a.InvoiceDate,a.CustomerCode,(select b.PartyName "
'  Sql = Sql + " from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid"
'  Sql = Sql + " and b.Compid=a.Compid) PartyName,a.ItemCode,(select c.ItemDescription from INV_MATERIAL_MASTER c"
'  Sql = Sql + " where c.ItemCode=a.ItemCode and c.Branchid=a.Branchid and c.Compid=a.Compid and c.status = 1) ItemDescription,"
'  Sql = Sql + " a.InvoiceQty,a.Rate,(a.InvoiceQty*a.Rate) GV,a.ItemGroupKey from SalesInvoice a"
  
  
  Sql = " Select a.ItemCode, b.partno, c.itemdescription, a.CustomerCode, (select b.PartyName "
  Sql = Sql + " from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode and b.Branchid='" & gBranchID & "'"
  Sql = Sql + " and b.Compid=" & gCompID & ") PartyName, sum(a.InvoiceQty)invoiceqty"
  Sql = Sql + " from SalesInvoice a"
  Sql = Sql + " left outer join SAL_CUSTOMER_ITEMMASTER b on a.itemcode=b.itemcode"
  Sql = Sql + " and a.customercode=b.customercode and a.compid=b.compid"
  Sql = Sql + " left outer join INV_MATERIAL_MASTER c On c.itemcode=a.itemcode and c.compid = a.compid "
  
  Sql = Sql + " where a.InvoiceDate BETWEEN CONVERT(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)  and "
  Sql = Sql + " CONVERT(datetime,'" + Format(dpTodt, "yyyy-MM-dd") + "',111) "
  If OptAccRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and a.CustomerCode in("
      sRow = 0
      For i = 1 To spCust.SelectionCount
        sRow = spCust.GetMultiSelItem(sRow)
        spCust.Row = sRow
spCust.Col = 1
        Sql = Sql + "'" + Trim$(spCust.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  If OptPrdrandom Then
    If spItem.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
Exit Function
End If
    Sql = Sql & " and a.ItemCode in("
      sRow = 0
      For i = 1 To spItem.SelectionCount
        sRow = spItem.GetMultiSelItem(sRow)
        spItem.Row = sRow
spItem.Col = 1
        Sql = Sql + "'" + Trim$(spItem.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql + ")"
  End If
  Sql = Sql + " and a.Branchid ='" & gBranchID & "' and a.Compid = '" & gCompID & "' "
'  Sql = Sql + " and a.finyear = '" & gFinyear & "'"
  If cmbType.Text <> "All" Then
 Sql = Sql + " and a.Type = '" & Trim$(cmbType.Text) & "'"
End If
  Sql = Sql + " group by a.CustomerCode, a.ItemCode, b.partno,c.itemdescription order by a.CustomerCode, a.ItemCode, b.partno"
  If RSOpen(rsDetails, Sql) = False Then Exit Function
  If rsDetails.RecordCount = 0 Then
      MsgBox "No Record(s) Found"
      Screen.MousePointer = vbNormal
      Sales = False
      GoTo ExitHere
  End If
  
  pheight = 75
  pPaper = vbPRPSA4
  Orient = 1
  iPageWidth = 80
  iLineNo = 1
  
  fOpenPort Me
  
  sSendToPort ""
  sSendToPort CAlign(iPageWidth, gCompName)
  sSendToPort ""
  mstr = "Component Despatch between " & dpfromdt & " and " & dpTodt
  sSendToPort CAlign(iPageWidth, mstr)
  sSendToPort ""
  
  sPageHeader
  
  i = 1
  tsqty = 0
  tqty = 0
tGV = 0
Gtot = 0
  Do While Not rsDetails.EOF
   
    If Custcode <> rsDetails("CustomerCode") Then
        If i <> 1 Then
            sSendToPort ""
            sSendToPort LAlign(9, "Sub Total") & Space(52) & RAlign(12, Format(TotQty, 0#))
            sSendToPort ""
            TotQty = 0
        End If
        
        sSendToPort LAlign(60, "CustomerName : " & rsDetails("PartyName"))
        sSendToPort ""
        Custcode = rsDetails("CustomerCode")
    End If
    
      If ItemCode <> rsDetails("ItemCode") Then
'        sSendToPort LAlign(15, Trim(rsDetails("Partno"))) & Space(1) & LAlign(40, rsDetails(("itemdescription"))) & Space(1) & RAlign(8, rsDetails(("InvoiceQty")))
        sSendToPort LAlign(20, Trim(rsDetails("ItemCode"))) & Space(1) & LAlign(39, CheckNull(rsDetails(("itemdescription")))) & Space(1) & RAlign(12, rsDetails(("InvoiceQty")))
        tGV = 0
        tsqty = tsqty + rsDetails("InvoiceQty")
        TotQty = TotQty + rsDetails("InvoiceQty")
      End If
      ItemCode = CheckNull(rsDetails("ItemCode"))
      i = i + 1
      rsDetails.MoveNext
  Loop
   sSendToPort ""
'   sSendToPort LAlign(6, "Total") & Space(45) & RAlign(14, Format(tsqty))
   sSendToPort LAlign(6, "Total") & Space(55) & RAlign(12, Format(tsqty))
   tqty = 0
tGV = 0
 
  rsDetails.Close
  Call EndofReport(iPageWidth, i - 1)
  Sales = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Public Sub sPageHeader()
  Try
  
'  'sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  If CustProdFlag = "C" Then
'    sSendToPort LAlign(12, "Item Code") & Space(1) & LAlign(25, "Item Description") & Space(15) & RAlign(12, "Quantity")
    
    'sSendToPort LAlign(45, " InvoiceDate")
  Else
'     sSendToPort LAlign(12, "Item Code") & Space(1) & LAlign(25, "Item Description") & Space(15) & RAlign(12, "Quantity")
     sSendToPort LAlign(20, "Item Code") & Space(1) & LAlign(39, "Item Description") & Space(1) & RAlign(12, "Quantity")
     
     
    'sSendToPort LAlign(45, "Customer Code") & Space(1) & RAlign(8, "Quantity") & Space(1) & RAlign(11, "Goods Value")
    'sSendToPort LAlign(45, "Customer Name") & Space(1) & RAlign(8, "Rate")
  End If
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sClosePort()
' On Error Resume Next
  Set frmCurrent = Nothing
  Close #iFile
End Sub

Public Function fOpenPort(ByVal frmMe As Form,Optional ByVal isPrintDate As Boolean = True,Optional ByVal isExcel As Boolean = False) As Boolean
  Dim Path As String
  
  Try
  
  Set frmCurrent = frmMe
  
  iPageNo = 1
  iLineNo = 1
  iFile = FreeFile
  Path = Environ("TMP")
  strFilename = IIf(Path = "", App.Path, Path) & "\" & App.Title & "_" & gUserID & Format(Now, "yyyyMMddHHmmss") & ".txt"
  Open strFilename For Output As #iFile
  Print #iFile, "~~1~~"
  
  If isPrintDate = True Then
    Print #iFile, LAlign(30, "Print Date: " & Format(Now, "dd/MM/yyyy HH:mm")) & RAlign(iPageWidth - 30, "Page No.: " & CStr(iPageNo))
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function CAlign(ByVal iLen As Integer,Optional ByVal strText As String = "") As String
  Dim tmp As Integer

' On Error Resume Next

  If Len(strText) = 0 Then
    CAlign = Space(iLen)
  ElseIf iLen <= Len(strText) Then
    CAlign = Left$(strText, iLen)
  Else
    tmp = (iLen - Len(strText)) / 2
    CAlign = Space(tmp) & strText & Space((iLen - (tmp + Len(strText))))
  End If
End Function

Public Function LAlign(ByVal iLen As Integer,Optional ByVal strText As String = "") As String
' On Error Resume Next
  
  If Len(strText) = 0 Then
    LAlign = Space(iLen)
  ElseIf iLen <= Len(strText) Then
  
    LAlign = Left$(strText, iLen)
  Else
    LAlign = strText & Space((iLen - Len(strText)))
  End If
End Function

Public Function RAlign(ByVal iLen As Integer,Optional ByVal strText As String = "") As String
' On Error Resume Next
  
  If Len(strText) = 0 Then
    RAlign = Space(iLen)
  ElseIf iLen <= Len(strText) Then
    RAlign = Left$(strText, iLen)
  Else
    RAlign = Space((iLen - Len(strText))) & strText
  End If
End Function

Public Sub EndofReport(ByVal pagewidth As Integer,ByVal  A As Long)
' On Error Resume Next
  sSendToPort String(pagewidth, "-")
  sSendToPort "No. of Records : " & A & CAlign(iPageWidth - 35, "***** End of Report *****")
sSendToPort String(pagewidth, "-")
End Sub

Public Sub sSendToPort(Optional ByVal strText As String = "",Optional ByVal blnDoNotTerminateLine As Boolean)
' On Error Resume Next
  
  If blnDoNotTerminateLine Then
    Print #iFile, strText;
  Else
    Print #iFile, strText
    iLineNo = iLineNo + 1
    sPerformPageOperations
  End If
  Exit Sub
End Sub

'This routine checks if end of page is reached and if so, prints the page footer, header etc.
Public Sub sPerformPageOperations()
' On Error Resume Next
  If iLineNo > pheight Then
    'sSendToPort LAlign(10, "Total") & Space(1) & RAlign(28, Format(GV, "0.00")) & Space(1) & RAlign(15, Format(EDAmt, "0.00")) & Space(1) & RAlign(15, Format(CessAmt, "0.00"))
    sNewPage
    'frmCurrent.sPageHeader
  End If
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim StDate As Variant
  Dim toDate As Variant
  Dim TmpDate As Variant
  Dim i As Long
  
  Try

  dpfromdt = Format(ServerDate, "dd-mm-yyyy")
  StDate = "01" + "-" + Mid$(dpfromdt, 4, 7)
  toDate = Mid$(dpfromdt, 4, 7)
  
  i = 31
  Do While True
    TmpDate = Trim$(str(i)) + "-" + toDate
    If IsDate(TmpDate) Then Exit Do
    i = i - 1
 Loop
  dpTodt = TmpDate
  cmbType.ListIndex = 0
  
'  If CustProdFlag = "C" Then
''    rptCustomerwiseProductwiseSalesReport.Caption = "Customerwise Productwise Sales"
'    rptfrmdaily.Caption = "Daily Component Deapatch"
'  ElseIf CustProdFlag = "P" Then   'Productwise Supplierwise
''    rptCustomerwiseProductwiseSalesReport.Caption = "Productwise Customerwise Sales"
'  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
  End Sub

Private Sub optAll_Click()

  Try
  Me.MousePointer = vbHourglass

Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub optIAll_Click()
Try
  Me.MousePointer = vbHourglass

  spItem.Visible = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

'Private Sub optIRandom_Click()
'Dim Sql As String
'  Dim sRow As Long
'  Dim i As Long
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  Sql = " select distinct a.ItemCode [Item Code],(select b.ItemDescription from INV_MATERIAL_MASTER b"
'  Sql = Sql + " where b.ItemCode=a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.compid and b.status = 1) [Item Description] "
'  Sql = Sql + " from SalesInvoice a where a.InvoiceDate between convert(datetime,'" + Format(dpfromdt.Value, "yyyy-MM-dd") + "',111)"
'  Sql = Sql + " and convert(datetime,'" + Format(dpTodt.Value, "yyyy-MM-dd") + "',111) "
'  If OptAccRandom Then
'    If spCust.SelectionCount < 1 Then: sShowMessage "Please select atlease one item": OptPAll = True: GoTo ExitHere
'    Sql = Sql & " and a.CustomerCode in("
'    With spCust
'      sRow = 0
'      For i = 1 To .SelectionCount
'        sRow = .GetMultiSelItem(sRow)
'        .Row = sRow: .Col = 1
'        Sql = Sql + "'" + Trim$(.Text) + "',"
'      Next i
'      Sql = Left$(Sql, Len(Sql) - 1)
'    End With
'    Sql = Sql + ")"
'  End If
'  If cmbType <> "All" Then
'    Sql = Sql + " and type = '" + Trim$(cmbType) + "'"
'  End If
'  Sql = Sql & " and a.Branchid='" & gBranchID & "' and compid='" & gCompID & "'"
'  Sql = Sql + "  and finyear='" & gFinyear & "' order by a.ItemCode"
'  adoList1.ConnectionString = db
'  adoList1.RecordSource = Sql
'  adoList1.Refresh
'  spItem.MaxRows = spItem.DataRowCnt
'  spItem.Visible = True
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = 0
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub
'
'
'Private Sub optRandom_Click()
'Dim Sql As String
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  Sql = " select distinct a.CustomerCode [Customer Code],(select b.PartyName from FAS_PARTY_MASTER b where"
'  Sql = Sql + " b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) [Customer Name]"
'  Sql = Sql + " from SAL_INVOICE_HDR a where a.InvoiceDate between"
'  Sql = Sql + " convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
'  Sql = Sql + " and convert(datetime,'" + Format(dpTodt, "yyyy-MM-dd") + "',111)"
'  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
'  Sql = Sql + " and finyear = '" & gFinyear & "'"
'  If cmbType <> "All" Then
'    Sql = Sql + " and type = '" + Trim$(cmbType) + "'"
'  End If
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spCust.MaxRows = spCust.DataRowCnt
'  spCust.Visible = True
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = 0
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub

Private Sub OptAAll_Click()
Try
  Me.MousePointer = vbHourglass

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub OptAccRandom_Click()

  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  Me.MousePointer = vbHourglass

  Sql = " select distinct a.CustomerCode [Customer Code],(select b.PartyName from FAS_PARTY_MASTER b where"
  Sql = Sql + " b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) [Customer Name]"
  Sql = Sql + " from SAL_INVOICE_HDR a where a.InvoiceDate between"
  Sql = Sql + " convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dpTodt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and finyear = '" & gFinyear & "'"
  If cmbType <> "All" Then
    Sql = Sql + " and type = '" + Trim$(cmbType) + "'"
  End If
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
    GoTo ExitHere
  Else
    spCust.Enabled = True
 spCust.Visible = True
    spCust.ClearSelection
    Call ADORS_Spread_DataFill(rs, spCust, True)
    
    spCust.MaxRows = spCust.DataRowCnt
    Call Spread_Set_Focus(spCust, 1, 1)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = 0
  Exit Sub

End Sub

Private Sub OptPAll_Click()
 Try
  Me.MousePointer = vbHourglass

  spItem.Visible = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub OptPrdrandom_Click()
 Dim Sql As String
  Dim sRow As Long
  Dim i As Long
  Dim rs As AceADODB.Recordset

  Try
  Me.MousePointer = vbHourglass

  Sql = " select distinct a.ItemCode [Item Code],(select b.ItemDescription from INV_MATERIAL_MASTER b"
  Sql = Sql + " where b.ItemCode=a.ItemCode and b.Branchid=a.Branchid and b.Compid=a.compid and b.status = 1) [Item Description] "
  Sql = Sql + " from SalesInvoice a where a.InvoiceDate between convert(datetime,'" + Format(dpfromdt.Value, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dpTodt.Value, "yyyy-MM-dd") + "',111) "
  If OptAccRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
OptPAll = True
GoTo ExitHere
End If
    Sql = Sql & " and a.CustomerCode in("
      sRow = 0
      For i = 1 To spCust.SelectionCount
        sRow = spCust.GetMultiSelItem(sRow)
        spCust.Row = sRow
spCust.Col = 1
        Sql = Sql + "'" + Trim$(spCust.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)
    Sql = Sql + ")"
  End If
  If cmbType <> "All" Then
    Sql = Sql + " and type = '" + Trim$(cmbType) + "'"
  End If
  Sql = Sql & " and a.Branchid='" & gBranchID & "' and compid='" & gCompID & "'"
  Sql = Sql + "  and finyear='" & gFinyear & "' order by a.ItemCode"
'  adoList1.ConnectionString = db
'  adoList1.RecordSource = Sql
'  adoList1.Refresh
'  spItem.MaxRows = spItem.DataRowCnt
'  spItem.Visible = True
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    GoTo ExitHere
  Else
    spItem.Enabled = True
 spItem.Visible = True
    spItem.ClearSelection
    Call ADORS_Spread_DataFill(rs, spItem, True)
    
    spItem.MaxRows = spItem.DataRowCnt
    Call Spread_Set_Focus(spItem, 1, 1)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub Design_spCust(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spCust

End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem

End Sub

Private Sub InitializeSpreads()

	Call Design_spCust("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\Daily Component Despatch.frm", spCust, "")

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\Daily Component Despatch.frm", spItem, "")

End Sub
