VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptSalesTaxRegisterReport 
   Caption         =   "Sales Tax Register Report"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin VB.Frame frastep 
      Caption         =   "Invoice Type"
      Height          =   660
      Index           =   2
      Left            =   4650
      TabIndex        =   23
      Top             =   2340
      Width           =   2610
      Begin VB.ComboBox CmbList 
         Height          =   330
         ItemData        =   "SalesTaxRegisterReport.frx":0000
         Left            =   120
         List            =   "SalesTaxRegisterReport.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   210
         Width           =   2370
      End
   End
   Begin VB.CommandButton CmdStep 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   7365
      TabIndex        =   22
      Top             =   1905
      Width           =   405
   End
   Begin VB.CommandButton CmdStep 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   7365
      TabIndex        =   21
      Top             =   1245
      Width           =   405
   End
   Begin VB.Frame frastep 
      Caption         =   "Customer"
      Enabled         =   0   'False
      Height          =   555
      Index           =   1
      Left            =   4650
      TabIndex        =   18
      Top             =   1785
      Width           =   2610
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   60
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   570
      End
      Begin VB.OptionButton optRandom 
         Caption         =   "Selective"
         Height          =   225
         Left            =   1575
         TabIndex        =   19
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame frastep 
      Height          =   765
      Index           =   0
      Left            =   4650
      TabIndex        =   13
      Top             =   1020
      Width           =   2610
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1365
         TabIndex        =   14
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218890241
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   75
         TabIndex        =   15
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   218890241
         CurrentDate     =   37273
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   210
         Left            =   75
         TabIndex        =   17
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   210
         Left            =   1365
         TabIndex        =   16
         Top             =   150
         Width           =   555
      End
   End
   Begin VB.CommandButton CmdStep 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   7365
      TabIndex        =   12
      Top             =   2505
      Width           =   405
   End
   Begin VB.Frame frastep 
      Caption         =   "Report Type"
      Height          =   930
      Index           =   3
      Left            =   4650
      TabIndex        =   9
      Top             =   2985
      Width           =   2610
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "SalesTaxRegisterReport.frx":0038
         Left            =   90
         List            =   "SalesTaxRegisterReport.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   210
         Width           =   2370
      End
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print.Date"
         Height          =   255
         Left            =   90
         TabIndex        =   10
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
   End
   Begin VB.CommandButton CmdStep 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   3
      Left            =   7365
      TabIndex        =   8
      Top             =   3225
      Width           =   405
   End
   Begin VB.Frame Frame3 
      Height          =   555
      Left            =   4335
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   3090
      Begin VB.OptionButton OptCst 
         Caption         =   "CST"
         Height          =   210
         Left            =   915
         TabIndex        =   7
         Top             =   225
         Width           =   600
      End
      Begin VB.OptionButton optTngst 
         Caption         =   "TNGST / VAT"
         Height          =   210
         Left            =   1740
         TabIndex        =   6
         Top             =   225
         Width           =   1275
      End
      Begin VB.OptionButton optTAll 
         Caption         =   "ALL"
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   225
         Value           =   -1  'True
         Width           =   600
      End
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6660
      TabIndex        =   0
      Top             =   4125
      Width           =   1305
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5370
      TabIndex        =   3
      Top             =   4125
      Width           =   1305
   End
Begin AceSpread spCust
      Height          =   1290
      Left            =   7845
      TabIndex        =   2
      Top             =   1905
      Visible         =   0   'False
      Width           =   4095
      _Version        =   458752
      _ExtentX        =   7223
      _ExtentY        =   2275
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
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4080
      TabIndex        =   1
      Top             =   4125
      Width           =   1300
   End
End
Attribute VB_Name = "rptSalesTaxRegisterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim k As Integer
Dim j As Long

Public Sub sView()
  'No Change
End Sub
Public Sub sprint()
  'No Change
End Sub
Private Sub Header()
  Dim mstr As String
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
    frmPrint.sprPrint.MaxCols = 8

    frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
    
    frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j
    frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
    If chkPrintDate.Value = 1 Then
 frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(15) & "Printed Date : " & Format(ServerDate, "dd/MM/yyyy HH:MM:ss")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
End If
    j = j + 1
GridCount = GridCount + 1
    
'    Sql = " Select * from dbo.fn_SAL_Menu_List('" & Trim$(mnuFlg) & "')"
'    Set rs = New Recordset
'    If RSOpen(rs, Sql) = False Then GoTo ExitHere
'    If rs.RecordCount > 0 Then
'      .AddCellSpan 1, j, 3, 1
'      .Col = 1: .Row = j

'      .SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo"): .TypeHAlign = TypeHAlignLeft
'
'      .Col = 6: .Row = j
'      .AddCellSpan .Col, j, .MaxCols, 1
'      .SetText .Col, j, Space(15) & "TS Ref. No   : " & rs("TsNo"): .TypeHAlign = TypeHAlignLeft

'      j = j + 1: GridCount = GridCount + 1
'    End If

    If iPageNo = 1 Then
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.SetText 1, j, gCompName
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    End If
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
    frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
    frmPrint.sprPrint.SetText 1, j, "Sales Tax Register between " & dpfromdt & " and " & dptodt
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 10
frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
j = j + 2
GridCount = GridCount + 2
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.SetText 1, j, "Invoice"
frmPrint.sprPrint.ColWidth(1) = 6
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

    frmPrint.sprPrint.SetText 2, j, "Invoice"
frmPrint.sprPrint.ColWidth(2) = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
  
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

    frmPrint.sprPrint.SetText 3, j, "Customer Name"
frmPrint.sprPrint.ColWidth(3) = 25
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

    frmPrint.sprPrint.SetText 4, j, "Taxable"
frmPrint.sprPrint.ColWidth(4) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

    frmPrint.sprPrint.SetText 5, j, "SalesTax"
frmPrint.sprPrint.ColWidth(5) = 7
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

    frmPrint.sprPrint.SetText 6, j, "Sur"
frmPrint.sprPrint.ColWidth(6) = 6
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

    frmPrint.sprPrint.SetText 7, j, "SalesTax"
frmPrint.sprPrint.ColWidth(7) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

    frmPrint.sprPrint.SetText 8, j, "Invoice"
frmPrint.sprPrint.ColWidth(8) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
    j = j + 1
GridCount = GridCount + 1
    
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.SetText 1, j, "No"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

    frmPrint.sprPrint.SetText 2, j, "Date"
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

    frmPrint.sprPrint.SetText 4, j, "Amount"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

    frmPrint.sprPrint.SetText 5, j, "Amount"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

    frmPrint.sprPrint.SetText 6, j, "Charge"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

    frmPrint.sprPrint.SetText 7, j, "Total"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

    frmPrint.sprPrint.SetText 8, j, "Amount"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    Spread_BkColour frmPrint.sprPrint, &H8000000F, 1, j, frmPrint.sprPrint.MaxCols, j
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub RowProcess()
' On Error Resume Next
  If GridCount >= pheight Then
      GridCount = 1
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.MaxRows + pheight + 1
      frmPrint.sprPrint.Row = frmPrint.sprPrint.Row + 1
j = j + 1
      frmPrint.sprPrint.RowPageBreak = 1
      iPageNo = iPageNo + 1
      Header
      j = j + 1
GridCount = GridCount + 1
  End If
End Sub

Private Function sPrintSelect() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim sRow As Long
  Dim tamt As Double
  Dim StAmt As Double
  Dim suramt As Double
  Dim invamt As Double
  Dim GV As Double
  Dim Sql As String
  Dim i As Integer
  
  Try
  
  sPrintSelect = False
  
  Set rsDetails = New AceADODB.Recordset
  
  Sql = " select x.InvoiceNo, x.InvoiceDate, x.CustomerName, x.Labor, Case When x.ED=0 Then 0 Else (x.GV+x.ED+x.CESS) End as TaxableIncome, "
  Sql = Sql + " x.TNGST, x.CST, x.TNGSTSc, x.EOU, x.TotalAmount  "
  Sql = Sql + " from (select a.InvoiceNo, a.InvoiceDate, a.CustomerCode, "
  Sql = Sql + " (select d.PartyName from FAS_Party_Master d where d.PartyCode = a.CustomerCode and d.Branchid = a.Branchid and d.Compid = a.Compid) CustomerName,"
  Sql = Sql + " a.ItemCode, a.InvoiceQty, a.Rate, a.TotalAmount, (a.InvoiceQty * a.Rate) GV,"
  
  Sql = Sql + " isNull((select sum(c.ResultValue) from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from SAL_LSTCSTGrouping d where d.Description = 'ED'"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) ed,"

  Sql = Sql + " isNull((select c.ResultValue from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from SAL_LSTCSTGrouping d where d.Description = 'CESS'"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) CESS,"

  Sql = Sql + " isNull((select c.ResultValue from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from SAL_LSTCSTGrouping d where d.Description = 'LST'"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) TNGST,"
  
  Sql = Sql + " isNull((select c.ResultValue from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from SAL_LSTCSTGrouping d where d.Description = 'CST'"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) CST,"
  
  Sql = Sql + " isNull((select c.ResultValue from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from SAL_LSTCSTGrouping d where d.Description in ('CST Surcharge', 'LST Surcharge')"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) TNGSTSc,"
  
  Sql = Sql + " isNull((Select (InvoiceQty*Round(Rate,2))  From SalesInvoice Where Type = 'Labour' and a.invoicetype = invoicetype "
  Sql = Sql + " and a.invoiceno = invoiceno and a.invoicedate = invoicedate and ItemCode = a.ItemCode and CustomerCode = a.CustomerCode  and a.branchid = branchid and a.compid = compid),0) Labor,"
  
  Sql = Sql + " isNull((Select (InvoiceQty*Round(Rate,2))  From SalesInvoice Where Type = 'EOU' and a.invoicetype = invoicetype "
  Sql = Sql + " and a.invoiceno = invoiceno and a.invoicedate = invoicedate and ItemCode = a.ItemCode and CustomerCode = a.CustomerCode  and a.branchid = branchid and a.compid = compid),0) EOU"
  
  Sql = Sql + " from SalesInvoice a "
  Sql = Sql + " where a.branchid = '" & gBranchID & "' and a.compid = '" & gCompID & "'"
  Sql = Sql + " and a.invoiceDate between '" + Format(dpfromdt, "yyyy-MM-dd") + "' and '" + Format(dptodt, "yyyy-MM-dd") + "'"
  
  If optRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one item"
GoTo ExitHere
End If
    Sql = Sql & " and a.CustomerCode in ("
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
  If CmbList.ListIndex = 1 Then
    Sql = Sql + " where x.TNGST > 0"
  ElseIf CmbList.ListIndex = 2 Then
    Sql = Sql + " where x.CST > 0"
  ElseIf CmbList.ListIndex = 3 Then
    Sql = Sql + " where x.Labor > 0"
  ElseIf CmbList.ListIndex = 4 Then
    Sql = Sql + " where x.EOU > 0"
  End If
  Sql = Sql + " Order by x.InvoiceNo, x.InvoiceDate "
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    sShowMessage "No records exist"
    sPrintSelect = False
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
  gStrReportHeading = "Sales Tax Register between " & dpfromdt & " and " & dptodt
  sSendToPort CAlign(iPageWidth, gStrReportHeading)
  sSendToPort ""
  sPageHeader
  i = 1
  tamt = 0
  Do While Not rsDetails.EOF
    sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(35, rsDetails("CustomerName")) & Space(1) & RAlign(11, Format(rsDetails("TaxableIncome"), "0.00")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
    sSendToPort LAlign(10, rsDetails("InvoiceDate")) & Space(49) & RAlign(10, rsDetails("TNGSTSc")) & Space(1) & RAlign(10, Format(rsDetails("TotalAmount"), "0.00"))
    sSendToPort ""
    GV = GV + Format(rsDetails("TaxableIncome"), "0.00")
    StAmt = StAmt + (rsDetails("TNGST") + rsDetails("CST"))
    suramt = suramt + rsDetails("TNGSTSc")
    tamt = tamt + (rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
    invamt = invamt + rsDetails("TotalAmount")
    i = i + 1
    rsDetails.MoveNext
  Loop
  sSendToPort ""
  sSendToPort LAlign(10, "Total") & Space(37) & RAlign(11, Format(GV, "0.00")) & Space(1) & RAlign(10, Format(StAmt, "0.00")) & Space(1) & RAlign(10, Format(tamt, "0.00"))
  sSendToPort Space(59) & RAlign(10, Format(suramt, "0.00")) & Space(1) & RAlign(10, Format(invamt, "0.00"))
  
  rsDetails.Close
  Call EndofReport(iPageWidth, i - 1)
  
  sPrintSelect = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function
Private Sub cbExit_Click()
  Unload Me
End Sub

Private Sub cmdView_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If sPrintCurrent() = False Then
 GoTo ExitHere
End If
  
  If cmbReportType.ListIndex = 1 Then
 Call FinaliseAndShowPrintPreview
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set frmOutput = Nothing
  Set frmCurrent = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Public Sub sPageHeader()
  
  Try
  Me.MousePointer = vbHourglass

  ''sSendToPort "Print Date: " & Format(ServerDate, "dd/mm/yyyy HH:mm") & RAlign(iPageWidth - 31, "Page No.: " & CStr(iPageNo))
  sSendToPort String(iPageWidth, "-")
  sSendToPort LAlign(10, "Inv. No.") & Space(1) & LAlign(35, "Customer") & Space(1) & RAlign(11, "Taxable Amt") & Space(1) & RAlign(10, "Sales Tax") & Space(1) & RAlign(10, "ST Total")
  sSendToPort LAlign(10, "Inv. Date") & Space(47) & RAlign(12, "VAT %") & Space(1) & RAlign(10, "Inv. Amt")
  sSendToPort String(iPageWidth, "-")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Function sPrintCurrent() As Boolean
  Dim rsDetails As AceADODB.Recordset
  Dim sRow As Long
  Dim tamt As Double
  Dim StAmt As Double
  Dim suramt As Double
  Dim invamt As Double
  Dim GV As Double
  Dim Sql As String
  Dim i As Integer
  Dim tamt2 As Double
  Dim StAmt2 As Double
  Dim suramt2 As Double
  Dim invamt2 As Double
  Dim GV2 As Double
    
  Try
  
  sPrintCurrent = False
  
  Set rsDetails = New AceADODB.Recordset
  
  Sql = " select distinct x.InvoiceNo, x.InvoiceDate, x.CustomerName, x.Labor, Case When x.ED=0 Then 0 Else (x.GV+x.ED+x.CESS) End as TaxableIncome, "
  Sql = Sql + " x.TNGST, x.CST, x.TNGSTSc, x.EOU, x.TotalAmount  from (select a.InvoiceNo, "
  Sql = Sql + " a.InvoiceDate, a.CustomerCode, (select d.PartyName from FAS_Party_Master d"
  Sql = Sql + " where d.PartyCode = a.CustomerCode and d.Branchid = a.Branchid "
  Sql = Sql + " and d.Compid = a.Compid) CustomerName, a.ItemCode, a.InvoiceQty, "
  Sql = Sql + " a.Rate, a.TotalAmount, (a.InvoiceQty * a.Rate) GV, "
  
  Sql = Sql + " isNull((select sum(c.ResultValue) from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from COM_LSTCSTGrouping d where d.TaxDesc = 'ED'"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) ed,"

  Sql = Sql + " isNull((select sum(c.ResultValue) from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from COM_LSTCSTGrouping d where d.TaxDesc = 'CESS'"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) CESS,"

  Sql = Sql + " isNull((select sum(c.ResultValue) from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from COM_LSTCSTGrouping d where d.TaxDesc = 'LST'"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) TNGST,"
  
  Sql = Sql + " isNull((select sum(c.ResultValue) from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from COM_LSTCSTGrouping d where d.TaxDesc = 'CST'"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) CST,"
  
  Sql = Sql + " isNull((select sum(c.ResultValue) from salesCalculationDetails c where "
  Sql = Sql + " a.invoicetype = c.doctype and a.invoiceno = c.docno and a.invoicedate = c.docdate "
  Sql = Sql + " and a.branchid = c.branchid and a.compid = c.compid"
  Sql = Sql + " and c.ExprGroupId in (select d.ExprGroupId from COM_LSTCSTGrouping d where d.TaxDesc in ('CST Surcharge', 'LST Surcharge')"
  Sql = Sql + " and d.branchid = c.branchid and d.compid = c.compid)),0) TNGSTSc,"
  
  Sql = Sql + " isNull((Select (InvoiceQty*Round(Rate,2))  From SalesInvoice Where Type = 'Labour' and a.invoicetype = invoicetype "
  Sql = Sql + " and a.invoiceno = invoiceno and a.invoicedate = invoicedate and ItemCode = a.ItemCode and CustomerCode = a.CustomerCode  and a.branchid = branchid and a.compid = compid),0) Labor,"
  
  Sql = Sql + " isNull((Select (InvoiceQty * Round(Rate, 2)) From SalesInvoice"
  Sql = Sql + " Where Type = 'EOU' and a.invoicetype = invoicetype  and a.invoiceno = invoiceno and a.invoicedate = invoicedate and ItemCode = a.ItemCode"
  Sql = Sql + " and CustomerCode = a.CustomerCode  and a.branchid = branchid and a.compid = compid), 0) EOU"
  
  Sql = Sql + " from SalesInvoice a "
  Sql = Sql + " where a.branchid = '" & gBranchID & "' and a.compid = '" & gCompID & "'"
  Sql = Sql + " and a.invoiceDate between '" + Format(dpfromdt, "yyyy-MM-dd") + "' and '" + Format(dptodt, "yyyy-MM-dd") + "'"
  
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
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql & ")"
  End If
  Sql = Sql & ") x"
  Sql = Sql + " Order by x.InvoiceNo, x.InvoiceDate "
  
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    sShowMessage "No records exist"
    sPrintCurrent = False
    GoTo ExitHere
  End If
  
  If cmbReportType.ListIndex = 1 Then
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
    gStrReportHeading = "Sales Tax Register between " & dpfromdt & " and " & dptodt
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
    i = 1
    
    rsDetails.Filter = adFilterNone
    rsDetails.Filter = " TNGST > '0'"
    
    tamt = 0
    If Not rsDetails.EOF Then sSendToPort CAlign(80, "TNGST / VAT")
sSendToPort CAlign(80, "-----------")
    Do While Not rsDetails.EOF
      sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(35, rsDetails("CustomerName")) & Space(1) & RAlign(11, Format(rsDetails("TaxableIncome"), "0.00")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
      sSendToPort LAlign(10, rsDetails("InvoiceDate")) & Space(49) & RAlign(10, rsDetails("TNGSTSc")) & Space(1) & RAlign(10, Format(rsDetails("TotalAmount"), "0.00"))
      sSendToPort ""
      GV = GV + Format(rsDetails("TaxableIncome"), "0.00")
      StAmt = StAmt + (rsDetails("TNGST") + rsDetails("CST"))
      suramt = suramt + rsDetails("TNGSTSc")
      tamt = tamt + (rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
      invamt = invamt + rsDetails("TotalAmount")
      i = i + 1
      rsDetails.MoveNext
    Loop
    
    If (GV + StAmt + suramt + tamt + invamt) > 0 Then
      sSendToPort ""
      sSendToPort LAlign(10, "Total") & Space(37) & RAlign(11, Format(GV, "0.00")) & Space(1) & RAlign(10, Format(StAmt, "0.00")) & Space(1) & RAlign(10, Format(tamt, "0.00"))
      sSendToPort Space(59) & RAlign(10, Format(suramt, "0.00")) & Space(1) & RAlign(10, Format(invamt, "0.00"))
    End If
    
    GV2 = GV
    StAmt2 = StAmt
    suramt2 = suramt
    tamt2 = tamt
    invamt2 = invamt
    
    GV = 0
    StAmt = 0
    suramt = 0
    tamt = 0
    invamt = 0
    
    rsDetails.Filter = adFilterNone
    rsDetails.Filter = " CST > '0'"
    tamt = 0
    If Not rsDetails.EOF Then sSendToPort CAlign(80, "CST")
sSendToPort CAlign(80, "---")
    Do While Not rsDetails.EOF
      sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(35, rsDetails("CustomerName")) & Space(1) & RAlign(11, Format(rsDetails("TaxableIncome"), "0.00")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
      sSendToPort LAlign(10, rsDetails("InvoiceDate")) & Space(49) & RAlign(10, rsDetails("TNGSTSc")) & Space(1) & RAlign(10, Format(rsDetails("TotalAmount"), "0.00"))
      sSendToPort ""
      GV = GV + Format(rsDetails("TaxableIncome"), "0.00")
      StAmt = StAmt + (rsDetails("TNGST") + rsDetails("CST"))
      suramt = suramt + rsDetails("TNGSTSc")
      tamt = tamt + (rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
      invamt = invamt + rsDetails("TotalAmount")
      i = i + 1
      rsDetails.MoveNext
    Loop
    
    If (GV + StAmt + suramt + tamt + invamt) > 0 Then
      sSendToPort ""
      sSendToPort LAlign(10, "Total") & Space(37) & RAlign(11, Format(GV, "0.00")) & Space(1) & RAlign(10, Format(StAmt, "0.00")) & Space(1) & RAlign(10, Format(tamt, "0.00"))
      sSendToPort Space(59) & RAlign(10, Format(suramt, "0.00")) & Space(1) & RAlign(10, Format(invamt, "0.00"))
    End If
    
    GV2 = GV2 + GV
    StAmt2 = StAmt2 + StAmt
    suramt2 = suramt2 + suramt
    tamt2 = tamt2 + tamt
    invamt2 = invamt2 + invamt
    
    GV = 0
    StAmt = 0
    suramt = 0
    tamt = 0
    invamt = 0
    
    rsDetails.Filter = adFilterNone
    rsDetails.Filter = " Labor > '0'"
    tamt = 0
    If Not rsDetails.EOF Then sSendToPort CAlign(80, "Labor")
sSendToPort CAlign(80, "-----")
    Do While Not rsDetails.EOF
      sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(35, rsDetails("CustomerName")) & Space(1) & RAlign(11, Format(rsDetails("TaxableIncome"), "0.00")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
      sSendToPort LAlign(10, rsDetails("InvoiceDate")) & Space(49) & RAlign(10, rsDetails("TNGSTSc")) & Space(1) & RAlign(10, Format(rsDetails("TotalAmount"), "0.00"))
      sSendToPort ""
      GV = GV + Format(rsDetails("TaxableIncome"), "0.00")
      StAmt = StAmt + (rsDetails("TNGST") + rsDetails("CST"))
      suramt = suramt + rsDetails("TNGSTSc")
      tamt = tamt + (rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
      invamt = invamt + rsDetails("TotalAmount")
      i = i + 1
      rsDetails.MoveNext
    Loop
    If (GV + StAmt + suramt + tamt + invamt) > 0 Then
      sSendToPort ""
      sSendToPort LAlign(10, "Total") & Space(37) & RAlign(11, Format(GV, "0.00")) & Space(1) & RAlign(10, Format(StAmt, "0.00")) & Space(1) & RAlign(10, Format(tamt, "0.00"))
      sSendToPort Space(59) & RAlign(10, Format(suramt, "0.00")) & Space(1) & RAlign(10, Format(invamt, "0.00"))
    End If
    
    GV2 = GV2 + GV
    StAmt2 = StAmt2 + StAmt
    suramt2 = suramt2 + suramt
    tamt2 = tamt2 + tamt
    invamt2 = invamt2 + invamt
    
    GV = 0
    StAmt = 0
    suramt = 0
    tamt = 0
    invamt = 0
    
    rsDetails.Filter = adFilterNone
    rsDetails.Filter = " EOU > '0'"
    tamt = 0
    If Not rsDetails.EOF Then sSendToPort CAlign(80, "EOU")
sSendToPort CAlign(80, "-----")
    Do While Not rsDetails.EOF
      sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(35, rsDetails("CustomerName")) & Space(1) & RAlign(11, Format(rsDetails("TaxableIncome"), "0.00")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST")) & Space(1) & RAlign(10, rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
      sSendToPort LAlign(10, rsDetails("InvoiceDate")) & Space(49) & RAlign(10, rsDetails("TNGSTSc")) & Space(1) & RAlign(10, Format(rsDetails("TotalAmount"), "0.00"))
      sSendToPort ""
      GV = GV + Format(rsDetails("TaxableIncome"), "0.00")
      StAmt = StAmt + (rsDetails("TNGST") + rsDetails("CST"))
      suramt = suramt + rsDetails("TNGSTSc")
      tamt = tamt + (rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
      invamt = invamt + rsDetails("TotalAmount")
      i = i + 1
      rsDetails.MoveNext
    Loop
    If (GV + StAmt + suramt + tamt + invamt) > 0 Then
      sSendToPort ""
      sSendToPort LAlign(10, "Total") & Space(37) & RAlign(11, Format(GV, "0.00")) & Space(1) & RAlign(10, Format(StAmt, "0.00")) & Space(1) & RAlign(10, Format(tamt, "0.00"))
      sSendToPort Space(59) & RAlign(10, Format(suramt, "0.00")) & Space(1) & RAlign(10, Format(invamt, "0.00"))
    End If
    
    GV2 = GV2 + GV
    StAmt2 = StAmt2 + StAmt
    suramt2 = suramt2 + suramt
    tamt2 = tamt2 + tamt
    invamt2 = invamt2 + invamt
    
    sSendToPort String(iPageWidth, "-")
    
    sSendToPort ""
    sSendToPort LAlign(15, "GRAND TOTAL") & Space(32) & RAlign(11, Format(GV2, "0.00")) & Space(1) & RAlign(10, Format(StAmt2, "0.00")) & Space(1) & RAlign(10, Format(tamt2, "0.00"))
    sSendToPort Space(59) & RAlign(10, Format(suramt2, "0.00")) & Space(1) & RAlign(10, Format(invamt2, "0.00"))
    
    rsDetails.Close
    Call EndofReport(iPageWidth, i - 1)
  Else
      frmPrint.Show
  
      j = 1
GridCount = 1
k = 0
iPageNo = 1
pheight = 54
Orient = 1
      
      Header
      
      j = j + 1
GridCount = GridCount + 1
      tamt = 0
    
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      
  'TNGST/VAT--------------------
      rsDetails.Filter = adFilterNone
      rsDetails.Filter = " TNGST > '0'"
      
      If rsDetails.RecordCount > 0 Then
        frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "TNGST / VAT"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
      Else
        GoTo rec1
      End If
      Do While Not rsDetails.EOF
        frmPrint.sprPrint.SetText 1, j, rsDetails("InvoiceNo")
        frmPrint.sprPrint.SetText 2, j, rsDetails("InvoiceDate")
        frmPrint.sprPrint.SetText 3, j, rsDetails("CustomerName")
        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 4, j, Format(rsDetails("TaxableIncome"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 5, j, Val(rsDetails("TNGST") + rsDetails("CST"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 6, j, rsDetails("TNGSTSc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 7, j, Val(rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 8, j, Format(rsDetails("TotalAmount"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        GV = GV + Format(rsDetails("TaxableIncome"), "0.00")
        StAmt = StAmt + (rsDetails("TNGST") + rsDetails("CST"))
        suramt = suramt + rsDetails("TNGSTSc")
        tamt = tamt + (rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
        invamt = invamt + rsDetails("TotalAmount")
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j

        j = j + 1
GridCount = GridCount + 1
RowProcess
        rsDetails.MoveNext
      Loop
      
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Sub Total : "
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 4, j, Format(GV, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 5, j, Format(StAmt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 6, j, Format(suramt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 7, j, Format(tamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 8, j, Format(invamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

'      .Col = 8: .Row = j: .BackColor = RGB(255, 255, 0)

      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      GV2 = GV2 + GV
      StAmt2 = StAmt2 + StAmt
      suramt2 = suramt2 + suramt
      tamt2 = tamt2 + tamt
      invamt2 = invamt2 + invamt
      
      GV = 0
      StAmt = 0
      suramt = 0
      tamt = 0
      invamt = 0
rec1:
  'CST--------------------
      rsDetails.Filter = adFilterNone
      rsDetails.Filter = " CST > '0'"
      
      If rsDetails.RecordCount > 0 Then
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        
        frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "CST"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
        
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
      Else
        GoTo Rec2
      End If
      
      Do While Not rsDetails.EOF
        frmPrint.sprPrint.SetText 1, j, rsDetails("InvoiceNo")
        frmPrint.sprPrint.SetText 2, j, rsDetails("InvoiceDate")
        frmPrint.sprPrint.SetText 3, j, rsDetails("CustomerName")
        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 4, j, Format(rsDetails("TaxableIncome"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 5, j, Val(rsDetails("TNGST") + rsDetails("CST"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 6, j, rsDetails("TNGSTSc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 7, j, Val(rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 8, j, Format(rsDetails("TotalAmount"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        GV = GV + Format(rsDetails("TaxableIncome"), "0.00")
        StAmt = StAmt + (rsDetails("TNGST") + rsDetails("CST"))
        suramt = suramt + rsDetails("TNGSTSc")
        tamt = tamt + (rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
        invamt = invamt + rsDetails("TotalAmount")
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        rsDetails.MoveNext
      Loop
      
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Sub Total : "
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 4, j, Format(GV, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 5, j, Format(StAmt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 6, j, Format(suramt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 7, j, Format(tamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 8, j, Format(invamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)

      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      GV2 = GV2 + GV
      StAmt2 = StAmt2 + StAmt
      suramt2 = suramt2 + suramt
      tamt2 = tamt2 + tamt
      invamt2 = invamt2 + invamt
      
      GV = 0
      StAmt = 0
      suramt = 0
      tamt = 0
      invamt = 0
Rec2:
  'LABOR--------------------
      rsDetails.Filter = adFilterNone
      rsDetails.Filter = " Labor > '0'"
      
      If rsDetails.RecordCount > 0 Then
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "LABOR"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
      Else
        GoTo rec3
      End If
      
      Do While Not rsDetails.EOF
        frmPrint.sprPrint.SetText 1, j, rsDetails("InvoiceNo")
        frmPrint.sprPrint.SetText 2, j, rsDetails("InvoiceDate")
        frmPrint.sprPrint.SetText 3, j, rsDetails("CustomerName")
        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 4, j, Format(rsDetails("TaxableIncome"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 5, j, Val(rsDetails("TNGST") + rsDetails("CST"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 6, j, rsDetails("TNGSTSc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 7, j, Val(rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 8, j, Format(rsDetails("TotalAmount"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        GV = GV + Format(rsDetails("TaxableIncome"), "0.00")
        StAmt = StAmt + (rsDetails("TNGST") + rsDetails("CST"))
        suramt = suramt + rsDetails("TNGSTSc")
        tamt = tamt + (rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
        invamt = invamt + rsDetails("TotalAmount")
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        rsDetails.MoveNext
      Loop
      
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Sub Total : "
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 4, j, Format(GV, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 5, j, Format(StAmt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 6, j, Format(suramt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 7, j, Format(tamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 8, j, Format(invamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)

      j = j + 1
GridCount = GridCount + 1
RowProcess
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      GV2 = GV2 + GV
      StAmt2 = StAmt2 + StAmt
      suramt2 = suramt2 + suramt
      tamt2 = tamt2 + tamt
      invamt2 = invamt2 + invamt
      
      GV = 0
      StAmt = 0
      suramt = 0
      tamt = 0
      invamt = 0
rec3:
   'EOU--------------------
      rsDetails.Filter = adFilterNone
      rsDetails.Filter = " EOU > '0'"
      If rsDetails.RecordCount > 0 Then
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j
        frmPrint.sprPrint.AddCellSpan 1, j, frmPrint.sprPrint.MaxCols, 1
        frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 1, j, "EOU"
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 9
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
      Else
        GoTo Rec4
      End If

      Do While Not rsDetails.EOF
        frmPrint.sprPrint.SetText 1, j, rsDetails("InvoiceNo")
        frmPrint.sprPrint.SetText 2, j, rsDetails("InvoiceDate")
        frmPrint.sprPrint.SetText 3, j, rsDetails("CustomerName")

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 4, j, Format(rsDetails("TaxableIncome"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 5, j, Val(rsDetails("TNGST") + rsDetails("CST"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.SetText 6, j, rsDetails("TNGSTSc")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 8, j, Val(rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText 8, j, Format(rsDetails("TotalAmount"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        GV = GV + Format(rsDetails("TaxableIncome"), "0.00")
        StAmt = StAmt + (rsDetails("TNGST") + rsDetails("CST"))
        suramt = suramt + rsDetails("TNGSTSc")
        tamt = tamt + (rsDetails("TNGST") + rsDetails("CST") + rsDetails("TNGSTSc"))
        invamt = invamt + rsDetails("TotalAmount")
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        rsDetails.MoveNext
      Loop
      
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
     
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Sub Total : "
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 4, j, Format(GV, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 5, j, Format(StAmt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 6, j, Format(suramt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 7, j, Format(tamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 8, j, Format(invamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)

      j = j + 1
GridCount = GridCount + 1
RowProcess
      GV2 = GV2 + GV
      StAmt2 = StAmt2 + StAmt
      suramt2 = suramt2 + suramt
      tamt2 = tamt2 + tamt
      invamt2 = invamt2 + invamt
Rec4:
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Grant Total : "
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 4, j, Format(GV2, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 5, j, Format(StAmt2, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 6, j, Format(suramt2, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 7, j, Format(tamt2, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
      frmPrint.sprPrint.SetText 8, j, Format(invamt2, "0.00")
      frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.Row = j
frmPrint.sprPrint.BackColor = RGB(255, 255, 0)
      
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
    totalpages = iPageNo
  End If
  
  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Sub cmdClear_Click()
' On Error Resume Next
  
  Call StepClear(Me)
  StepEnable Me
  spCust.Visible = False
  OptAll.Value = True
  CmdStep(1).Enabled = False
  cmbReportType.ListIndex = 0
  CmbList.ListIndex = 0
'  MDISales.Toolbar1.Buttons(1).Enabled = True
'  MDISales.Toolbar1.Buttons(2).Enabled = True
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If fraStep.UBound = Index Then
    If optRandom = True Then
      If spCust.SelectionCount >= 1 Then
        spCust.Enabled = False
        CmdStep(Index).Enabled = False
      End If
    End If
  End If
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  Try
  Me.MousePointer = vbHourglass

  CurrDate Me
  cmbReportType.ListIndex = 0
  CmbList.ListIndex = 0
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  Else
'    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  CmdStep(1).Enabled = True
  spCust.Visible = False
End Sub

Private Sub optRandom_Click()

  Dim Sql As String
  Dim rs As AceADODB.Recordset
 
  Try
  Me.MousePointer = vbHourglass
  
  CmdStep(1).Enabled = False
  
  Sql = " select distinct a.CustomerCode [Customer Code],(select b.PartyName from FAS_PARTY_MASTER b"
  Sql = Sql + " where b.PartyCode=a.CustomerCode and b.Branchid=a.Branchid and b.Compid=a.Compid) [Customer Name]"
  Sql = Sql + " from SAL_Invoice_Hdr a where a.Branchid='" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Sql = Sql + " and InvoiceDate between convert(datetime,'" + Format(dpfromdt, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " and convert(datetime,'" + Format(dptodt, "yyyy-MM-dd") + "',111)"
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spCust.ClearSelection
'  spCust.MaxRows = spCust.DataRowCnt
'  If spCust.DataRowCnt = 0 Then
'    sShowMessage "No Records Exists."
'    OptAll.SetFocus
'    GoTo ExitHere
'  End If
'  spCust.Col = 1: spCust.ColHidden = True
'  spCust.Enabled = True
'  spCust.Visible = True
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    OptAll.Value = True
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

  Set rs = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spCust_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  If spCust.SelectionCount >= 1 Then
    CmdStep(1).Enabled = True
  Else
    CmdStep(1).Enabled = False
  End If
End Sub

Private Sub Design_spCust(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spCust


spCust.Click +=  new EventHandler(spCust_Click)

End Sub

Private Sub InitializeSpreads()

	Call Design_spCust("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesTaxRegisterReport.frm", spCust, "")

End Sub
