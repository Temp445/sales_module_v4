VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptExciseDutyRegisterReport 
   Caption         =   "Excise Duty Register Report"
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
      Left            =   7410
      TabIndex        =   5
      Top             =   2925
      Width           =   345
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
      Left            =   7410
      TabIndex        =   2
      Top             =   2235
      Width           =   345
   End
   Begin VB.Frame frastep 
      Caption         =   "Customer"
      Enabled         =   0   'False
      Height          =   555
      Index           =   1
      Left            =   4815
      TabIndex        =   16
      Top             =   2805
      Width           =   2520
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   135
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   570
      End
      Begin VB.OptionButton optRandom 
         Caption         =   "Selective"
         Height          =   225
         Left            =   1455
         TabIndex        =   4
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame frastep 
      Height          =   765
      Index           =   0
      Left            =   4815
      TabIndex        =   13
      Top             =   2040
      Width           =   2520
      Begin MSComCtl2.DTPicker dptodt 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   109379585
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dpfromdt 
         Height          =   315
         Left            =   75
         TabIndex        =   0
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   109379585
         CurrentDate     =   37273
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         Height          =   210
         Left            =   90
         TabIndex        =   15
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To Date"
         Height          =   210
         Left            =   1245
         TabIndex        =   14
         Top             =   180
         Width           =   555
      End
   End
   Begin VB.Frame FraReportType 
      Caption         =   "Report Type"
      Height          =   945
      Left            =   4815
      TabIndex        =   12
      Top             =   3360
      Width           =   2520
      Begin VB.ComboBox cmbReportType 
         Height          =   330
         ItemData        =   "ExciseDutyRegisterReport.frx":0000
         Left            =   90
         List            =   "ExciseDutyRegisterReport.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   210
         Width           =   2340
      End
      Begin VB.CheckBox chkPrintdate 
         Caption         =   "Print Date"
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   615
         Value           =   1  'Checked
         Width           =   1080
      End
   End
   Begin VB.CommandButton cbExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6795
      TabIndex        =   10
      Top             =   4440
      Width           =   1300
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5505
      TabIndex        =   9
      Top             =   4440
      Width           =   1300
   End
Begin AceSpread spCust
      Height          =   1305
      Left            =   7815
      TabIndex        =   11
      Top             =   2940
      Visible         =   0   'False
      Width           =   3900
      _Version        =   458752
      _ExtentX        =   6879
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
   Begin VB.CommandButton CmdView 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4215
      TabIndex        =   8
      Top             =   4440
      Width           =   1300
   End
End
Attribute VB_Name = "rptExciseDutyRegisterReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim k As Integer
Dim j As Long

Public Sub sprint()
  'No Change
End Sub
Public Sub sView()
  'No Change
End Sub
Private Sub Header()
  Dim mstr As String
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
    frmPrint.sprPrint.MaxCols = 7
    
    frmPrint.sprPrint.AddCellSpan 1, j, 2, 1
    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(iPageNo)
    
    frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
    frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
    If chkPrintDate.Value = 1 Then
 frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(15) & "Printed Date : " & Format(ServerDate, "dd/MM/yyyy HH:MM:ss")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
End If
    j = j + 1
GridCount = GridCount + 1
    
    Sql = " Select * from dbo.fn_SAL_Menu_List('" & Trim$(mnuFlg) & "')"
    Set rs = New Recordset
    If RSOpen(rs, Sql) = False Then GoTo ExitHere
    If rs.RecordCount > 0 Then
      frmPrint.sprPrint.AddCellSpan 1, j, 3, 1
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "ERP Ref. No   :" & Space(1) & rs("ErpNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j
      frmPrint.sprPrint.AddCellSpan frmPrint.sprPrint.Col, j, frmPrint.sprPrint.MaxCols, 1
      frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Space(15) & "TS Ref. No   : " & rs("TsNo")
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      j = j + 1
GridCount = GridCount + 1
    End If
    
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
    frmPrint.sprPrint.SetText 1, j, "Sales Excise Duty Register between " & dpfromdt & " and " & dptodt
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Invoice No"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Customer Name"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 27
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Goods"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "ED Value"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Cess"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sub Total"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Invoice"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
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

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Invoice Date"
 frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Value"
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Sur.Cess"
frmPrint.sprPrint.ColWidth(frmPrint.sprPrint.Col) = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
    
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

    frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, "Amount"
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

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdView_Click()
  
  Try
  Me.MousePointer = vbHourglass

'  Set frmCurrent = Me
'  DoEvents
'  iLineNo = 1
'  iPageNo = 1

  If sPrintCurrent() = False Then
 GoTo ExitHere
End If
  
  If cmbReportType.ListIndex = 1 Then
 Call FinaliseAndShowPrintPreview
End If
  
'  sClosePort
'  Set frmOutput = New frmReportViewer
'  frmOutput.Caption = gStrReportHeading
'  frmOutput.txtTotalPages.Text = CStr(iPageNo)
'  frmOutput.wbrView.Navigate strFileName
'  frmOutput.Tag = strFileName
'  gStrReportHeading = ""
'  frmOutput.Show vbModal
'  DoEvents
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
  sSendToPort LAlign(10, "Inv. No.") & Space(1) & LAlign(35, "Customer") & Space(1) & RAlign(11, "Goods Value") & Space(1) & RAlign(10, "Cess") & Space(1) & RAlign(10, "Total")
  sSendToPort LAlign(10, "Inv. Date") & Space(37) & RAlign(11, "ED Value") & Space(1) & RAlign(10, "Sh.Cess") & Space(1) & RAlign(10, "Inv. Amt")
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
  Dim EDAmt As Double
  Dim CessAmt As Double
  Dim SheAmt As Double
  Dim invamt As Double
  Dim GV As Double
  Dim Sql As String
  Dim i As Integer

  Try
  
  Set rsDetails = New AceADODB.Recordset
  
  Sql = " select x.InvoiceType,x.InvoiceNo, x.InvoiceDate, x.CustomerName, x.GV, "
  Sql = Sql + " x.ED, x.Cess2,x.SHE1, x.TotalAmount  from (select a.InvoiceType,a.InvoiceNo, "
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
  Sql = Sql + " and d.Compid=c.Compid)),0) CESS2,"
  
  Sql = Sql + " isNull((select c.ResultValue from salesCalculationDetails c where a.invoicetype = c.doctype"
  Sql = Sql + " and a.invoiceno = c.docno and a.invoicedate = c.docdate  and a.branchid = c.branchid"
  Sql = Sql + " and a.compid = c.compid and a.finyear = c.finyear"
  Sql = Sql + " and c.ExprGroupId in"
  Sql = Sql + " (select ExprGroupId from SAL_LSTCSTGrouping d where [Description] ='SHE' and d.Branchid=c.Branchid"
  Sql = Sql + " and d.Compid=c.Compid)),0) SHE1,"

  Sql = Sql + " a.TotalAmount from SalesInvoice a where a.branchid = '" & gBranchID & "'"
  Sql = Sql + " and a.compid = '" & gCompID & "' and a.finyear = '" & gFinyear & "'"
  Sql = Sql + " and a.invoiceDate between '" + Format(dpfromdt, "yyyy-MM-dd") + "'"
  Sql = Sql + " and '" + Format(dptodt, "yyyy-MM-dd") + "'"
  If optRandom Then
    If spCust.SelectionCount < 1 Then
 sShowMessage "Please select atlease one Vendor"
GoTo ExitHere
End If
    Sql = Sql & " and a.CustomerCode in"
    Sql = Sql & Spread_Get_SelectionList(spCust, 1)
      sRow = 0
      For i = 1 To spCust.SelectionCount
        sRow = spCust.GetMultiSelItem(sRow)
        spCust.Row = sRow
spCust.Col = 1
        spCust.Tag = spCust.Tag & "'" & Trim$(spCust.Text) & "',"
      Next i
      spCust.Tag = Left$(spCust.Tag, Len(spCust.Tag) - 1)  End If
  Sql = Sql & ") x"
  Sql = Sql + " where (x.ED > 0) order by x.InvoiceNo"
  If RSOpen(rsDetails, Sql) = False Then GoTo ExitHere
  If rsDetails.RecordCount = 0 Then
    MsgBox "No records exist"
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
    gStrReportHeading = "Sales Excise Duty Register between " & dpfromdt & " and " & dptodt
    sSendToPort CAlign(iPageWidth, gStrReportHeading)
    sSendToPort ""
    sPageHeader
   tamt = 0
  Do While Not rsDetails.EOF
    sSendToPort LAlign(10, rsDetails("InvoiceNo")) & Space(1) & LAlign(35, rsDetails("CustomerName")) & Space(1) & RAlign(11, Format(rsDetails("GV"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("Cess2"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("ED") + rsDetails("Cess2") + rsDetails("SHE1"), "0.00"))
    sSendToPort LAlign(10, rsDetails("InvoiceDate")) & Space(38) & RAlign(10, Format(rsDetails("ED"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("SHE1"), "0.00")) & Space(1) & RAlign(10, Format(rsDetails("TotalAmount"), "0.00"))
    sSendToPort ""
    GV = GV + Format(rsDetails("GV"), "0.00")
    EDAmt = EDAmt + Format(rsDetails("ed"), "0.00")
    CessAmt = CessAmt + Format(rsDetails("cess2"), "0.00")
    SheAmt = SheAmt + Format(rsDetails("SHE1"), "0.00")
    tamt = tamt + Format(rsDetails("ed") + rsDetails("cess2") + rsDetails("SHE1"), "0.00")
    invamt = invamt + Format(rsDetails("TotalAmount"), "0.00")
    i = i + 1
    rsDetails.MoveNext
  Loop
  sSendToPort ""
  sSendToPort LAlign(10, "Total") & Space(37) & RAlign(11, Format(GV, "0.00")) & Space(1) & RAlign(10, Format(CessAmt, "0.00")) & Space(1) & RAlign(10, Format(tamt, "0.00"))
  sSendToPort Space(48) & RAlign(10, Format(EDAmt, "0.00")) & Space(1) & RAlign(10, Format(SheAmt, "0.00")) & Space(1) & RAlign(10, Format(invamt, "0.00"))
    
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
      Do While Not rsDetails.EOF
        If GridCount = 53 Then
          Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        End If
        frmPrint.sprPrint.SetText 1, j, rsDetails("InvoiceType") & rsDetails("InvoiceNo")
        frmPrint.sprPrint.SetText 2, j, rsDetails("CustomerName")
        frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsDetails("GV"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsDetails("ED"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsDetails("Cess2"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsDetails("ED") + rsDetails("Cess2"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsDetails("TotalAmount"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
        
        GV = GV + Format(rsDetails("GV"), "0.00")
        EDAmt = EDAmt + Format(rsDetails("ed"), "0.00")
        CessAmt = CessAmt + Format(rsDetails("cess2"), "0.00")
        SheAmt = SheAmt + Format(rsDetails("SHE1"), "0.00")
        tamt = tamt + Format(rsDetails("ed") + rsDetails("cess2") + rsDetails("SHE1"), "0.00")
        invamt = invamt + Format(rsDetails("TotalAmount"), "0.00")
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        
        frmPrint.sprPrint.SetText 1, j, rsDetails("InvoiceDate")
        
        frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

        frmPrint.sprPrint.SetText frmPrint.sprPrint.Col, j, Format(rsDetails("SHE1"), "0.00")
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
        Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
        j = j + 1
GridCount = GridCount + 1
RowProcess
        rsDetails.MoveNext
      Loop
      
      frmPrint.sprPrint.AddCellSpan 1, j, 2, 2
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexTop, 1, j, frmPrint.sprPrint.MaxCols, j

      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 1, j, "Grant Total : "
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignTop
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 3, j, Format(GV, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 4, j, Format(EDAmt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 5, j, Format(CessAmt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 6, j, Format(tamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 7, j, Format(invamt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight

      j = j + 1
GridCount = GridCount + 1
RowProcess
      
      frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.Row = j

      frmPrint.sprPrint.SetText 5, j, Format(SheAmt, "0.00")
frmPrint.sprPrint.Font = "Arial"
frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
      
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexLeft, 1, j, frmPrint.sprPrint.MaxCols, j
      Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexBottom, 1, j, frmPrint.sprPrint.MaxCols, j
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
  End If
  totalpages = iPageNo
  sPrintCurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDetails = Nothing
End Try
End Function

Private Sub cmdClear_Click()
' On Error Resume Next
  CmdStep(1).Enabled = False
  spCust.Visible = False
  OptAll.Value = True
  Call Spread_Clear(spCust)
  Call StepClear(Me)
  Call StepEnable(Me)
  chkPrintDate.Value = 1
  cmbReportType.ListIndex = 0
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next
  Call StepProceedNext(Me, Index)
  If fraStep.UBound = Index Then
    CmdStep(Index).Enabled = False
    cmdview.Enabled = True
  End If
  spCust.Enabled = False
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  Try
  Me.MousePointer = vbHourglass

  CurrDate Me
  cmbReportType.ListIndex = 0
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
    RptFlg = ""
    MDISales.Toolbar1.Visible = False
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
'    sShowMessage "No Records Exists"
'    OptAll.SetFocus
'    GoTo ExitHere
'  End If
'  spCust.Col = 1: spCust.ColHidden = True
'  spCust.Enabled = True
'  spCust.Visible = True
'  spCust.SetFocus
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

Private Sub spCust_KeyUp(sender as object, e as KeyUpEventArgs)
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
spCust.KeyUp +=  new EventHandler(spCust_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_spCust("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\ExciseDutyRegisterReport.frm", spCust, "")

End Sub
