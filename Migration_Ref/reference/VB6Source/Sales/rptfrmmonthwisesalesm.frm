VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form rptfrmmonthwisesalesm 
   Caption         =   "MonthwiseDespatch-UCAL MM Nager"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Full View"
      Height          =   450
      Left            =   5835
      TabIndex        =   12
      Top             =   4140
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sal Pln Vs Req Rm"
      Height          =   450
      Left            =   4320
      TabIndex        =   11
      Top             =   4140
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      Caption         =   "Vendor"
      Height          =   600
      Index           =   1
      Left            =   4410
      TabIndex        =   7
      Top             =   3165
      Width           =   2490
      Begin VB.OptionButton optAll 
         Caption         =   "All"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         Top             =   255
         Value           =   -1  'True
         Width           =   570
      End
      Begin VB.OptionButton optRandom 
         Caption         =   "Random"
         Height          =   195
         Left            =   1410
         TabIndex        =   8
         Top             =   255
         Width           =   900
      End
   End
   Begin VB.ComboBox cmbDailyReNo 
      Height          =   315
      Left            =   5655
      TabIndex        =   5
      Top             =   2370
      Width           =   870
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   5640
      TabIndex        =   4
      Top             =   3750
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Index           =   0
      Left            =   4440
      TabIndex        =   1
      Top             =   2610
      Width           =   2445
      Begin MSComCtl2.DTPicker dtpProcessMonth 
         Height          =   315
         Left            =   1065
         TabIndex        =   2
         Top             =   180
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM"
         Format          =   106692611
         CurrentDate     =   37991
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Year && Month"
         Height          =   210
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   400
      Left            =   4320
      TabIndex        =   0
      Top             =   3750
      Width           =   1320
   End
Begin AceSpread spList
      Height          =   1530
      Left            =   6990
      TabIndex        =   10
      Top             =   2655
      Visible         =   0   'False
      Width           =   4695
      _Version        =   458752
      _ExtentX        =   8281
      _ExtentY        =   2699
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
   Begin VB.Label Label2 
      Caption         =   "Revision No"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   2385
      Width           =   1215
   End
End
Attribute VB_Name = "rptfrmmonthwisesalesm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim j As Long
Dim FDATE As Date
Dim TDATE As Date
Dim Head As Long
Public pageno As Long

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdView_Click()
  Dim rec As New AceADODB.Recordset
  Dim rs As New AceADODB.Recordset
  Dim ed As New AceADODB.Recordset
  Dim rec1 As AceADODB.Recordset
  Dim Sql As String
  Dim k As Long
  Dim i As Long
  Dim WH As Integer
  Dim tmp As Variant
  Dim PlanQty As Long
  Dim invoiceqty As Long
  Dim balqty As Long
  Dim tbalqty As Long
  Dim tbalamt As Long
  Dim percomp As Long
  Dim plnamount As Long
  Dim invoiceamount As Long
  Dim totbalamt As Long
  Try
  Me.MousePointer = vbHourglass
  
  If optRandom Then
    If spList.SelectionCount < 1 Then
      sShowMessage "Please Select atleast one Customer"
      sPrintCurrent = False
      GoTo ExitHere
    End If
  End If
  
  FDATE = Format(dtpProcessMonth, "yyyy-MM-01")
  TDATE = DateAdd("d", -1, DateAdd("m", 1, FDATE))
  
  j = 1
  GridCount = 1
  pageno = 1
  
'    Sql = " select distinct a.partno,b.itemcode,c.planqty,a.rate,sum(b.invoiceqty)invoiceqty,(planqty-sum(b.invoiceqty))BalanceQty,(sum(b.invoiceqty)/planqty*100)percomplete,(c.planqty*a.rate)PlanAmt,(sum(b.invoiceqty)*a.rate)InvAmt,b.customercode  from"
'    Sql = Sql + " SAL_CUSTOMER_ITEMMASTER a left outer join salesinvoice b on a.itemcode=b.itemcode"
'    Sql = Sql + " and a.customercode=b.customercode left outer join SAL_SalesPlan c on a.itemcode=c.itemcode and a.customercode=c.customercode"
'    Sql = Sql + " where  c.planqty<>'0' and b.Type in ('OE', 'Labour') and month(b.invoiceDate) = " & Month(dtpProcessMonth) & ""
'    Sql = Sql + " and year(b.invoiceDate) = " & Year(dtpProcessMonth) & "  and month(c.planperiod) = " & Month(dtpProcessMonth) & " and year(c.planperiod) = " & Year(dtpProcessMonth) & " "
'    Sql = Sql + " and c.planfor='2' and c.RevisionNo = " & cmbDailyReNo.Text & " "
'
'  If optRandom Then
'    If spList.SelectionCount < 1 Then
'      sShowMessage "Please select atleast one item"
'      sPrintCurrent = False
'      GoTo ExitHere
'    End If
'    Sql = Sql & " and b.customerCode in("
'    With spList
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
'  'Sql = Sql + " GROUP BY a.itemcode order by a.itemcode "
'  Sql = Sql + "group by b.customercode ,a.partno,b.itemcode,c.planqty,a.rate order  by b.itemcode "
      
  Sql = " Select distinct a.PartNo, b.ItemCode, Sum(c.PlanQty) as PlanQty, a.Rate, Sum(b.InvoiceQty) as InvoiceQty,"
  Sql = Sql & " (Sum(c.PlanQty)-Sum(b.InvoiceQty)) as BalanceQty,"
  Sql = Sql & " ((Sum(b.InvoiceQty)/Sum(c.PlanQty))*100) as PerComplete, (Sum(c.PlanQty)*a.Rate) as PlanAmt,"
  Sql = Sql & " (Sum(b.InvoiceQty)*a.Rate) as InvAmt"
  Sql = Sql & " From SAL_CUSTOMER_ITEMMASTER a"
  Sql = Sql & " Left Outer Join(Select ItemCode, CustomerCode, Sum(InvoiceQty) as InvoiceQty"
  Sql = Sql & " From SalesInvoice b"
  Sql = Sql & " Where Type in ('OE', 'Labour') and Month(InvoiceDate) = " & Month(dtpProcessMonth.Value) & " and Year(InvoiceDate) = " & Year(dtpProcessMonth.Value) & ""
  Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  
  If optRandom Then
    Sql = Sql & " and CustomerCode in("
      sRow = 0
      For i = 1 To spList.SelectionCount
        sRow = spList.GetMultiSelItem(sRow)
        spList.Row = sRow
spList.Col = 1
        Sql = Sql + "'" + Trim$(spList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  
  Sql = Sql & " Group by ItemCode, CustomerCode)b"
  Sql = Sql & " On a.ItemCode=b.ItemCode and a.CustomerCode=b.CustomerCode"
  Sql = Sql & " Left Outer Join(Select ItemCode, CustomerCode, Sum(Planqty) as Planqty"
  Sql = Sql & " From SAL_SalesPlan"
  Sql = Sql & " Where PlanQty<>'0' and Month(Planperiod) = " & Month(dtpProcessMonth) & " and Year(Planperiod) = " & Year(dtpProcessMonth) & "   "
  Sql = Sql & " and PlanFor=2 and RevisionNo = " & cmbDailyReNo.Text & ""
  Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  
  If optRandom Then
    Sql = Sql & " and customerCode in("
      sRow = 0
      For i = 1 To spList.SelectionCount
        sRow = spList.GetMultiSelItem(sRow)
        spList.Row = sRow
spList.Col = 1
        Sql = Sql + "'" + Trim$(spList.Text) + "',"
      Next i
      Sql = Left$(Sql, Len(Sql) - 1)    Sql = Sql + ")"
  End If
  
  Sql = Sql & " Group by ItemCode, CustomerCode)c"
  Sql = Sql & " On a.ItemCode=c.ItemCode and a.CustomerCode=c.CustomerCode"
  Sql = Sql & " Where a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and IsNull(c.PlanQty,0)<>0 "
  Sql = Sql & " and b.ItemCode is Not Null and c.ItemCode is Not Null"
  Sql = Sql & " Group by a.PartNo, b.ItemCode, a.Rate"
  Sql = Sql & " Order by b.ItemCode"

  If RSOpen(rec, Sql) = False Then Exit Sub
  If rec.EOF Then
 sShowMessage "No Data(s) Found"
Screen.MousePointer = vbNormal
Exit Sub
End If

  frmPrint.Show
  
  Header           'to display the heading
  
  j = j + 1
GridCount = GridCount + 1

  rec.MoveLast
rec.MoveFirst
  k = 0
    For k = 1 To rec.RecordCount
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
      plnamount = plnamount + rec!PlanAmt
      invoiceamount = invoiceamount + rec!invamt
      percomp = ((invoiceamount / plnamount) * 100)
      tbalqty = PlanQty - invoiceqty
     ' tbalamt = plnamount - invoiceamount
     tbalamt = rec!BalanceQty * rec!Rate
     totbalamt = totbalamt + tbalamt
    
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      frmPrint.sprPrint.SetText 3, j, Trim$(CheckNull((rec!PlanQty)))
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      frmPrint.sprPrint.SetText 2, j, Trim$(CheckNull((rec!ItemCode)))
     frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.SetText 4, j, Trim$(CheckNull((rec!Rate)))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

     frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.SetText 1, j, Trim$(CheckNull((rec!PartNo)))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.SetText 5, j, Trim$(CheckNull((rec!invoiceqty)))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.SetText 6, j, Trim$(CheckNull(Round((rec!BalanceQty))))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.SetText 7, j, Trim$(CheckNull(Round((rec!PlanAmt))))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.SetText 8, j, Trim$(CheckNull(Round((rec!invamt))))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.SetText 9, j, Trim$(CheckNull(Round((tbalamt))))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

      frmPrint.sprPrint.SetText 10, j, Trim$(CheckNull(Round((rec!percomplete))))
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
      PlanQty = PlanQty + rec!PlanQty
      invoiceqty = invoiceqty + rec!invoiceqty
      balqty = balqty + rec!BalanceQty

      If OptAll Then
      
        Sql = " select distinct a.invoicedate,sum(a.invoiceqty)invoiceqty from salesinvoice a"
        Sql = Sql + " left outer join SAL_CUSTOMER_ITEMMASTER b on a.itemcode=b.itemcode"
        Sql = Sql + " and a.customercode=b.customercode and a.compid=b.compid and a.branchid=b.branchid "
        Sql = Sql + " where a.Type in ('OE', 'Labour') and b.itemcode = '" + rec("itemcode") + "'" 'and a.customercode = '" & rec("customercode") & "'"
        
        Sql = Sql + " and month(a.invoiceDate) = " & Month(dtpProcessMonth) & " And year(a.invoiceDate) = " & Year(dtpProcessMonth) & ""
        Sql = Sql + " GROUP BY invoicedate"

      Else
      
        Sql = " select distinct a.invoicedate,sum(a.invoiceqty)invoiceqty from salesinvoice a"
        Sql = Sql + " left outer join SAL_CUSTOMER_ITEMMASTER b on a.itemcode=b.itemcode"
        Sql = Sql + " and a.customercode=b.customercode  "
        Sql = Sql + " where b.itemcode = '" + rec("itemcode") + "' and a.Type in ('OE', 'Labour') and "
        Sql = Sql + " month(a.invoiceDate) = " & Month(dtpProcessMonth) & " And year(a.invoiceDate) = " & Year(dtpProcessMonth) & ""
      
        If optRandom Then
          Sql = Sql & " and a.customercode in("
          With spList
            sRow = 0
            For i = 1 To frmPrint.sprPrint.SelectionCount
              sRow = frmPrint.sprPrint.GetMultiSelItem(sRow)
              frmPrint.sprPrint.Row = sRow
frmPrint.sprPrint.Col = 1
              Sql = Sql + "'" + Trim$(frmPrint.sprPrint.Text) + "',"
            Next i
            Sql = Left$(Sql, Len(Sql) - 1)
            sRow = 0
            For i = 1 To spList.SelectionCount
              sRow = spList.GetMultiSelItem(sRow)
              spList.Row = sRow
spList.Col = 1
              Sql = Sql + "'" + Trim$(spList.Text) + "',"
            Next i
            Sql = Left$(Sql, Len(Sql) - 1)
          Sql = Sql + ")"
        End If
        Sql = Sql + " GROUP BY invoicedate"
      End If

      If RSOpen(ed, Sql) = False Then Exit Sub
      If Not ed.EOF Then
        ed.MoveLast
ed.MoveFirst
        For i = 1 To ed.RecordCount
          frmPrint.sprPrint.SetText 11 + Day(ed("invoicedate")) - 1, j, IIf(Left(ed("invoiceqty"), 1) = "F", ed("invoiceqty"), Left(ed("invoiceqty"), ed("invoiceqty")))
.TypeHAlign = TypeHAlignLeft
          ed.MoveNext
        Next i
      End If

      rec.MoveNext

      j = j + 1
GridCount = GridCount + 1
RowProcess
    End With
  Next k

  'frmPrint.sprPrint.SetText 1, j, "": j = j + 1: GridCount = GridCount + 1: RowProcess

  'frmPrint.sprPrint.SetText 1, j, "": j = j + 1: GridCount = GridCount + 1: RowProcess
  Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 2, j, frmPrint.sprPrint.MaxCols, j
  frmPrint.sprPrint.SetText 2, j, "Total:"
  frmPrint.sprPrint.SetText 3, j, PlanQty
  frmPrint.sprPrint.SetText 5, j, invoiceqty
  frmPrint.sprPrint.SetText 6, j, balqty
  frmPrint.sprPrint.SetText 7, j, plnamount
  frmPrint.sprPrint.SetText 8, j, invoiceamount
  frmPrint.sprPrint.SetText 9, j, totbalamt
  frmPrint.sprPrint.SetText 10, j, percomp
j = j + 1
GridCount = GridCount + 1
RowProcess

  j = j + 1
GridCount = GridCount + 1
RowProcess
  frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
  Screen.MousePointer = 1
  totalpages = pageno
  Set rec = Nothing
  Set ed = Nothing
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
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
        pageno = pageno + 1
        Header
        j = j + 1
GridCount = GridCount + 1
  End If
End Sub

Private Sub Header()
Dim    X As Date
Dim  k As Long

  
  Try
  
    frmPrint.sprPrint.SetText 1, j, gCompName
j = j + 2
GridCount = GridCount + 2

    frmPrint.sprPrint.SetText 1, j, "Monthwise Despatch Report for the month of " & MonthName(Month(dtpProcessMonth)) & "-" & Year(dtpProcessMonth)
frmPrint.sprPrint.ColWidth(1) = 15
j = j + 1
GridCount = GridCount + 1
    frmPrint.sprPrint.SetText 1, j, "Page Number :" + str(pageno)
    j = j + 2
GridCount = GridCount + 2
    Spread_Set_CellBorder frmPrint.sprPrint, CellBorderIndexOutline, 1, j, frmPrint.sprPrint.MaxCols, j
    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.SetText 1, j, "PartNo"
frmPrint.sprPrint.ColWidth(1) = 12
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 2

    frmPrint.sprPrint.SetText 2, j, "Itemcode"
frmPrint.sprPrint.ColWidth(1) = 10
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 3

    frmPrint.sprPrint.SetText 3, j, "PlanQty"
frmPrint.sprPrint.ColWidth(2) = 13
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 4

    frmPrint.sprPrint.SetText 4, j, "Rate"
frmPrint.sprPrint.ColWidth(2) = 13
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 5

    frmPrint.sprPrint.SetText 5, j, "CompQty"
frmPrint.sprPrint.ColWidth(2) = 13
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 6

    frmPrint.sprPrint.SetText 6, j, "BalQty"
frmPrint.sprPrint.ColWidth(2) = 13
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 7

    frmPrint.sprPrint.SetText 7, j, "PlnAmt"
frmPrint.sprPrint.ColWidth(2) = 13
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 8

    frmPrint.sprPrint.SetText 8, j, "InvAmt"
frmPrint.sprPrint.ColWidth(2) = 13
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 9

    frmPrint.sprPrint.SetText 9, j, "BsalAmt"
frmPrint.sprPrint.ColWidth(2) = 13
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = 10

    frmPrint.sprPrint.SetText 10, j, "%Comp"
frmPrint.sprPrint.ColWidth(2) = 13
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft

    '.Row = j: .Col = 5

    'frmPrint.sprPrint.SetText 5, j, "BalQty": .ColWidth(2) = 10: .TypeHAlign = TypeHAlignLeft
    Head = j
  
    k = 1
    For X = FDATE To TDATE
      frmPrint.sprPrint.Row = j
frmPrint.sprPrint.Col = k + 9

      frmPrint.sprPrint.SetText k + 10, j, Format(X, "dd/MM")
frmPrint.sprPrint.ColWidth(k + 2) = 5
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.FontBold = True
      k = k + 1
    Next X
    frmPrint.sprPrint.MaxCols = frmPrint.sprPrint.DataColCnt
    Spread_BkColour frmPrint.sprPrint, &H80000004, 1, j, frmPrint.sprPrint.MaxCols, j
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Command1_Click()
rptfrmmonthwisesalesplan.Show
End Sub

Private Sub Command2_Click()
rptfrmmonthwisesalesfullview.Show
End Sub

Private Sub dtpProcessMonth_Change()
 'optYear.SetFocus
  Dim adoRs As AceADODB.Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim dat As Date
  Try
  Dim Sql As String
  CurDate = ServerDate
  'dat = Format(dtpDate.Value, "yyyy") + 1
  'dateFram(0).Visible = True
  'dateFram(1).Visible = False
  cmbDailyReNo.Clear
  'dtpDate.CustomFormat = "yyyy"
  'schedule.Caption = "Schedule Year"
  'dtpDate.Value = Format(CurDate, "dd-MM-yyyy")
  Set adoRs = New AceADODB.Recordset

  Sql = "select max(a.RevisionNo) RevNo from SAL_SalesPlan a "
  Sql = Sql & " where a.planqty<>'0'  and a.planFor=2 and month(a.planperiod) = " & Month(dtpProcessMonth) & " and year(a.planperiod) = " & Year(dtpProcessMonth) & " "
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "

  adoRs.Open Sql, db, adOpenStatic, adLockReadOnly
  adoRs.ActiveConnection = Nothing
  If adoRs.EOF Then
    cmbDailyReNo.Text = ""
    GoTo ExitHere
  End If
  cmbDailyReNo.Text = adoRs("RevNo")
  For i = 1 To adoRs.RecordCount
  cmbDailyReNo.AddItem (adoRs("RevNo"))
  adoRs.MoveNext
  Next i
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Sub

Private Sub Form_Activate()
 'optYear.SetFocus
  Dim adoRs As AceADODB.Recordset
  Dim i As Integer
  Dim CurDate As Date
  Dim dat As Date
  Try
  Dim Sql As String
  CurDate = ServerDate
  'dat = Format(dtpDate.Value, "yyyy") + 1
  'dateFram(0).Visible = True
  'dateFram(1).Visible = False
  cmbDailyReNo.Clear
  'dtpDate.CustomFormat = "yyyy"
  'schedule.Caption = "Schedule Year"
  'dtpDate.Value = Format(CurDate, "dd-MM-yyyy")
  Set adoRs = New AceADODB.Recordset

  Sql = "select max(a.RevisionNo) RevNo from SAL_SalesPlan a "
  Sql = Sql & " where a.planqty<>'0'  and a.planFor=2 and month(a.planperiod) = " & Month(dtpProcessMonth) & " and year(a.planperiod) = " & Year(dtpProcessMonth) & " "
  Sql = Sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' "
  
  adoRs.Open Sql, db, adOpenStatic, adLockReadOnly
  adoRs.ActiveConnection = Nothing
  If adoRs.EOF Then
    cmbDailyReNo.Text = ""
    GoTo ExitHere
  End If
  cmbDailyReNo.Text = adoRs("RevNo")
  For i = 1 To adoRs.RecordCount
  cmbDailyReNo.AddItem (adoRs("RevNo"))
  adoRs.MoveNext
  Next i
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adoRs = Nothing
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  Try

  dtpProcessMonth = Format(ServerDate, "yyyy-MM")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Function ServerDate() As Date
  Dim sdate As AceADODB.Recordset
  Dim Sql As String

  Try

  Sql = "select getdate()"
  Set sdate = New AceADODB.Recordset
  If RSOpen(sdate, Sql) = False Then
    GoTo ExitHere
  End If
  ServerDate = Format(sdate(0), "dd/MM/yyyy HH:mm:ss")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set sdate = Nothing
End Try
End Function

Private Sub optAll_Click()
' On Error Resume Next
  spList.Visible = False
End Sub

Private Sub optRandom_Click()
  
  Dim Sql As String
  Dim dat As Double
  Dim rs As AceADODB.Recordset
  
  Try
  Me.MousePointer = vbHourglass
  dat = Format(dtpProcessMonth, "yyyy") + 1
  
  Sql = "select distinct a.Customercode [Customer Code],c.Partyname [Customer Name] from SAL_SalesPlan a"
  Sql = Sql & " left outer join FAS_Party_Master c on a.CustomerCode =c.PartyCode"
  Sql = Sql & " and a.Branchid=c.Branchid and a.Compid=c.Compid"
  If optYear Then
  Sql = Sql & " where a.PlanFor = 1 and a.RevisionNo =' " & cmbReNo.Text & "' and a.PlanPeriod between '" & Format(dtpDate.Value, "yyyy") & "-04-01'and '" & CInt(dat) & "-03-31'"
  ElseIf optMonth Then
  Sql = Sql & " where a.PlanFor = 2 and a.RevisionNo = '" & cmbReNo.Text & "' and month(a.PlanPeriod)='" & Format(dtpDate.Value, "MM") & "'and year(a.planPeriod)='" & Format(dtpDate.Value, "yyyy") & "'"
  ElseIf optWeek Then
  Sql = Sql & " where a.PlanFor = 3 and a.RevisionNo = '" & cmbReNo.Text & " 'and month(a.PlanPeriod)='" & Format(dtpDate.Value, "MM") & "'and year(a.planPeriod)='" & Format(dtpDate.Value, "yyyy") & "'"
  ElseIf optDay Then
  Sql = Sql & " where a.PlanFor = 4 and a.RevisionNo = '" & cmbDailyReNo.Text & "' and month(a.PlanPeriod)='" & Format(dtpDailyDate.Value, "MM") & "'and year(a.planPeriod)='" & Format(dtpDailyDate.Value, "yyyy") & "'"
  End If
  Sql = Sql & " and a.Branchid='" & gBranchID & "' "
  Sql = Sql & " and a.Compid = '" & gCompID & "'"
 
'  adoList.ConnectionString = db
'  adoList.RecordSource = Sql
'  adoList.Refresh
'  spList.MaxRows = spList.DataRowCnt
'  spList.Visible = True
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If rs.EOF Then
    sShowMessage "Record(s) Not Found"
    GoTo ExitHere
  Else
    spList.Enabled = True
 spList.Visible = True
    spList.ClearSelection
    Call ADORS_Spread_DataFill(rs, spList, True)
    
    spList.MaxRows = spList.DataRowCnt
    Call Spread_Set_Focus(spList, 1, 1)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub Design_spList(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spList

End Sub

Private Sub InitializeSpreads()

	Call Design_spList("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptfrmmonthwisesalesm.frm", spList, "")

End Sub
