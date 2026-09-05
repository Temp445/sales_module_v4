VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesOrderAuthorisation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order Authorisation"
   ClientHeight    =   6450
   ClientLeft      =   -135
   ClientTop       =   1770
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   9060
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   2895
      TabIndex        =   31
      Top             =   0
      Width           =   2220
      Begin VB.OptionButton OptSal 
         Caption         =   "Sales"
         Height          =   225
         Left            =   150
         TabIndex        =   33
         Top             =   255
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton OptJob 
         Caption         =   "JobWork"
         Height          =   375
         Left            =   1125
         TabIndex        =   32
         Top             =   180
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      Height          =   480
      Left            =   6165
      TabIndex        =   10
      Top             =   6570
      Width           =   1830
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      CausesValidation=   0   'False
      Height          =   480
      Left            =   4350
      TabIndex        =   29
      Top             =   6570
      Width           =   1830
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   480
      Left            =   2535
      TabIndex        =   30
      Top             =   6570
      Width           =   1830
   End
   Begin VB.Frame fraDate 
      Height          =   675
      Left            =   3765
      TabIndex        =   24
      Top             =   585
      Width           =   3705
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   315
         Left            =   3150
         TabIndex        =   25
         Top             =   240
         Width           =   435
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   84606977
         CurrentDate     =   37946
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   1830
         TabIndex        =   27
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   84606977
         CurrentDate     =   37946
      End
      Begin VB.Label lblDocDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1515
         TabIndex        =   28
         Top             =   270
         Width           =   165
      End
   End
   Begin VB.Frame fraDetail 
      Height          =   2280
      Left            =   525
      TabIndex        =   22
      Top             =   1215
      Width           =   7170
Begin AceSpread sprOrder
         Height          =   1980
         Left            =   75
         TabIndex        =   23
         Top             =   210
         Width           =   7020
         _Version        =   458752
         _ExtentX        =   12383
         _ExtentY        =   3493
         _StockProps     =   64
         ButtonDrawMode  =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxRows         =   1
         ScrollBars      =   2
      End
   End
   Begin VB.Frame fraOrderDetails 
      Height          =   2760
      Left            =   525
      TabIndex        =   18
      Top             =   3705
      Width           =   10200
      Begin VB.TextBox txtVendor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   120
         TabIndex        =   19
         ToolTipText     =   "Enter  The Reason"
         Top             =   345
         Width           =   3870
      End
Begin AceSpread sprItem
         Height          =   1980
         Left            =   60
         TabIndex        =   20
         Top             =   675
         Width           =   10065
         _Version        =   458752
         _ExtentX        =   17754
         _ExtentY        =   3492
         _StockProps     =   64
         ButtonDrawMode  =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   7
         MaxRows         =   1
         ScrollBars      =   2
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   60
         TabIndex        =   21
         Top             =   165
         Width           =   1140
      End
   End
   Begin VB.Frame fraCalc 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   7770
      TabIndex        =   11
      Top             =   825
      Visible         =   0   'False
      Width           =   3660
      Begin VB.TextBox txtGoodsValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Select the Enquiry Number. Press F2 / Double Click for Help."
         Top             =   150
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.TextBox txtUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   600
         TabIndex        =   13
         Top             =   150
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtGrandTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1875
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2580
         Width           =   1470
      End
Begin AceSpread sprCalc
         Height          =   2085
         Left            =   45
         TabIndex        =   15
         Top             =   480
         Width           =   3570
         _Version        =   458752
         _ExtentX        =   6297
         _ExtentY        =   3678
         _StockProps     =   64
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
         MaxCols         =   2
         MaxRows         =   1
         ScrollBars      =   2
      End
      Begin VB.Label lblUserKey 
         AutoSize        =   -1  'True
         Caption         =   "User Key"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   75
         TabIndex        =   17
         Top             =   165
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblGrandTotal 
         AutoSize        =   -1  'True
         Caption         =   "Grand Total"
         Height          =   195
         Left            =   975
         TabIndex        =   16
         Top             =   2625
         Width           =   840
      End
   End
   Begin VB.Frame fraEffective 
      Height          =   720
      Left            =   7620
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   4545
      Begin VB.ComboBox cmbEffStatus 
         Height          =   330
         ItemData        =   "SalesOrderAuthorisation.frx":0DBE
         Left            =   75
         List            =   "SalesOrderAuthorisation.frx":0DC8
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   1170
      End
      Begin MSComCtl2.DTPicker dtpCreatedDate 
         Height          =   315
         Left            =   2565
         TabIndex        =   5
         Top             =   330
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   84606979
         CurrentDate     =   37946
      End
      Begin MSComCtl2.DTPicker dtpEffDate 
         Height          =   315
         Left            =   1260
         TabIndex        =   6
         Top             =   330
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   84606979
         CurrentDate     =   37946
      End
      Begin VB.Label lblCreatedDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Created Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2565
         TabIndex        =   9
         Top             =   150
         Width           =   945
      End
      Begin VB.Label lblEffDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "EffDate"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1260
         TabIndex        =   8
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lblEffStatus 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "EffStatus"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   75
         TabIndex        =   7
         Top             =   150
         Width           =   645
      End
   End
   Begin VB.Frame fraOrderType 
      Height          =   600
      Left            =   690
      TabIndex        =   0
      Top             =   600
      Width           =   3060
      Begin VB.OptionButton optAmend 
         Caption         =   "AmendmentDate"
         Height          =   375
         Left            =   1425
         TabIndex        =   2
         Top             =   180
         Width           =   1515
      End
      Begin VB.OptionButton optOrder 
         Caption         =   "OrderDate"
         Height          =   225
         Left            =   165
         TabIndex        =   1
         Top             =   255
         Value           =   -1  'True
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmSalesOrderAuthorisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'For Order Spread Column Nos
Private Enum eOrder
  mOrdType = 1
  mOrdNo = 2
  mOrdDate = 3
  mAuthorise = 4
  mCurStatus = 5
  mStatusDate = 6
  mVendor = 7
  mOrderingType = 8
  mCreatedDate = 9
  mStatusDateKey = 10
End Enum

'For Order Item Spread Column Nos
Private Enum eItem
  mItem = 1
  mDesc = 2
  mUOM = 3
  mOrdQty = 4
  mRate = 5
  mAssRate = 6
  mActive = 7
End Enum

'For Calculation Spread Column Nos
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Public Function gFinStart() As Date
' On Error Resume Next
  gFinStart = CDate("01-" & Left$(gFinyear, 7))
End Function

Public Function gFinEnd() As Date
' On Error Resume Next
  gFinEnd = CDate("31-" & Mid$(gFinyear, 8))
End Function

Private Sub SetSheet(ByVal fRow As Long)
  
' On Error Resume Next
    If SprItem.SheetCount < fRow Then
      SprItem.SheetCount = fRow
      Call CopySpreadStructure(SprItem, 1, , fRow)
    End If
    SprItem.Sheet = fRow
    SprItem.ActiveSheet = fRow
    If sprCalc.SheetCount < fRow Then
      sprCalc.SheetCount = fRow
      Call CopySpreadStructure(sprCalc, 1, , fRow)
    End If
    sprCalc.Sheet = fRow
    sprCalc.ActiveSheet = fRow
End Sub

Private Sub CmdClear_Click()
  
  sprCalc.SheetCount = 1
  sprCalc.MaxRows = 1
  SprItem.SheetCount = 1
  SprItem.MaxRows = 1
  sprOrder.MaxRows = 1
  
  Call Form_Clear(Me)
  Fradate.Enabled = False
  CmdSave.Enabled = False
  Fradate.Enabled = True
  fraDetail.Enabled = False
  fraOrderType.Enabled = True
  Frame1.Enabled = True
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim sql As String
  Dim rs As AceADODB.Recordset             'Order
  Dim rsDet As AceADODB.Recordset          'Item
  Dim rCnt As Long      'Order Row Count
  Dim irCnt As Long     'Item Row Count
  Dim curdate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  If dtpToDate.Value < dtpFromDate.Value Then
    sShowMessage "To Date should be less than or equal to From Date."
    dtpToDate.SetFocus
    GoTo ExitHere
  End If
  
  curdate = ServerDate
  sprOrder.Enabled = False
  SprItem.Enabled = False
  
  If OptSal Then
    sql = "Select distinct a.OrderType, a.OrderNo, a.OrderDate,Case a.Status when 'N' then 'New' when 'M' then 'Amendment' when 'A' then 'Authorised' when 'H' then 'Hold' End as Status,"
    sql = sql & " a.CreatedDate,(select PartyName From Fas_Party_Master"
    sql = sql & " Where PartyCode=a.CustomerCode and Branchid=a.Branchid and Compid=a.Compid)CustomerName,a.CustomerCode,a.OrderingType"
    sql = sql & " From dbo.fn_Sal_Order_Hdr('" & Format(curdate, "yyyy-MM-dd 23:59:59") & "', '" & Format(curdate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    If optAmend Then
      sql = sql & " Inner join dbo.fn_SAL_Order_Amendment('" & Format(curdate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')b"
      sql = sql & " On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.Orderdate=b.Orderdate"
      sql = sql & " Where a.Status<>'C' and b.AmendmentDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd 23:59:59") & "'"
    Else
      sql = sql & " Where a.Status<>'C' and a.OrderDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
    End If
    sql = sql & " Order by a.OrderType, a.OrderNo, a.OrderDate"
  Else
    sql = " Select distinct a.OrderType, a.OrderNo, a.OrderDate,Case a.Status when 'N' then 'New' when 'M' then 'Amendment' when 'A' then 'Authorised' when 'H' then 'Hold' End as Status,"
    sql = sql & " a.CreatedDate,(select PartyName from Fas_Party_Master"
    sql = sql & " Where PartyCode=a.CustomerCode and Branchid=a.Branchid and Compid=a.Compid)CustomerName,a.CustomerCode,a.OrderingType"
    sql = sql & " From dbo.fn_Sal_JobWorkOrder_Hdr('" & Format(curdate, "yyyy-MM-dd 23:59:59") & "', '" & Format(curdate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    If optAmend Then
      sql = sql & " Inner join dbo.fn_SAL_Order_Amendment('" & Format(curdate, "yyyy-MM-dd 23:59:59") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')b"
      sql = sql & " On a.OrderType=b.OrderType and a.OrderNo=b.OrderNo and a.Orderdate=b.Orderdate"
      sql = sql & " Where a.Status<>'C' and b.AmendmentDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd 23:59:59") & "'"
    Else
      sql = sql & " Where a.Status<>'C' and a.OrderDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
    End If
    sql = sql & " Order by a.OrderType, a.OrderNo, a.OrderDate"
  End If
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    sShowMessage "No Order Details Found"
    GoTo ExitHere
  End If
  txtVendor.Tag = rs("CustomerCode")
  Fradate.Enabled = False
  
  If Not rs.EOF Then
    rs.MoveLast
rs.MoveFirst
      rCnt = rs.RecordCount
      sprOrder.MaxRows = rCnt
      For rCnt = 1 To rs.RecordCount
        sprOrder.SetText eOrder.mOrdType, rCnt, rs("OrderType")
        sprOrder.SetText eOrder.mOrdNo, rCnt, rs("OrderNo")
        sprOrder.SetText eOrder.mOrdDate, rCnt, Format(rs("OrderDate"), "dd/MM/yyyy")
        sprOrder.SetText eOrder.mCurStatus, rCnt, rs("Status")
        sprOrder.SetText eOrder.mStatusDate, rCnt, Format(rs("CreatedDate"), "dd/MM/yyyy")
        sprOrder.SetText eOrder.mVendor, rCnt, rs("CustomerName")
        sprOrder.SetText eOrder.mOrderingType, rCnt, rs("OrderingType")
        sprOrder.SetText eOrder.mCreatedDate, rCnt, Format(rs("CreatedDate"), "yyyy-MM-dd HH:mm:ss")
        sprOrder.SetText eOrder.mStatusDateKey, rCnt, Format(rs("CreatedDate"), "yyyy-MM-dd HH:mm:ss")
        
        'Item Details loaded here
        If OptSal Then
          sql = "Select a.ItemCode, a.ItemDescription,a.OrderQty as Qty,a.Rate,a.AssessableRate,"
          sql = sql & " (Select sDesc from INV_UOM_MASTER where Code=a.UOM and BRanchid=a.Branchid and Compid=a.Compid)UOM,"
          sql = sql & " (Select NoofDecimals from INV_UOM_MASTER where Code=a.UOM and BRanchid=a.Branchid and Compid=a.Compid)NoOfDec,Case a.Status when 'N' then 'New' when 'M' then 'Amendment' when 'A' then 'Authorised' when 'H' then 'Hold' End as Status"
          sql = sql & " From dbo.fn_Sal_Order_Dtl('" & Format(curdate, "yyyy-MM-dd 23:59:59") & "', '" & Format(curdate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
          sql = sql & " Where Status<>'C' and OrderType='" & rs("OrderType") & "' and OrderNo=" & rs("OrderNo") & " and OrderDate='" & Format(rs("OrderDate"), "yyyy-MM-dd") & "'"
          sql = sql & " Order by a.ItemCode"
        Else
          sql = "Select distinct a.SentItemCode ItemCode, b.ItemDescription, b.UOMDesc UOM,b.UOMDecimal NoOfDec, a.OrderQty Qty, a.Rate, "
          sql = sql & " a.AssRate AssessableRate,a.Status"
          sql = sql & " From dbo.fn_SAL_JobWorkOrder ('" & Format(curdate, "yyyy-MM-dd 23:59:59") & "', '" & Format(curdate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a Left Outer Join Material_Master b "
          sql = sql & " On a.SentItemCode = b.ItemCode and a.Branchid = b.Branchid and a.Compid = b.Compid "
          sql = sql & " Where a.OrderType = '" & rs("OrderType") & "'  and a.OrderNo = "
          sql = sql & " " & rs("OrderNo") & " and a.OrderDate = '" & Format(rs("OrderDate"), "yyyy-MM-dd") & "'"
          sql = sql & " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "' and a.Status<>'C' and a.ItemStatus<>'C' and b.Status = 1"
        End If
        Set rsDet = New AceADODB.Recordset
        If RSOpen(rsDet, sql) = False Then
          GoTo ExitHere
        End If
        
        If Not rsDet.EOF Then
          rsDet.MoveLast
rsDet.MoveFirst
          With SprItem
            Call SetSheet(rCnt)
            sprOrder.MaxRows = rsDet.RecordCount
            For irCnt = 1 To rsDet.RecordCount
              sprOrder.SetText eItem.mItem, irCnt, rsDet("ItemCode")
              sprOrder.SetText eItem.mDesc, irCnt, rsDet("ItemDescription")
              sprOrder.SetText eItem.mUOM, irCnt, rsDet("UOM")
              Call Spread_Set_Decimal(SprItem, Val(CheckNull(rsDet("NoOfDec"))), eItem.mOrdQty, irCnt, eItem.mOrdQty)
              sprOrder.SetText eItem.mOrdQty, irCnt, IIf(rsDet("Qty") > 0, rsDet("Qty"), "")
              sprOrder.SetText eItem.mRate, irCnt, rsDet("Rate")
              sprOrder.SetText eItem.mAssRate, irCnt, rsDet("AssessableRate")
              sprOrder.SetText eItem.mActive, irCnt, rsDet("Status")
              rsDet.MoveNext
            Next irCnt
            Call SetSheet(rCnt)
            SprItem.MaxRows = rsDet.RecordCount
            For irCnt = 1 To rsDet.RecordCount
              SprItem.SetText eItem.mItem, irCnt, rsDet("ItemCode")
              SprItem.SetText eItem.mDesc, irCnt, rsDet("ItemDescription")
              SprItem.SetText eItem.mUOM, irCnt, rsDet("UOM")
              Call Spread_Set_Decimal(SprItem, Val(CheckNull(rsDet("NoOfDec"))), eItem.mOrdQty, irCnt, eItem.mOrdQty)
              SprItem.SetText eItem.mOrdQty, irCnt, IIf(rsDet("Qty") > 0, rsDet("Qty"), "")
              SprItem.SetText eItem.mRate, irCnt, rsDet("Rate")
              SprItem.SetText eItem.mAssRate, irCnt, rsDet("AssessableRate")
              SprItem.SetText eItem.mActive, irCnt, rsDet("Status")
              rsDet.MoveNext
            Next irCnt
        End If
        'Calculation Details shown here for specific order
        If rs("OrderingType") = "Specific" Then
          If LoadCalcDetails(rs("OrderType"), rs("OrderNo"), rs("OrderDate"), rCnt) = False Then
            GoTo ExitHere
          End If
        End If
        
        rs.MoveNext
      Next rCnt
    End With
    Call OrderRowChangeOperation(1)
  End If
  
  sprOrder.Enabled = True
  SprItem.Enabled = True
  CmdSave.Enabled = True
  fraDetail.Enabled = True
  Fradate.Enabled = False
  fraOrderType.Enabled = False
  Frame1.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set rsDet = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Function LoadCalcDetails(ByVal fOrdType As String,ByVal  fOrdNo As Long,ByVal  fOrdDate As Date,ByVal  fRow As Long) As Boolean
  Dim sql As String
  Dim rsCalc As AceADODB.Recordset
  Dim rCnt As Long
  
  Try
  LoadCalcDetails = False
  
  sql = " Select * from dbo.fn_COM_GetExprResults_Saved('" & fOrdType & "'," & fOrdNo & ",'" & Format(fOrdDate, "yyyy-MM-dd") & "', 4, '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') "
  sql = sql & " Order by ExprSeq"
  Set rsCalc = New AceADODB.Recordset
  If RSOpen(rsCalc, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsCalc.EOF Then
    rsCalc.MoveNext
rsCalc.MoveFirst
      Call SetSheet(fRow)
      sprCalc.MaxRows = rsCalc.RecordCount - 1
      For rCnt = 1 To rsCalc.RecordCount - 1
        sprCalc.SetText eCalc.mDesc, rCnt, rsCalc("ExprName")
        sprCalc.SetText eCalc.mValue, rCnt, rsCalc("ResultValue")
        rsCalc.MoveNext
      Next rCnt
      txtGrandTotal.Text = Format(Val(rsCalc("ResultValue")), "0.00")
  End If
  
  LoadCalcDetails = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsCalc = Nothing
End Try
End Function

Private Sub FillTotalValue(ByVal fOrdType As String,ByVal  fOrdNo As Long,ByVal  fOrdDate As Date,ByVal  fRow As Long)
  Dim sql As String
  Dim rsTot As AceADODB.Recordset
  
  Try
  
  sql = " select Sum(GrandTotalFlag * ResultValue) as TotalValue from dbo.fn_COM_GetExprResults_Saved('" & fOrdType & "'," & fOrdNo & ",'" & Format(fOrdDate, "yyyy-MM-dd") & "', 4, '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')"
  Set rsTot = New AceADODB.Recordset
  If RSOpen(rsTot, sql) = False Then
    GoTo ExitHere
  End If
  
  txtGrandTotal.Text = Format(rsTot("TotalValue"), "0.00")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsTot = Nothing
End Try
End Sub

Private Sub OrderRowChangeOperation(ByVal fRow As Long)
  Dim retVal(1 To 3) As Variant
  
  Try
  
  If fRow > 0 Then
      sprOrder.GetText eOrder.mVendor, fRow, retVal(1)
      txtVendor.Text = Trim$(retVal(1))
      SprItem.Sheet = fRow
      SprItem.ActiveSheet = fRow
      
      sprOrder.GetText eOrder.mOrderingType, fRow, retVal(2)
      If Trim$(retVal(2)) = "Specific" Then
        sprCalc.Sheet = fRow
        sprCalc.ActiveSheet = fRow
        
        sprOrder.GetText eOrder.mOrdType, fRow, retVal(1)
        sprOrder.GetText eOrder.mOrdNo, fRow, retVal(2)
        sprOrder.GetText eOrder.mOrdDate, fRow, retVal(3)
        Call FillTotalValue(CStr(retVal(1)), Val(retVal(2)), CDate(retVal(3)), fRow)
        fraCalc.Visible = True
      Else
        fraCalc.Visible = False
      End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function ValidateData() As Boolean
  Dim rCnt As Long
  Dim retVal As Variant
  
  Try
  ValidateData = True
  
    For rCnt = 1 To sprOrder.DataRowCnt
      sprOrder.GetText eOrder.mAuthorise, rCnt, retVal
      If Trim$(retVal) <> "" Then
        GoTo ExitHere
      End If
    Next rCnt
  
  sShowMessage "No Data found to Save."
  ValidateData = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdSave_Click()
  Dim sql As String
  Dim rCnt As Long
  Dim retVal As Variant
  Dim tmp As Variant
  Dim curdate As Date
  Dim cmd As AceADODB.Command
  
  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  Set cmd = New Command
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  curdate = ServerDate
  
    For rCnt = 1 To sprOrder.DataRowCnt
      sprOrder.GetText eOrder.mAuthorise, rCnt, retVal
      If Len(Trim$(retVal)) > 0 Then
        If OptSal Then
          sql = "Insert into Sal_Order_Hdr "
          sql = sql & " Select OrderType,OrderNo,OrderDate,OrderSource,CustomerCode,ReferenceNo,ReferenceDate,ValidFrom,ValidTo,"
          sql = sql & " OrderingType,Remarks,FormCode,PackingCode,DeliveryCode,PaymentCode,InsuranceCode,ModeCode,Freightcode,"
          sprOrder.GetText eOrder.mAuthorise, rCnt, retVal
          sql = sql & " FormulaCode,TotalAmount,PostingAc,RevisionLevel,GeneralDetails, '" & Left$(retVal, 1) & "','" & Format(curdate, "yyyy-MM-dd HH:mm:ss") & "',"
          sql = sql & " 'A', '" & Format(curdate, "yyyy-MM-dd HH:mm:ss") & "','" & gUserID & "','" & gUserID & "',Branchid,Compid,Finyear, '" & Format(curdate, "yyyy-MM-dd HH:mm:ss") & "'"
          sql = sql & " From dbo.fn_Sal_Order_Hdr('" & Format(curdate, "yyyy-MM-dd 23:59:59") & "', '" & Format(curdate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a "
          sprOrder.GetText eOrder.mOrdType, rCnt, retVal
          sql = sql & " Where OrderType='" & Trim$(retVal) & "'"
          sprOrder.GetText eOrder.mOrdNo, rCnt, retVal
          sql = sql & " and OrderNo=" & Val(retVal)
          sprOrder.GetText eOrder.mOrdDate, rCnt, retVal
          sql = sql & " and OrderDate='" & Format(retVal, "yyyy-MM-dd") & "'"
          sprOrder.GetText eOrder.mAuthorise, rCnt, retVal
          sql = sql & " and Status<>'" & Left$(retVal, 1) & "'"
          sql = sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
          If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        Else
          sql = "Insert into SAL_JOBWORKORDER_HDR "
          sql = sql & " Select OrderType,OrderNo,OrderDate,OrderSource,CustomerCode,ReferenceNo,ReferenceDate,ValidFrom,ValidTo,OrderingType,"
          sprOrder.GetText eOrder.mAuthorise, rCnt, retVal
          sql = sql & " Remarks, FormCode, PackingCode, DeliveryCode, PaymentCode, InsuranceCode, ModeCode, Freightcode,"
          sql = sql & " FormulaCode, TotalAmount, PostingAc,'" & Left$(retVal, 1) & "','" & Format(curdate, "yyyy-MM-dd HH:mm:ss") & "',"
          sql = sql & " Status, OrderAuthoriZationDate, OrderAuthoriZationBy, '" & gUserID & "',Branchid,Compid,Finyear, '" & Format(curdate, "yyyy-MM-dd HH:mm:ss") & "'"
          sql = sql & " From dbo.fn_Sal_JobWorkOrder_Hdr('" & Format(curdate, "yyyy-MM-dd 23:59:59") & "', '" & Format(curdate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
          sprOrder.GetText eOrder.mOrdType, rCnt, retVal
          sql = sql & " Where OrderType='" & Trim$(retVal) & "'"
          sprOrder.GetText eOrder.mOrdNo, rCnt, retVal
          sql = sql & " and OrderNo=" & Val(retVal)
          sprOrder.GetText eOrder.mOrdDate, rCnt, retVal
          sql = sql & " and OrderDate='" & Format(retVal, "yyyy-MM-dd") & "'"
          sprOrder.GetText eOrder.mAuthorise, rCnt, retVal
          sql = sql & " and Status<>'" & Left$(retVal, 1) & "'"
          sql = sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
          If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        End If
      End If
    Next rCnt
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record Saved Successfully"
  Call CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim curdate As Date
  
  Try

  curdate = ServerDate
'  Call CurrDate(Me)
  
'  dtpFromDate.MinDate = gFinStart
  dtpFromDate.Value = Format(curdate, "dd/MM/yyyy")
  dtpFromDate.MaxDate = gFinEnd
  
  dtpToDate.Value = Format(curdate, "dd/MM/yyyy")
  dtpToDate.MaxDate = dtpToDate.Value
  
  sprOrder.Enabled = False
'  fraDate.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = 1
  End If
End Sub

Private Sub optAmend_Click()

  Try
     
  If optAmend Then
    Fradate.Enabled = True
    Fradate.Caption = "Amendment Date"
   Else
    Fradate.Caption = " "
    Fradate.Enabled = False
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub optorder_Click()
  Try
  
  If optOrder Then
    Fradate.Enabled = True
    Fradate.Caption = "Order Date"
  Else
    Fradate.Caption = " "
    Fradate.Enabled = False
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub sprOrder_LeaveCell(sender as object, e as LeaveCellEventArgs)
  
  Try
  Me.MousePointer = vbHourglass

  If Row <> NewRow Then
    Call OrderRowChangeOperation(NewRow)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Design_sprOrder(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprOrder


sprOrder.LeaveCell +=  new EventHandler(sprOrder_LeaveCell)

End Sub

Private Sub Design_sprItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprItem

End Sub

Private Sub Design_sprCalc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprCalc

End Sub

Private Sub InitializeSpreads()

	Call Design_sprOrder("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderAuthorisation.frm", sprOrder, "")

	Call Design_sprItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderAuthorisation.frm", sprItem, "")

	Call Design_sprCalc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesOrderAuthorisation.frm", sprCalc, "")

End Sub
