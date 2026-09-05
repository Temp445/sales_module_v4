VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSalesOrderShortClose 
   Caption         =   "Order ShortClose/ Cancel"
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
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame fraDate 
      Caption         =   "Order Date"
      Height          =   645
      Left            =   4140
      TabIndex        =   4
      Top             =   960
      Width           =   3420
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   315
         Left            =   2865
         TabIndex        =   3
         Top             =   210
         Width           =   435
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55771137
         CurrentDate     =   37946
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         Top             =   210
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   55771137
         CurrentDate     =   37946
      End
      Begin VB.Label lblDocDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "to"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1335
         TabIndex        =   1
         Top             =   240
         Width           =   165
      End
   End
   Begin VB.Frame fraOrderDetails 
      Height          =   2760
      Left            =   285
      TabIndex        =   7
      Top             =   3825
      Width           =   11130
      Begin VB.TextBox txtVendor 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   120
         TabIndex        =   9
         ToolTipText     =   "Enter  The Reason"
         Top             =   345
         Width           =   3870
      End
Begin AceSpread sprItem
         Height          =   1980
         Left            =   90
         TabIndex        =   10
         Top             =   675
         Width           =   10920
         _Version        =   458752
         _ExtentX        =   19262
         _ExtentY        =   3493
         _StockProps     =   64
         ButtonDrawMode  =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   10
         MaxRows         =   1
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   165
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6495
      TabIndex        =   13
      Top             =   6735
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   5200
      TabIndex        =   12
      Top             =   6735
      Width           =   1300
   End
   Begin VB.Frame fraDetail 
      Height          =   2250
      Left            =   1778
      TabIndex        =   6
      Top             =   1575
      Width           =   8145
Begin AceSpread sprOrder
         Height          =   1980
         Left            =   75
         TabIndex        =   5
         Top             =   180
         Width           =   7980
         _Version        =   458752
         _ExtentX        =   14076
         _ExtentY        =   3492
         _StockProps     =   64
         ButtonDrawMode  =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   10
         MaxRows         =   1
         ScrollBars      =   2
      End
   End
   Begin TabDlg.SSTab StabItem 
      Height          =   3210
      Left            =   360
      TabIndex        =   14
      Top             =   7995
      Visible         =   0   'False
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   5662
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      Enabled         =   0   'False
      TabCaption(0)   =   "ItemCode"
      TabPicture(0)   =   "frmSalesOrderShortClose.frx":0CF6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Schedule"
      TabPicture(1)   =   "frmSalesOrderShortClose.frx":0D12
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "sprSchedule"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
Begin AceSpread sprSchedule
         Height          =   1575
         Left            =   1740
         TabIndex        =   15
         Top             =   855
         Width           =   7425
         _Version        =   458752
         _ExtentX        =   13097
         _ExtentY        =   2778
         _StockProps     =   64
         ButtonDrawMode  =   4
         DAutoCellTypes  =   0   'False
         DAutoHeadings   =   0   'False
         DAutoSave       =   0   'False
         DAutoSizeCols   =   0
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
         MaxCols         =   7
         MaxRows         =   1
         ScrollBars      =   2
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3915
      TabIndex        =   11
      Top             =   6735
      Width           =   1300
   End
End
Attribute VB_Name = "frmSalesOrderShortClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private sButtonClick As Boolean
Public FrmFlg As String

Private Enum eOrder
  mOrdType = 1
  mOrdNo = 2
  mOrdDate = 3
  mShortClose = 4
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
  mInvoiceQty = 5
  mPendQty = 6
  mRatePer = 7
  mRate = 8
  mAssRate = 9
  mShortClose = 10
End Enum

Private Enum eSch
  mSchType = 1
  mSchNo = 2
  mSchDate = 3
  mItemCode = 4
  mSchQty = 5
  mComQty = 6
  mComStatus = 7
End Enum

Private Function SaveDetailData(ByVal CurDate As Date,ByVal  rCnt As Long,ByVal  iCnt As Long) As Boolean
'  Dim Sql As String
'  Dim rs As ADODB.Recordset
'  Dim retVal As Variant
'
'  On Error GoTo ErrHandler
'  SaveDetailData = False
'
'  Sql = " Select *, case When Exists(SSelect * From "
'  Sql = Sql & " SAL_Invoice_Dtl c"
'  Sql = Sql & " Where c.OrderType=a.OrderType and c.OrderNo=a.OrderNo and c.OrderDate=a.OrderDate and c.Branchid='" & gBranchID & "' and c.Compid='" & gCompID & "')"
'  Sql = Sql & "  then 'S' else 'C' end as ShorClose"
'  Sql = Sql & " From dbo.fn_SAL_order_Dtl('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')a"
'  Sql = Sql & " Where "
'  sprOrder.GetText eOrder.mOrdType, rCnt, retVal
'  Sql = Sql & " OrderType='" & Trim$(retVal) & "' and "
'  sprOrder.GetText eOrder.mOrdNo, rCnt, retVal
'  Sql = Sql & " OrderNo='" & Val(retVal) & "' and "
'  sprOrder.GetText eOrder.mOrdDate, rCnt, retVal
'  Sql = Sql & " OrderDate='" & Format(retVal, "yyyy-MM-dd") & "' and "
'  Sql = Sql & " BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
'  Set rs = New ADODB.Recordset
'  If RSOpen(rs, Sql) = False Then
'    GoTo ExitHere
'  End If
'  If Not rs.EOF Then
'    Sql = "Insert into SAL_Order_Dtl Values("
'    Sql = Sql & " '" & Trim$(rs("OrderType")) & "',"
'    Sql = Sql & " " & Val(rs("OrderNo")) & ","
'    Sql = Sql & " '" & Format(rs("OrderDate"), "yyyy-MM-dd") & "',"
'    sprItem.GetText eItem.mItem, iCnt, retVal
'    Sql = Sql & " '" & Trim$(retVal) & "',"
'    sprItem.GetText eItem.mDesc, iCnt, retVal
'    Sql = Sql & " '" & Trim$(retVal) & "',"
'    Sql = Sql & " '" & Trim$(rs("PartNo")) & "',"
'    Sql = Sql & " '" & Val(rs("RevisionNo")) & "',"
'    Sql = Sql & " '" & Val(rs("UOM")) & "',"
'    Sql = Sql & " '" & rs("ScheduleType") & "',"
'    Sql = Sql & " '" & Val(rs("OrderQty")) & "',"
'    Sql = Sql & " '" & rs("Rate") & "',"
'    Sql = Sql & " '" & Val(rs("AssessableRate")) & "',"
'    Sql = Sql & " '" & Trim$(rs("ShorClose")) & "',"
'    Sql = Sql & " '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
'    Sql = Sql & " '" & gUserID & "',"
'    Sql = Sql & " '" & gBranchID & "',"
'    Sql = Sql & " '" & gCompID & "',"
'    Sql = Sql & " '" & gFinyear & "'"
'    Sql = Sql & " )"
'    db.Execute Sql
'  End If
'
'  SaveDetailData = True
'
'ExitHere:
'  On Error GoTo 0
'  Exit Function
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Function

Private Function SaveJobworkData(ByVal CurDate As Date,ByVal  rCnt As Long) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim retVal As Variant

  Try
  
  SaveJobworkData = False
  
  Sql = " Select *, Case When Exists (Select * From "
  Sql = Sql & " SAL_Invoice_Dtl c"
  Sql = Sql & " Where c.OrderType=a.OrderType and c.OrderNo=a.OrderNo and c.OrderDate=a.OrderDate and c.Branchid='" & gBranchID & "' and c.Compid='" & gCompID & "'"
  Sql = Sql & " ) then 'S' else 'C' end as ShorClose"
  Sql = Sql & " From dbo.fn_SAL_JobworkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')a"
  Sql = Sql & " Where "
  sprOrder.GetText eOrder.mOrdType, rCnt, retVal
  Sql = Sql & " OrderType='" & Trim$(retVal) & "' and "
  sprOrder.GetText eOrder.mOrdNo, rCnt, retVal
  Sql = Sql & " OrderNo='" & Val(retVal) & "' and "
  sprOrder.GetText eOrder.mOrdDate, rCnt, retVal
  Sql = Sql & " OrderDate='" & Format(retVal, "yyyy-MM-dd") & "' "
  Sql = Sql & " and Status Not in('C','S','H') and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  If Not rs.EOF Then
    Sql = "Insert into SAL_JobworkOrder_Hdr Values("
    Sql = Sql & " '" & Trim$(rs("OrderType")) & "',"
    Sql = Sql & " " & Val(rs("OrderNo")) & ","
    Sql = Sql & " '" & Format(rs("OrderDate"), "yyyy-MM-dd") & "',"
    Sql = Sql & " '" & Trim$(rs("OrderSource")) & "',"
    Sql = Sql & " " & Val(rs("CustomerCode")) & ","
    Sql = Sql & " '" & Val(rs("ReferenceNo")) & "',"
    Sql = Sql & " '" & Format(rs("ReferenceDate"), "yyyy-MM-dd") & "',"
    Sql = Sql & " '" & Format(rs("ValidFrom"), "yyyy-MM-dd") & "',"
    Sql = Sql & " '" & Format(rs("ValidTo"), "yyyy-MM-dd") & "',"
    Sql = Sql & " '" & rs("OrderingType") & "',"
    Sql = Sql & " '" & rs("Remarks") & "',"
    Sql = Sql & " " & Val(rs("FormCode")) & ","
    Sql = Sql & " " & Val(rs("PackingCode")) & ","
    Sql = Sql & " " & Val(rs("DeliveryCode")) & ","
    Sql = Sql & " " & Val(rs("PaymentCode")) & ","
    Sql = Sql & " " & Val(rs("InsuranceCode")) & ","
    Sql = Sql & " " & Val(rs("ModeCode")) & ","
    Sql = Sql & " " & Val(rs("Freightcode")) & ","
    Sql = Sql & " " & Val(rs("FormulaCode")) & ","
    Sql = Sql & " " & Val(rs("TotalAmount")) & ","
    Sql = Sql & " " & Val(rs("PostingAc")) & ","
    Sql = Sql & " '" & Trim$(rs("ShorClose")) & "',"
    Sql = Sql & " '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
    Sql = Sql & " '" & rs("OrderAuthorization") & "',"
    Sql = Sql & " '" & Format(rs("OrderAuthorizationDate"), "yyyy-MM-dd HH:mm:ss") & "',"
    Sql = Sql & " '" & rs("OrderAuthorizationBy") & "',"
    Sql = Sql & " '" & gUserID & "',"
    Sql = Sql & " '" & gBranchID & "',"
    Sql = Sql & " '" & gCompID & "',"
    Sql = Sql & " '" & gFinyear & "',"
    Sql = Sql & " '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " )"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If

  SaveJobworkData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function SaveScheduleData(ByVal CurDate As Date,ByVal  rCnt As Long,ByVal  iCnt As Long) As Boolean
'  Dim Sql As String
'  Dim rs As ADODB.Recordset
'  Dim retVal As Variant
'
'  On Error GoTo ErrHandler
'
'  SaveScheduleData = False
'
'  Sql = " Select *, Case When Exists(Select a.ScheduleQty, Sum(b.InvoiceQty) as Qty"
'  Sql = Sql & " From dbo.fn_SAL_Order_Schedule('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
'  Sql = Sql & " Left Outer Join SAL_Invoice_Dtl b"
'  Sql = Sql & " On a.OrderDate = b.OrderDate and a.OrderNo = b.OrderNo and a.OrderType = b.OrderType"
'  Sql = Sql & " and a.BranchID = b.BranchID and a.COmpID = b.COmpID"
'  sprOrder.GetText eOrder.mOrdType, rCnt, retVal
'  Sql = Sql & " Where a.OrderType='" & Trim$(retVal) & "'and "
'  sprOrder.GetText eOrder.mOrdNo, rCnt, retVal
'  Sql = Sql & " a.OrderNo='" & Val(retVal) & "'and "
'  sprOrder.GetText eOrder.mOrdDate, rCnt, retVal
'  Sql = Sql & " a.OrderDate='" & Format(retVal, "yyyy-MM-dd") & "'and "
'  Sql = Sql & " a.CompID='" & gCompID & "'"
'  Sql = Sql & " Group By a.ScheduleQty"
'  Sql = Sql & " Having Sum(b.InvoiceQty) < a.ScheduleQty)"
'  Sql = Sql & " then 'less' else 'greater' end as Qty" 'InvoiceQty less or Greater
'  Sql = Sql & " From dbo.fn_SAL_Order_Schedule('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')a"
'  sprOrder.GetText eOrder.mOrdType, rCnt, retVal
'  Sql = Sql & " Where a.OrderType='" & Trim$(retVal) & "'and "
'  sprOrder.GetText eOrder.mOrdNo, rCnt, retVal
'  Sql = Sql & " a.OrderNo='" & Val(retVal) & "'and "
'  sprOrder.GetText eOrder.mOrdDate, rCnt, retVal
'  Sql = Sql & " a.OrderDate='" & Format(retVal, "yyyy-MM-dd") & "'and "
'  Sql = Sql & " a.CompID='" & gCompID & "'"
'  Set rs = New ADODB.Recordset
'  If RSOpen(rs, Sql) = False Then
'    GoTo ExitHere
'  End If
'  If Not rs.EOF Then
'    If rs("Qty") = "Less" Then
'      Sql = "Insert into SAL_Order_Schedule Values("
'      Sql = Sql & " '" & Trim$(rs("OrderType")) & "',"
'      Sql = Sql & " " & Val(rs("OrderNo")) & ","
'      Sql = Sql & " '" & Format(rs("OrderDate"), "yyyy-MM-dd") & "',"
'      sprItem.GetText eItem.mItem, rCnt, retVal
'      Sql = Sql & " '" & Trim$(retVal) & "',"
'      Sql = Sql & " '" & Val(rs("PartNo")) & "',"
'      Sql = Sql & " '" & Format(rs("ScheduleDate"), "yyyy-MM-dd") & "',"
'      Sql = Sql & " '" & -Val(rs("ScheduleQty")) & "',"
'      Sql = Sql & " 'C',"
'      Sql = Sql & " '" & Trim$(rs("Location")) & "',"
'      Sql = Sql & " '" & Format(CurDate, "yyyy-MM-dd") & "',"
'      Sql = Sql & " '" & gUserID & "',"
'      Sql = Sql & " '" & gBranchID & "',"
'      Sql = Sql & " '" & gCompID & "',"
'      Sql = Sql & " '" & gFinyear & "'"
'      Sql = Sql & " )"
'      db.Execute Sql
'    End If
'  End If
'
'  SaveScheduleData = True
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Function
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Function

Private Function ValidateData() As Boolean
  Dim rCnt As Long
  Dim Cnt As Long
  Dim retVal As Variant
  Dim sflag As Boolean
  
  Try
  
  ValidateData = False
  sflag = False
  
    For rCnt = Cnt To sprOrder.DataRowCnt
      sprOrder.GetText eOrder.mShortClose, rCnt, retVal
      If Val(retVal) = 1 Then
        sflag = True
Exit For
'      Else
'        sprItem.Sheet = rCnt
'        For Cnt = 1 To sprItem.DataRowCnt
'         sprItem.GetText eItem.mShortClose, Cnt, retVal
'          If Val(retVal) = 1 Then
'            sflag = True: Exit For
'          End If
'        Next Cnt
      End If
    Next rCnt
  
  If sflag = False Then
    sShowMessage "Please select atleast one order to shortclose"
    GoTo ExitHere
  End If
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function SaveHeaderData(ByVal CurDate As Date,ByVal  rCnt As Long) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim retVal As Variant

  Try
  SaveHeaderData = False
  
  Sql = " Select *, Case When Exists (Select * From "
  Sql = Sql & " SAL_Invoice_Dtl c"
  Sql = Sql & " Where c.OrderType=a.OrderType and c.OrderNo=a.OrderNo and c.OrderDate=a.OrderDate and c.Branchid='" & gBranchID & "' and c.Compid='" & gCompID & "'"
  Sql = Sql & " ) then 'S' else 'C' end as ShorClose"
  Sql = Sql & " From dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "','" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "')a"
  Sql = Sql & " Where "
  sprOrder.GetText eOrder.mOrdType, rCnt, retVal
  Sql = Sql & " OrderType='" & Trim$(retVal) & "' and "
  sprOrder.GetText eOrder.mOrdNo, rCnt, retVal
  Sql = Sql & " OrderNo='" & Val(retVal) & "' and "
  sprOrder.GetText eOrder.mOrdDate, rCnt, retVal
  Sql = Sql & " OrderDate='" & Format(retVal, "yyyy-MM-dd") & "' "
  Sql = Sql & " and Status Not in('C','S','H') and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  If Not rs.EOF Then
    Sql = "Insert into SAL_Order_Hdr Values("
    Sql = Sql & " '" & Trim$(rs("OrderType")) & "',"
    Sql = Sql & " " & Val(rs("OrderNo")) & ","
    Sql = Sql & " '" & Format(rs("OrderDate"), "yyyy-MM-dd") & "',"
    Sql = Sql & " '" & Trim$(rs("OrderSource")) & "',"
    Sql = Sql & " " & Val(rs("CustomerCode")) & ","
    Sql = Sql & " '" & Val(rs("ReferenceNo")) & "',"
    Sql = Sql & " '" & Format(rs("ReferenceDate"), "yyyy-MM-dd") & "',"
    Sql = Sql & " '" & Format(rs("ValidFrom"), "yyyy-MM-dd") & "',"
    Sql = Sql & " '" & Format(rs("ValidTo"), "yyyy-MM-dd") & "',"
    Sql = Sql & " '" & rs("OrderingType") & "',"
    Sql = Sql & " '" & rs("Remarks") & "',"
    Sql = Sql & " " & Val(rs("FormCode")) & ","
    Sql = Sql & " " & Val(rs("PackingCode")) & ","
    Sql = Sql & " " & Val(rs("DeliveryCode")) & ","
    Sql = Sql & " " & Val(rs("PaymentCode")) & ","
    Sql = Sql & " " & Val(rs("InsuranceCode")) & ","
    Sql = Sql & " " & Val(rs("ModeCode")) & ","
    Sql = Sql & " " & Val(rs("Freightcode")) & ","
    Sql = Sql & " " & Val(rs("FormulaCode")) & ","
    Sql = Sql & " " & Val(rs("TotalAmount")) & ","
    Sql = Sql & " " & Val(rs("PostingAc")) & ","
    Sql = Sql & " " & Val(rs("RevisionLevel")) & ","
    Sql = Sql & " " & Val(rs("GeneralDetails")) & ","
    Sql = Sql & " '" & Trim$(rs("ShorClose")) & "',"
    Sql = Sql & " '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
    Sql = Sql & " '" & rs("OrderAuthorization") & "',"
    Sql = Sql & " '" & Format(rs("OrderAuthorizationDate"), "yyyy-MM-dd HH:mm:ss") & "',"
    Sql = Sql & " '" & rs("OrderAuthorizationBy") & "',"
    Sql = Sql & " '" & gUserID & "',"
    Sql = Sql & " '" & gBranchID & "',"
    Sql = Sql & " '" & gCompID & "',"
    Sql = Sql & " '" & gFinyear & "',"
    Sql = Sql & " '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    Sql = Sql & " )"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If

  SaveHeaderData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub cmdSave_Click()
  Dim Sql As String
  Dim rCnt As Long
  Dim iCnt As Long
  Dim retVal As Variant
  Dim CurDate As Date
  Dim schk As Variant
  
  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  CurDate = ServerDate
  
  For rCnt = 1 To sprOrder.DataRowCnt
    sprOrder.GetText eOrder.mShortClose, rCnt, schk
    If Val(schk) = 1 Then
      If FrmFlg = "S" Then
        If SaveHeaderData(CurDate, rCnt) = False Then
          GoTo ExitHere
        End If
      ElseIf FrmFlg = "J" Then
        If SaveJobworkData(CurDate, rCnt) = False Then
          GoTo ExitHere
        End If
      End If
    End If
  
'    sprItem.Sheet = rCnt
'    For iCnt = 1 To sprItem.DataRowCnt
''      sprItem.GetText eItem.mShortClose, iCnt, schk
''      If Val(schk) = 1 Then
'        If SaveDetailData(CurDate, rCnt, iCnt) = False Then
'          GoTo ExitHere
'        End If
'
'        If SaveScheduleData(CurDate, rCnt, iCnt) = False Then
'          GoTo ExitHere
'        End If
''      End If
'    Next iCnt
  Next rCnt
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Order(s) ShortClosed/Cancelled Successfully"
  Call cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
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

Private Sub sprOrder_ButtonClicked(sender as object, e as ButtonClickedEventArgs)
  Dim retVal As Integer
  Dim i As Integer
    
  Try
  Me.MousePointer = vbHourglass

  If sButtonClick = True Then
      If Col = eOrder.mShortClose Then
        sprOrder.Col = Col
sprOrder.Row = Row
        If sprOrder.Value = 1 Then
          sprItem.Sheet = Row
          For i = 1 To sprItem.DataRowCnt
            sprItem.Col = eItem.mShortClose
sprItem.Row = i
sprItem.Value = 1
sprItem.Lock = True
          Next i
        Else
          sprItem.Sheet = Row
          For i = 1 To sprItem.DataRowCnt
            sprItem.Col = eItem.mShortClose
sprItem.Row = i
sprItem.Value = 0
sprItem.Lock = False
          Next i
        End If
      End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_ButtonClicked(sender as object, e as ButtonClickedEventArgs)
  Dim retVal As Variant
  Dim schk As Long
  Dim Cnt As Integer
  Dim i As Integer
  
  Try
  Me.MousePointer = vbHourglass

  If sButtonClick = False Then
       If sprItem.ActiveCol = eItem.mShortClose Then
         sprItem.Sheet = sprItem.ActiveSheet
         Cnt = 0
           For i = 1 To sprItem.DataRowCnt
             sprItem.GetInteger eItem.mShortClose, i, schk
             If Val(schk) = 1 Then
               Cnt = Cnt + 1
             End If
           Next i
           If Cnt = sprItem.DataRowCnt Then
             sprOrder.Col = eOrder.mShortClose
sprOrder.Row = sprItem.ActiveSheet
sprOrder.Value = 1
           Else
             sprOrder.Col = eOrder.mShortClose
sprOrder.Row = sprItem.ActiveSheet
sprOrder.Value = 0
           End If
       End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  sButtonClick = False
End Sub

Private Sub SetSheet(ByVal fRow As Long)
' On Error Resume Next
  
    If sprItem.SheetCount < fRow Then
      sprItem.SheetCount = fRow
      Call CopySpreadStructure(sprItem, 1, , fRow)
    End If
    sprItem.Sheet = fRow
    sprItem.ActiveSheet = fRow
  
End Sub

Private Sub OrderRowChangeOperation(ByVal fRow As Long)
  Dim retVal As Variant
  
  Try
  
  If fRow > 0 Then
    sprOrder.GetText eOrder.mVendor, fRow, retVal
    txtVendor.Text = Trim$(retVal)
    sprItem.Sheet = fRow
    sprItem.ActiveSheet = fRow
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdClear_Click()

  Try
  Me.MousePointer = vbHourglass

  sprItem.SheetCount = 1
  sprItem.MaxRows = 1
  sprOrder.MaxRows = 1
    
  Call ClearForm(Me)
  
  fraDate.Enabled = False
  cmdsave.Enabled = False
  StabItem.Enabled = False
  fraDate.Enabled = True
  StabItem.Tab = 0
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rsDet As AceADODB.Recordset
  Dim OrderNo As String
  Dim CurDate As Date
  Dim rCnt As Long   'Order Row Count
  Dim irCnt As Long  'Item Row Count
  Dim Cnt As Integer
  Dim PendingQty As Double

  Try
  Me.MousePointer = vbHourglass
  
  CurDate = ServerDate

  If dtpToDate.Value < dtpFromDate.Value Then
    sShowMessage "To Date should be less than or equal to From Date."
    dtpToDate.SetFocus
    GoTo ExitHere
  End If
  
  If FrmFlg = "S" Then
    Sql = " Select Distinct a.OrderType, a.OrderNo, a.OrderDate, a.CreatedDate, a.CustomerCode, b.PartyName,"
    Sql = Sql & " Case a.Status when 'N' then 'New' when 'M' then 'Amendment' when 'A' then 'Authorised'"
    Sql = Sql & " when 'H' then 'Hold' End as Status, a.OrderingType" ',c.OrderQty
    Sql = Sql & " From  dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " Left Outer Join FAS_Party_Master b"
    Sql = Sql & " On b.PartyCode = a.CustomerCode and b.BranchID = a.BranchID and b.CompID = a.COmpID"
    Sql = Sql & " where a.Status not in ('C','S','H')"
    Sql = Sql & " and a.orderDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " and Exists(Select Distinct OrderType, OrderNo, OrderDate"
    Sql = Sql & " From (Select a.OrderType, a.OrderNo, a.OrderDate, a.OrderQty,  isNull(Sum(c.InvoiceQty), 0) as InvQty" ', Sum(d.DCQty) as DCQty"
    Sql = Sql & " From dbo.fn_SAL_Order_Dtl('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " Left Outer Join SAL_Invoice_Dtl c"
    Sql = Sql & " On a.OrderNo = c.OrderNo and a.OrderDate = c.OrderDate and a.OrderType = c.OrderType"
    Sql = Sql & " and a.ItemCode = c.ItemCode and a.BranchId = c.BranchID and a.CompID = c.CompID"
    Sql = Sql & " Where a.Status Not in ('C','S','H') and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'and a.OrderQty > 0"
    Sql = Sql & " and a.OrderDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " Group by a.OrderType, a.OrderNo, a.OrderDate, a.OrderQty, a.ItemCode"
    Sql = Sql & " ) x Where OrderQty > InvQty and a.OrderType=x.OrderType and a.OrderNo=x.OrderNo and a.OrderDate=x.OrderDate)"
    Sql = Sql & " and a.OrderingType='Specific' "
  ElseIf FrmFlg = "J" Then
    Sql = " Select Distinct a.OrderType, a.OrderNo, a.OrderDate, a.CreatedDate, a.CustomerCode, Case a.Status when 'N' then 'New' when 'M' then 'Amendment' when 'A' "
    Sql = Sql & " then 'Authorised' when 'H' then 'Hold' End as Status, a.OrderingType,b.PartyName"
    Sql = Sql & " From dbo.fn_SAL_JobworkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " Left Outer Join FAS_Party_Master b"
    Sql = Sql & " On b.PartyCode = a.CustomerCode and b.BranchID = a.BranchID and b.CompID = a.COmpID"
    Sql = Sql & " Where a.Status not in ('C','S','H') and a.OrderDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " and Exists(Select Distinct OrderType, OrderNo, OrderDate"
    Sql = Sql & " From (Select a.OrderType, a.OrderNo, a.OrderDate, a.OrderQty,  isNull(Sum(c.InvoiceQty), 0) as InvQty"
    Sql = Sql & " From dbo.fn_SAL_JobworkOrder_Dtl('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " Left Outer Join SAL_Invoice_Dtl c On a.OrderNo = c.OrderNo and a.OrderDate = c.OrderDate"
    Sql = Sql & " and a.OrderType = c.OrderType and a.SentItemCode = c.ItemCode and a.BranchId = c.BranchID"
    Sql = Sql & " and a.CompID = c.CompID Where a.Status Not in ('C','S','H') and a.Branchid='" & gBranchID & "'"
    Sql = Sql & " and a.Compid='" & gCompID & "'and a.OrderQty > 0 and a.OrderDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " Group by a.OrderType, a.OrderNo, a.OrderDate, a.OrderQty, a.SentItemCode"
    Sql = Sql & " ) x"
    Sql = Sql & " Where OrderQty > InvQty And a.OrderType = x.OrderType And a.OrderNo = x.OrderNo"
    Sql = Sql & " and a.OrderDate=x.OrderDate) and a.OrderingType='Specific'"
  End If
  
  Set rsDet = New AceADODB.Recordset
  If RSOpen(rsDet, Sql) = False Then
 GoTo ExitHere
End If
  
  If rsDet.EOF Then
    sShowMessage "No Order Details Found"
    GoTo ExitHere
  End If
  
  If FrmFlg = "S" Then
    Sql = " Select Distinct a.OrderType, a.OrderNo, a.OrderDate, a.ItemCode, d.ItemDescription,"
    Sql = Sql & " a.Rate, a.AssessableRate as AssRate, d.UOMdesc, d.UOMDecimal, a.OrderQty, Sum(e.InvoiceQty) as CQty "
    Sql = Sql & " From dbo.fn_SAL_Order_dtl('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " Left Outer Join Material_Master d"
    Sql = Sql & " On a.ItemCode = d.ItemCode and a.BranchID = d.BranchID and a.CompID = d.CompId"
    Sql = Sql & " Left Outer Join SAL_Invoice_Dtl e "
    Sql = Sql & " On a.OrderDate = e.OrderDate and a.OrderNo = e.OrderNo"
    Sql = Sql & " and a.OrderType = e.OrderType and a.ItemCode = e.ItemCode  and a.BranchID = e.BranchID and a.CompID = e.CompID"
    Sql = Sql & " where a.Status not in ('C','S','H') "
    Sql = Sql & " and Exists(Select Distinct OrderType, OrderNo, OrderDate"
    Sql = Sql & " From (Select a.OrderType, a.OrderNo, a.OrderDate, a.ItemCode, a.OrderQty,  isNull(Sum(c.InvoiceQty), 0) as InvQty" ', Sum(d.DCQty) as DCQty"
    Sql = Sql & " From dbo.fn_SAL_Order_Dtl('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " Left Outer Join SAL_Invoice_Dtl c"
    Sql = Sql & " On a.OrderNo = c.OrderNo and a.OrderDate = c.OrderDate and a.OrderType = c.OrderType"
    Sql = Sql & " and a.ItemCode = c.ItemCode and a.BranchId = c.BranchID and a.CompID = c.CompID"
    Sql = Sql & " Where a.Status Not in ('C','S','H')and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'and a.OrderQty > 0"
    Sql = Sql & " and a.OrderDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " Group by a.OrderType, a.OrderNo, a.OrderDate, a.OrderQty, a.ItemCode"
    Sql = Sql & " ) x Where OrderQty > InvQty and a.OrderType=x.OrderType and a.OrderNo=x.OrderNo and a.OrderDate=x.OrderDate and a.ItemCode=x.ItemCode)"
    Sql = Sql & " Group by a.OrderType, a.OrderNo, a.OrderDate, a.ItemCode, d.ItemDescription, a.OrderQty, a.Rate, a.AssessableRate, d.UOMdesc, d.UOMDecimal"
    Sql = Sql & " Having Sum(e.InvoiceQty) < a.OrderQty"
  ElseIf FrmFlg = "J" Then
    Sql = " Select Distinct a.OrderType, a.OrderNo, a.OrderDate, a.SentItemCode, a.Rate, a.OrderQty,"
    Sql = Sql & " d.ItemDescription, Sum(e.InvoiceQty)as CQty, a.AssRate, d.UOMdesc, d.UOMDecimal"
    Sql = Sql & " From dbo.fn_SAL_JobworkOrder_Dtl('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " Left Outer Join Material_Master d"
    Sql = Sql & " On a.SentItemCode = d.ItemCode and a.BranchID = d.BranchID and a.CompID = d.CompId"
    Sql = Sql & " Left Outer Join SAL_Invoice_Dtl e"
    Sql = Sql & " On a.OrderDate = e.OrderDate and a.OrderNo = e.OrderNo and a.OrderType = e.OrderType"
    Sql = Sql & " and a.SentItemCode = e.ItemCode  and a.BranchID = e.BranchID and a.CompID = e.CompID"
    Sql = Sql & " where a.Status not in ('C','S','H')"
    Sql = Sql & " and Exists(Select Distinct OrderType, OrderNo, OrderDate"
    Sql = Sql & " From (Select a.OrderType, a.OrderNo, a.OrderDate, a.SentItemCode, a.OrderQty,"
    Sql = Sql & " isNull(Sum(c.InvoiceQty), 0) As InvQty"
    Sql = Sql & " From dbo.fn_SAL_JobworkOrder_Dtl('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
    Sql = Sql & " Left Outer Join SAL_Invoice_Dtl c"
    Sql = Sql & " On a.OrderNo = c.OrderNo and a.OrderDate = c.OrderDate"
    Sql = Sql & " and a.OrderType = c.OrderType and a.SentItemCode = c.ItemCode"
    Sql = Sql & " and a.BranchId = c.BranchID and a.CompID = c.CompID"
    Sql = Sql & " Where a.Status Not in ('C','S','H')and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'and a.OrderQty > 0"
    Sql = Sql & " and a.OrderDate between '" & Format(dtpFromDate, "yyyy-MM-dd") & "' and '" & Format(dtpToDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " Group by a.OrderType, a.OrderNo, a.OrderDate, a.OrderQty, a.SentItemCode ) x"
    Sql = Sql & " Where a.OrderQty > x.InvQty And a.OrderType = x.OrderType And a.OrderNo = x.OrderNo"
    Sql = Sql & " and a.OrderDate = x.OrderDate and a.SentItemCode=x.SentItemCode)"
    Sql = Sql & " Group by a.OrderType, a.OrderNo, a.OrderDate, a.SentItemCode, d.ItemDescription,"
    Sql = Sql & " a.OrderQty , a.Rate, a.AssRate, d.UOMdesc, d.UOMDecimal"
    Sql = Sql & " Having Sum(e.InvoiceQty) < a.OrderQty"
  End If
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsDet.EOF Then
    rsDet.MoveLast
rsDet.MoveFirst
      sprOrder.MaxRows = rsDet.RecordCount
      For rCnt = 1 To rsDet.RecordCount
          sprOrder.SetText eOrder.mOrdType, rCnt, rsDet("OrderType")
          sprOrder.SetText eOrder.mOrdNo, rCnt, rsDet("OrderNo")
          sprOrder.SetText eOrder.mOrdDate, rCnt, Format(rsDet("OrderDate"), "dd/MM/yyyy")
          sprOrder.SetText eOrder.mCurStatus, rCnt, rsDet("Status")
          sprOrder.SetText eOrder.mStatusDate, rCnt, Format(rsDet("CreatedDate"), "dd/MM/yyyy")
          sprOrder.SetText eOrder.mVendor, rCnt, rsDet("PartyName")
          sprOrder.SetText eOrder.mOrderingType, rCnt, rsDet("OrderingType")
          sprOrder.SetText eOrder.mCreatedDate, rCnt, Format(rsDet("CreatedDate"), "yyyy-MM-dd HH:mm:ss")
          sprOrder.SetText eOrder.mStatusDateKey, rCnt, Format(rsDet("CreatedDate"), "yyyy-MM-dd HH:mm:ss")
          
          rs.Filter = adFilterNone
          rs.Filter = "OrderNo=" & rsDet("OrderNo") & " and OrderType='" & rsDet("OrderType") & "' and OrderDate='" & Format(rsDet("OrderDate"), "yyyy-MM-dd") & "'"
          
          Call SetSheet(rCnt)
          If Not rs.EOF Then
            rs.MoveLast
rs.MoveFirst
            With sprItem
              sprOrder.MaxRows = rs.RecordCount
              For irCnt = 1 To rs.RecordCount
                If FrmFlg = "S" Then
                  sprOrder.SetText eItem.mItem, irCnt, rs("ItemCode")
                Else
                  sprOrder.SetText eItem.mItem, irCnt, rs("SentItemCode")
                End If
                sprOrder.SetText eItem.mDesc, irCnt, rs("ItemDescription")
                sprOrder.SetText eItem.mUOM, irCnt, rs("UOMDesc")
                Call Spread_Set_Decimal(sprItem, Val(CheckNull(rs("UOMDecimal"))), eItem.mOrdQty, irCnt, eItem.mRatePer)
                sprOrder.SetText eItem.mOrdQty, irCnt, IIf(rs("OrderQty") > 0, rs("OrderQty"), "")
                sprOrder.SetText eItem.mInvoiceQty, irCnt, rs("CQty")
                PendingQty = rs("OrderQty") - CheckNull(rs("CQty"))
                sprOrder.SetText eItem.mPendQty, irCnt, PendingQty
                sprOrder.SetText eItem.mRate, irCnt, rs("Rate")
                sprOrder.SetText eItem.mAssRate, irCnt, rs("AssRate")
                rs.MoveNext
              Next irCnt
              sprItem.MaxRows = rs.RecordCount
              For irCnt = 1 To rs.RecordCount
                If FrmFlg = "S" Then
                  sprItem.SetText eItem.mItem, irCnt, rs("ItemCode")
                Else
                  sprItem.SetText eItem.mItem, irCnt, rs("SentItemCode")
                End If
                sprItem.SetText eItem.mDesc, irCnt, rs("ItemDescription")
                sprItem.SetText eItem.mUOM, irCnt, rs("UOMDesc")
                Call Spread_Set_Decimal(sprItem, Val(CheckNull(rs("UOMDecimal"))), eItem.mOrdQty, irCnt, eItem.mRatePer)
                sprItem.SetText eItem.mOrdQty, irCnt, IIf(rs("OrderQty") > 0, rs("OrderQty"), "")
                sprItem.SetText eItem.mInvoiceQty, irCnt, rs("CQty")
                PendingQty = rs("OrderQty") - CheckNull(rs("CQty"))
                sprItem.SetText eItem.mPendQty, irCnt, PendingQty
                sprItem.SetText eItem.mRate, irCnt, rs("Rate")
                sprItem.SetText eItem.mAssRate, irCnt, rs("AssRate")
                rs.MoveNext
              Next irCnt
          Else
            sprItem.MaxRows = 1
          End If
        rsDet.MoveNext
      Next rCnt
    End With
    Call OrderRowChangeOperation(1)
  End If
  
  sprOrder.Enabled = True
  cmdsave.Enabled = True
  StabItem.Enabled = True
  fraDate.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set rsDet = Nothing
'  Set rsChk = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurDate As Date
  
  Try
  
  CurDate = ServerDate

  dtpFromDate = Format(CurDate, "dd-MM-yyyy")
  dtpToDate = dtpFromDate
   
  sprOrder.Enabled = False
  
  If FrmFlg = "J" Then
    frmSalesOrderShortClose.Caption = "Jobwork Short Close"
  ElseIf FrmFlg = "S" Then
    frmSalesOrderShortClose.Caption = "SalesOrder Short Close"
  End If
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

Private Sub sprOrder_Click(sender as object, e as ClickEventArgs)
' On Error Resume Next
  sButtonClick = True
End Sub

Private Sub StabItem_Click(ByVal PreviousTab As Integer)
'  Dim Sql As String
'  Dim retVal As Variant
'  Dim i As Integer
'  Dim dtpCurDate As Date
'  Dim rs As ADODB.Recordset
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  dtpCurDate = ServerDate
'  sprSchedule.ClearRange -1, -1, sprSchedule.MaxCols, sprSchedule.MaxRows, True
'
'  If PreviousTab = 0 Then
'    sprItem.Sheet = sprItem.ActiveSheet
'    Sql = " Select b.OrderType, b.OrderNo, b.OrderDate, b.ItemCode,a.OrderQty, b.ScheduleQty, Sum(c.InvoiceQty) as InvQty"
'    Sql = Sql & " From dbo.fn_SAL_Order_Dtl('" & Format(dtpCurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(dtpCurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')a"
'    Sql = Sql & " Left Outer Join dbo.fn_SAL_Order_Schedule('" & Format(dtpCurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format(dtpCurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') b"
'    Sql = Sql & " On a.OrderDate = b.OrderDate and a.OrderNo = b.OrderNo and a.OrderType =b.OrderType"
'    Sql = Sql & " and a.ItemCode = b.ItemCode and a.BranchID = b.BranchID and a.CompID = b.COmpID"
'    Sql = Sql & " Left Outer Join  SAL_Invoice_Dtl c"
'    Sql = Sql & " On a.OrderDate = c.OrderDate and a.OrderNo = c.OrderNo and a.OrderType = c.OrderType"
'    Sql = Sql & " and a.ItemCode = c.ItemCode and a.BranchID = c.BranchID and a.CompID = c.CompID"
'    Sql = Sql & " Where a.OrderQty > 0 and b.ScheduleQty >0 and a.Status not in('C','S')"
'    sprItem.GetText eItem.mItem, sprItem.ActiveRow, retVal
'    Sql = Sql & " and a.ItemCode='" & Trim$(retVal) & "' "
'    sprOrder.GetText eOrder.mOrdType, sprOrder.ActiveRow, retVal
'    Sql = Sql & " and a.OrderType='" & Trim$(retVal) & "'"
'    sprOrder.GetText eOrder.mOrdNo, sprOrder.ActiveRow, retVal
'    Sql = Sql & " and a.OrderNo=" & Val(retVal) & ""
'    sprOrder.GetText eOrder.mOrdDate, sprOrder.ActiveRow, retVal
'    Sql = Sql & " and a.OrderDate='" & Format(retVal, "yyyy-MM-dd") & "'"
'    Sql = Sql & " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
'    Sql = Sql & " Group By  b.ScheduleType, b.ScheduleNo, b.ScheduleDate, b.ItemCode,a.OrderQty, b.ScheduleQty"
'    Sql = Sql & " Order by ScheduleNo"
'    Set rs = New ADODB.Recordset
'    If RSOpen(rs, Sql) = False Then: GoTo ExitHere
'    With sprSchedule
'      If Not rs.EOF Then
'        .MaxRows = rs.RecordCount
'        For i = 1 To rs.RecordCount
'          .SetText eSch.mSchType, i, rs("ScheduleType")
'          .SetText eSch.mSchNo, i, rs("ScheduleNo")
'          .SetText eSch.mSchDate, i, rs("ScheduleDate")
'          .SetText eSch.mItemCode, i, rs("ItemCode")
'          .SetText eSch.mSchQty, i, rs("ScheduleQty")
'          .SetText eSch.mComQty, i, rs("InvQty")
'          If Val(rs("ScheduleQty")) = Val(rs("InvQty")) Then
'            .SetText eSch.mComStatus, i, "1"
'          Else
'            .SetText eSch.mComStatus, i, "0"
'          End If
'          rs.MoveNext
'        Next i
'      Else
'        .MaxRows = 1
'      End If
'    End With
'  End If
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Sub

Private Sub Design_sprItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprItem


sprItem.ButtonClicked +=  new EventHandler(sprItem_ButtonClicked)
sprItem.Click +=  new EventHandler(sprItem_Click)

End Sub

Private Sub Design_sprOrder(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprOrder


sprOrder.LeaveCell +=  new EventHandler(sprOrder_LeaveCell)
sprOrder.ButtonClicked +=  new EventHandler(sprOrder_ButtonClicked)
sprOrder.Click +=  new EventHandler(sprOrder_Click)

End Sub

Private Sub Design_sprSchedule(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprSchedule

End Sub

Private Sub InitializeSpreads()

	Call Design_sprItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmSalesOrderShortClose.frm", sprItem, "")

	Call Design_sprOrder("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmSalesOrderShortClose.frm", sprOrder, "")

	Call Design_sprSchedule("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmSalesOrderShortClose.frm", sprSchedule, "")

End Sub
