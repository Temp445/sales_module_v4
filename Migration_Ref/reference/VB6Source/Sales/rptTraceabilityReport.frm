VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form rptTraceabilityReport 
   Caption         =   "Traceability Report"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   15735
   ScaleWidth      =   28680
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLegend 
      Caption         =   "Legend"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1110
      Left            =   615
      TabIndex        =   32
      Top             =   2145
      Width           =   7620
      Begin VB.Label lblLegend 
         AutoSize        =   -1  'True
         Caption         =   $"rptTraceabilityReport.frx":0000
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   675
         Left            =   135
         TabIndex        =   33
         Top             =   300
         Width           =   7410
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11760
      TabIndex        =   30
      Top             =   2955
      Width           =   510
   End
   Begin VB.Frame fraReportFormat 
      Enabled         =   0   'False
      Height          =   5550
      Left            =   -40
      TabIndex        =   28
      Top             =   3390
      Width           =   20070
Begin AceSpread SprPrint
         Height          =   5190
         Left            =   90
         TabIndex        =   31
         Top             =   225
         Width           =   19875
         _Version        =   458752
         _ExtentX        =   35057
         _ExtentY        =   9155
         _StockProps     =   64
         DisplayColHeaders=   0   'False
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
         GridShowHoriz   =   0   'False
         GridShowVert    =   0   'False
      End
   End
   Begin VB.Frame fraSearch 
      Height          =   2160
      Index           =   0
      Left            =   12450
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   7545
Begin AceSpread sprList_0
         Height          =   1800
         Left            =   105
         TabIndex        =   27
         Top             =   240
         Width           =   7320
         _Version        =   458752
         _ExtentX        =   12912
         _ExtentY        =   3175
         _StockProps     =   64
         DAutoSizeCols   =   1
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
         MaxRows         =   5
         OperationMode   =   4
         SelectBlockOptions=   0
      End
   End
   Begin VB.CommandButton cmdStep 
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
      Left            =   11445
      TabIndex        =   23
      Top             =   2205
      Width           =   345
   End
   Begin VB.Frame fraStep 
      Caption         =   "Item "
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   585
      Index           =   1
      Left            =   8683
      TabIndex        =   20
      Top             =   2040
      Width           =   2625
      Begin VB.OptionButton OptAll 
         Caption         =   "&All"
         Height          =   210
         Left            =   195
         MaskColor       =   &H00400000&
         TabIndex        =   22
         Top             =   255
         Value           =   -1  'True
         Width           =   570
      End
      Begin VB.OptionButton OptSelective 
         Caption         =   "&Selective"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1545
         TabIndex        =   21
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   10635
      TabIndex        =   14
      Top             =   9075
      Width           =   1300
   End
   Begin VB.Frame fraPreRequest 
      Height          =   1770
      Left            =   3255
      TabIndex        =   16
      Top             =   11625
      Visible         =   0   'False
      Width           =   2295
      Begin VB.TextBox txtItemUserKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   45
         TabIndex        =   24
         Top             =   1365
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker dtpAsOnDate 
         Height          =   315
         Left            =   45
         TabIndex        =   17
         Top             =   330
         Visible         =   0   'False
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   372375555
         CurrentDate     =   39743
      End
      Begin VB.Label lblItemUserKey 
         AutoSize        =   -1  'True
         Caption         =   "Item UserKey"
         Height          =   210
         Left            =   75
         TabIndex        =   25
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label lblAsOnDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AsOnDate"
         Height          =   210
         Left            =   45
         TabIndex        =   19
         Top             =   135
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblTagCustomerCode 
         AutoSize        =   -1  'True
         Caption         =   "CustomerCode"
         Height          =   210
         Left            =   120
         TabIndex        =   18
         Top             =   750
         Width           =   1065
      End
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   9345
      TabIndex        =   15
      Top             =   9075
      Width           =   1300
   End
   Begin VB.CommandButton cmdStep 
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
      Left            =   11535
      TabIndex        =   13
      Top             =   1590
      Width           =   345
   End
   Begin VB.Frame fraDocDtl 
      Enabled         =   0   'False
      Height          =   780
      Left            =   8398
      TabIndex        =   4
      Top             =   2625
      Width           =   3195
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label lblCustomer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3210
         TabIndex        =   12
         Top             =   180
         Width           =   690
      End
      Begin VB.Label lblDocType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc.Type"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   10
         Top             =   180
         Width           =   690
      End
      Begin VB.Label lblDocNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc.No."
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   990
         TabIndex        =   9
         Top             =   180
         Width           =   570
      End
      Begin VB.Label lblDocDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Doc. Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   2040
         TabIndex        =   8
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.Frame fraStep 
      Caption         =   "Period"
      Height          =   630
      Index           =   0
      Left            =   8540
      TabIndex        =   0
      Top             =   1380
      Width           =   2910
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   315
         Left            =   1605
         TabIndex        =   1
         Top             =   210
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   372375553
         CurrentDate     =   37273
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   210
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Format          =   372375553
         CurrentDate     =   37273
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   210
         Left            =   1395
         TabIndex        =   3
         Top             =   255
         Width           =   180
      End
   End
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export To Excel"
      Enabled         =   0   'False
      Height          =   400
      Left            =   8055
      TabIndex        =   29
      Top             =   9075
      Width           =   1300
   End
   Begin MSComDlg.CommonDialog dlgMain 
      Left            =   1365
      Top             =   11850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "rptTraceabilityReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const DefHdrRow As Long = 6
Private sprow As Long

Private Enum eRptTraceable
  mItemCode = 1
  mItemDesc = 2
  mUOM = 3
  mDocNo = 4
  mTranDate = 5
  mWarehouse = 6
  mQty = 7
  mRemarks = 8
End Enum

Private Sub cmdClear_Click()
  Dim CurDate As Date
  
  Try
    
  'Delete ItemCode into Tmp table
  Call DeleteTmpTable
  
  Call ClearForm(Me)
  
  'Getting Server Date
  CurDate = SetServerDate(dtpAsOnDate)
  
  If CurDate < FinStart Then
    GoTo ExitHere
  End If
  
  'Assign ServerDate
  dtpAsOnDate.Value = Format(CurDate, "dd/MM/yyyy HH:mm:ss")
  
  'Assigning Server Date
  dtpFromDate.MinDate = FinStart
  dtpFromDate.MaxDate = Format(CurDate, "dd/MM/yyyy")
  dtpFromDate.Value = Format(CurDate, "dd/MM/yyyy")
  
  dtpToDate.MinDate = FinStart
  dtpToDate.MaxDate = Format(CurDate, "dd/MM/yyyy")
  dtpToDate.Value = Format(CurDate, "dd/MM/yyyy")
    
  Call Spread_Clear(SprPrint)
SprPrint.MaxCols = eRptTraceable.mRemarks
SprPrint.MaxRows = 1
  
  CmdExport.Enabled = False
  OptAll.Value = True
  fraDocDtl.Enabled = False
  
  Call StepClear(Me)
  Call StepEnable(Me)
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

Private Sub CmdExport_Click()
  Try
  Me.MousePointer = vbHourglass

  If Spread_Export(SprPrint, dlgMain, "Traceability Report") = False Then
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass
          
  If sSpread_sPrintcurrent = False Then
 GoTo ExitHere
End If
Catch ex As Exception
  Call ErrMsgShow(ex)
Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
' On Error Resume Next

  If Index = 0 Then
    If DateValue(dtpToDate.Value) < DateValue(dtpFromDate.Value) Then
      sShowMessage "To Date Should be greater than From Date"
      dtpToDate.SetFocus
      Exit Sub
    End If
  Else
    
    If OptSelective Then
      If sprList(0).SelectionCount < 1 Then
        sShowMessage "Please select atlease one Item"
        Call Spread_Set_Focus(sprList(0), 1, 1)
       Exit Sub
      End If
    End If
    fraSearch(0).Enabled = False
    fraDocDtl.Enabled = True
  End If
  
  Call StepProceedNext(Me, Index)
End Sub

Private Function sSpread_sPrintcurrent() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Dim RptFrm As Form
  
  Dim rCnt As Long
  
  Try
  sSpread_sPrintcurrent = False
  
  If contrl_valid(txtType, "Please select Doc Type") = False Then
 GoTo ExitHere
End If
  
  Call SelLstFill
  
  Sql = " Execute SAL_Rpt_TraceabilityReport_Gen"
  Sql = Sql & " @ItemUserKey='" & Trim$(txtItemUserKey.Text) & "',"
  Sql = Sql & " @DocType='" & Trim$(txtType.Text) & "',"
  Sql = Sql & " @DocNo=" & Val(txtNo.Text) & ","
  Sql = Sql & " @DocDate='" & Format(txtDate.Text, "yyyy-MM-dd") & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "', "
  Sql = Sql & " @CompID='" & gCompID & "',"
  Sql = Sql & " @FinYear='" & gFinyear & "'"

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then GoTo ExitHere

  If rs.RecordCount = 0 Then
    MsgBox "No Record(s) Found"
    GoTo ExitHere
  End If
  
  Call Spread_Clear(SprPrint)
SprPrint.MaxCols = 1
SprPrint.MaxRows = 1
  fraReportFormat.Enabled = True
  
   
    sprow = 1
SprPrint.MaxRows = rs.RecordCount + DefHdrRow + 1
SprPrint.MaxCols = eRptTraceable.mRemarks
    Call spSetColWidth(SprPrint, eRptTraceable.mItemCode, 15, eRptTraceable.mItemDesc, 50, eRptTraceable.mUOM, 5, eRptTraceable.mDocNo, 17, eRptTraceable.mTranDate, 17, eRptTraceable.mWarehouse, 12, eRptTraceable.mQty, 12, eRptTraceable.mRemarks, 35)
    
    SprPrint.MaxRows = rs.RecordCount + DefHdrRow + 1
    
    SprPrint.MaxCols = eRptTraceable.mRemarks
    
    Call Header       'Printing the Header Details
    
    If Not rs.EOF() Then
      rCnt = 1
      Do While Not rs.EOF
        'Border
        Call spSetCellBorder(SprPrint, eRptTraceable.mItemCode, sprow, SprPrint.MaxCols, , mBorderAllSides)
        Call spSetTextColArray(SprPrint, eRptTraceable.mItemCode, sprow, 1, mAlignLeftCenter, False, rs("ItemCode"), rs("ItemDesc"), rs("UOMDesc"), rs("DocType") & " - " & rs("DocNo") & " - " & rs("DocDate"), Format(rs("TransactionDate"), "dd/MM/yyyy HH:mm:ss"), rs("ToWarehouseCode"), "", rs("HeatNo"))
        Call spSetText(SprPrint, eRptTraceable.mQty, sprow, Format(rs("Qty"), FormatDecimalString(rs("UOMDecimal"))), mAlignRightCenter)

        'Set Row Height
        SprPrint.RowHeight(sprow) = SprPrint.MaxTextRowHeight(sprow)
        sprow = sprow + 1
        
        rs.MoveNext
      Loop

    End If
    
    SprPrint.MaxRows = SprPrint.DataRowCnt
    CmdExport.Enabled = True
  
  'Delete ItemCode into Tmp table
  Call DeleteTmpTable
  
  sSpread_sPrintcurrent = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Set RptFrm = Nothing
End Try
End Function

Private Sub Header()

  Try

    
    'Printed Date
    Call spSetText(SprPrint, eRptTraceable.mItemCode, sprow, "Printed Date : " & Format(dtpAsOnDate.Value, "dd/MM/yyyy HH:mm:ss"), mAlignleftTop, True, SprPrint.MaxCols, 1)
    sprow = sprow + 1
    
    'Company Name
    SprPrint.Col = eRptTraceable.mItemCode
SprPrint.Row = sprow:
SprPrint.Col2 = SprPrint.MaxCols
SprPrint.Row2 = sprow + 1
SprPrint.FontSize = 14
    Call spSetText(SprPrint, eRptTraceable.mItemCode, sprow, gCompName, mAlignCenterCenter, True, SprPrint.MaxCols, 1)
    sprow = sprow + 2

    'Heading
    SprPrint.Col = eRptTraceable.mItemCode
SprPrint.Row = sprow
SprPrint.FontSize = 12
    Call spSetText(SprPrint, eRptTraceable.mItemCode, sprow, "Traceability Report [" & Format(dtpFromDate.Value, "dd/mm/yyyy") & " To " & Format(dtpToDate.Value, "dd/mm/yyyy") & "]", mAlignCenterCenter, True, SprPrint.MaxCols, 1)
    sprow = sprow + 1
    SprPrint.AddCellSpan 1, sprow, SprPrint.MaxCols, 1
    sprow = sprow + 1

    sprow = DefHdrRow
    
    Call spSetTextColArray(SprPrint, eRptTraceable.mItemCode, sprow, 1, mAlignCenterCenter, True, "Item Code", "Item Description", "UOM", "Doc No", "Tran Date", "Warehouse", "Qty", "Remarks")
    
    'BackColor
    Call Spread_BkColour(SprPrint, &H8000000F, eRptTraceable.mItemCode, sprow, SprPrint.MaxCols, sprow)
    Call spSetCellBorder(SprPrint, eRptTraceable.mItemCode, sprow, SprPrint.MaxCols, , mBorderAllSides)
    
    sprow = sprow + 1
        
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub DeleteTmpTable()
  Dim Sql As String
  
  Try
    
  Sql = " Execute COM_SelectList_Del"
  Sql = Sql & " @UserKey='" & Trim$(txtItemUserKey.Text) & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "',"
  Sql = Sql & " @CompID='" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub SelLstFill()
  Dim Sql As String
  Dim retVal As Variant
  
  Dim lcnt As Long
  Dim sRow As Long
  
  Try
  
  txtItemUserKey.Text = CreateUserKey
  
  If OptAll = True Then
    Call FillItem(True)
  Else
      For lcnt = 1 To sprList(0).SelectionCount
        sRow = sprList(0).GetMultiSelItem(sRow)
        sprList(0).GetText 1, sRow, retVal
               
        Sql = " Execute COM_SelectList_Ins '" & Trim$(txtItemUserKey.Text) & "','', '" & Trim$(retVal) & "', '" & gBranchID & "', '" & gCompID & "', '" & Format(dtpAsOnDate.Value, "yyyy-MM-dd HH:mm:ss") & "'"
        If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
       Next lcnt
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_Activate()
' On Error Resume Next

  'Assigning FormName to RptFrmFlg variable
  Set RptFrmFlg = rptTraceabilityReport
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Dim CurDate As Date
  
  Try
   
  'Getting Server Date
  CurDate = SetServerDate(dtpAsOnDate)
  
  If CurDate < FinStart Then
    GoTo ExitHere
  End If
  
  'Assign ServerDate
  dtpAsOnDate.Value = Format(CurDate, "dd/MM/yyyy HH:mm:ss")
  
  'Assigning Server Date
  dtpFromDate.MinDate = FinStart
  dtpFromDate.MaxDate = Format(CurDate, "dd/MM/yyyy")
  dtpFromDate.Value = Format(CurDate, "dd/MM/yyyy")
  
  dtpToDate.MinDate = FinStart
  dtpToDate.MaxDate = Format(CurDate, "dd/MM/yyyy")
  dtpToDate.Value = Format(CurDate, "dd/MM/yyyy")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  Else
    MDISales.Toolbar1.Visible = False
  End If
End Sub

Private Sub optAll_Click()
' On Error Resume Next
  
  fraSearch(0).Visible = False
  txtItemUserKey.Text = ""     'Empty must because, Step not used
End Sub

Private Sub OptSelective_Click()

  Try
  Me.MousePointer = vbHourglass
  
  If FillItem(False) = False Then
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Function FillItem(ByVal isAll As Boolean) As Boolean
  Dim rs As AceADODB.Recordset
  Dim Sql As String
    
  Try

  FillItem = False
  
  fraSearch(0).Enabled = True
  sprList(0).ClearSelection

  Sql = " Execute SAL_Rpt_TraceabilityReport_AllSel"
  Sql = Sql & " @ItemUserkey = '" & Trim$(txtItemUserKey.Text) & "',"
  Sql = Sql & " @isSelective = '" & IIf(isAll, 0, 1) & "',"
  Sql = Sql & " @FromDate='" & Format(dtpFromDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @ToDate = '" & Format(dtpToDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @AsOnDate = '" & Format(dtpAsOnDate, "yyyy-MM-dd HH:mm:ss") & "',"
  Sql = Sql & " @BranchID = '" & gBranchID & "', "
  Sql = Sql & " @CompID = '" & gCompID & "',"
  Sql = Sql & " @FinYear='" & gFinyear & "'"
   
  If isAll Then
    If DBExecuteSql(db, Sql) = False Then GoTo ExitHere
  Else
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
      
    If rs.EOF Then
      sShowMessage "Record(s) Not Found."
      OptAll.Value = True
      GoTo ExitHere
    Else
      Call ADORS_Spread_DataFill(rs, sprList(0), True)
      fraSearch(0).Visible = True
     
      sprList(0).MaxRows = sprList(0).DataRowCnt
    End If
    
    Call Spread_Set_Focus(sprList(0), 1, 1)
  
    ReDim SortOrder(0 To sprList(0).MaxCols) As Integer
  End If

  FillItem = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
  End Function

Private Sub txtcustomer_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = " Execute SAL_Rpt_TraceabilityReport_Hlp"
  Sql = Sql & " @ItemUSerKey='" & Trim$(txtItemUserKey.Text) & "',"
  Sql = Sql & " @FromDate='" & Format(dtpFromDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @ToDate = '" & Format(dtpToDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "', "
  Sql = Sql & " @CompID='" & gCompID & "',"
  Sql = Sql & " @FinYear='" & gFinyear & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "111001", "List of Invoice No's", "Invoice Type", "Invoice No", "Invoice Date", "Customer Name") = True Then
    txtType.Text = getvalue(1)
    txtNo.Text = getvalue(2)
    txtDate.Text = getvalue(3)
    lblTagCustomerCode.Caption = getvalue(4)
    txtCustomer.Text = Trim$(getvalue(5))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtcustomer_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  
  If KeyCode = vbKeyF2 Then
    Call txtcustomer_DblClick
  End If
  
End Sub

Private Sub txttype_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Call SelLstFill
  
  Sql = " Execute SAL_Rpt_TraceabilityReport_Hlp"
  Sql = Sql & " @ItemUSerKey='" & Trim$(txtItemUserKey.Text) & "',"
  Sql = Sql & " @FromDate='" & Format(dtpFromDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @ToDate = '" & Format(dtpToDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "', "
  Sql = Sql & " @CompID='" & gCompID & "',"
  Sql = Sql & " @FinYear='" & gFinyear & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "111001", "List of Document's", "Doc Type", "Doc No", "Doc Date", "Customer Name") = True Then
    txtType.Text = getvalue(1)
    txtNo.Text = getvalue(2)
    txtDate.Text = getvalue(3)
'    lblTagCustomerCode.Caption = getvalue(4)
'    txtCustomer.Text = Trim$(getvalue(5))
  End If
  
  'Delete ItemCode into Tmp table
  Call DeleteTmpTable
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtType_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  
  If KeyCode = vbKeyF2 Then
    Call txttype_DblClick
  End If
  
End Sub

Private Sub Design_SprPrint(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'SprPrint

End Sub

Private Sub Design_sprList_0(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprList(0)

End Sub

Private Sub InitializeSpreads()

	Call Design_SprPrint("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptTraceabilityReport.frm", SprPrint, "")

	Call Design_sprList_0("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\rptTraceabilityReport.frm", sprList(0), "0")

End Sub
