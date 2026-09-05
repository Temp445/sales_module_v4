VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form rptLCReport 
   Caption         =   "LC Report"
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame FraTagDetail 
      Caption         =   "Tag Detail"
      Height          =   1035
      Left            =   2310
      TabIndex        =   24
      Top             =   9855
      Visible         =   0   'False
      Width           =   2925
      Begin VB.Label TaglblClosingBal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Balance"
         Height          =   210
         Left            =   1605
         TabIndex        =   28
         Top             =   690
         Width           =   1155
      End
      Begin VB.Label TaglblOPBalance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OP Bal"
         Height          =   210
         Left            =   1710
         TabIndex        =   27
         Top             =   345
         Width           =   480
      End
      Begin VB.Label TaglblLCAmount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LC Amount"
         Height          =   210
         Left            =   240
         TabIndex        =   26
         Top             =   690
         Width           =   795
      End
      Begin VB.Label TaglblCustomerCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
         Height          =   210
         Left            =   210
         TabIndex        =   25
         Top             =   345
         Width           =   1110
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   7785
      TabIndex        =   11
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   6000
      Width           =   1300
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
      Index           =   0
      Left            =   11055
      TabIndex        =   23
      Top             =   3165
      Width           =   345
   End
   Begin VB.Frame fraInvType 
      Caption         =   "Type of Invoice"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6328
      TabIndex        =   21
      Top             =   4260
      Width           =   1635
      Begin VB.ComboBox CmbInvType 
         Height          =   330
         ItemData        =   "rptLCReport.frx":0000
         Left            =   105
         List            =   "rptLCReport.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Select the Type of Invoice"
         Top             =   270
         Width           =   1440
      End
   End
   Begin VB.Frame fraPreRequest 
      Height          =   780
      Left            =   120
      TabIndex        =   18
      Top             =   10110
      Visible         =   0   'False
      Width           =   2160
      Begin MSComCtl2.DTPicker dtpAsOnDate 
         Height          =   315
         Left            =   75
         TabIndex        =   19
         Top             =   345
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy HH:mm:ss"
         Format          =   57868291
         CurrentDate     =   39743
      End
      Begin VB.Label lblAsOnDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AsOnDate"
         Height          =   210
         Left            =   105
         TabIndex        =   20
         Top             =   150
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   6495
      TabIndex        =   16
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   6000
      Width           =   1300
   End
   Begin VB.Frame fraReportType 
      Height          =   945
      Left            =   5578
      TabIndex        =   12
      Top             =   4935
      Width           =   3135
      Begin VB.ComboBox cmbType 
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "rptLCReport.frx":0095
         Left            =   1035
         List            =   "rptLCReport.frx":009F
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   210
         Width           =   1995
      End
      Begin VB.CheckBox chkPrintDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Print Date"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   585
         Value           =   1  'Checked
         Width           =   990
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Type"
         Height          =   210
         Left            =   105
         TabIndex        =   15
         Top             =   255
         Width           =   885
      End
   End
   Begin VB.Frame fraStep 
      Enabled         =   0   'False
      Height          =   765
      Index           =   0
      Left            =   3290
      TabIndex        =   4
      Top             =   2790
      Width           =   7710
      Begin VB.TextBox txtLCDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   345
         Width           =   1080
      End
      Begin VB.TextBox txtCustomerName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3465
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   345
         Width           =   4155
      End
      Begin VB.TextBox txtLCNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   345
         Width           =   2250
      End
      Begin VB.Label lblCustomerName 
         AutoSize        =   -1  'True
         Caption         =   "Customer Name"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   3465
         TabIndex        =   10
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label lblLCDate 
         Caption         =   "LC Date"
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   2370
         TabIndex        =   9
         Top             =   150
         Width           =   660
      End
      Begin VB.Label lblLCNo 
         AutoSize        =   -1  'True
         Caption         =   "LC No"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   8
         Top             =   150
         Width           =   435
      End
   End
   Begin VB.Frame fraInvDate 
      Caption         =   "Invoice Date Restriction"
      Enabled         =   0   'False
      Height          =   675
      Left            =   5705
      TabIndex        =   0
      Top             =   3555
      Width           =   2880
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   255
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57868289
         CurrentDate     =   40621
      End
      Begin MSComCtl2.DTPicker dtpTodate 
         Height          =   315
         Left            =   1605
         TabIndex        =   2
         Top             =   255
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   556
         _Version        =   393216
         Format          =   57868289
         CurrentDate     =   40621
      End
      Begin VB.Label lblTo 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   210
         Left            =   1350
         TabIndex        =   3
         Top             =   315
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdview 
      Caption         =   "&Proceed"
      Enabled         =   0   'False
      Height          =   400
      Left            =   5205
      TabIndex        =   17
      ToolTipText     =   "Click to View the Report"
      Top             =   6000
      Width           =   1300
   End
End
Attribute VB_Name = "rptLCReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private sprow As Long

Dim k As Long

Private Function sPrintCurrent_Print() As Boolean
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  
  Dim TotAmt As Double
  Dim TotAdjAmt As Double
  
  Dim Cnt As Long
  
  Try
  sPrintCurrent_Print = False
  
  Sql = " Execute SAL_Rpt_LCReport_Gen"
  Sql = Sql & " @CustomerCode=" & Val(TaglblCustomerCode.Caption) & ","
  Sql = Sql & " @LCNo='" & Trim$(txtLCNo.Text) & "',"
  Sql = Sql & " @Type='" & Trim$(CmbInvType.Text) & "',"
  Sql = Sql & " @FromDate='" & Format(dtpFromDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @ToDate='" & Format(dtpTodate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "',"
  Sql = Sql & " @CompID='" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
  If snap.EOF Then
    sShowMessage "No Data(s) Found"
    GoTo ExitHere
  End If
    
  sprow = 1
GridCount = 1
iPageNo = 1
pheight = 35
Orient = 2
  
  Cnt = 1
  
    frmPrint.sprPrint.MaxCols = 11
    
    TaglblOPBalance.Caption = snap("OB")
    TotAmt = snap("TotalInvAmt")
    
    Call Header
          
    Do While Not snap.EOF
    
      Call spSetCellBorder(frmPrint.sprPrint, 1, sprow, frmPrint.sprPrint.MaxCols, sprow, mBorderAllSides)
      Call spSetText(frmPrint.sprPrint, 1, sprow, Val(Cnt), mAlignCenterCenter)
      Call spSetTextColArray(frmPrint.sprPrint, 2, sprow, 1, mAlignLeftCenter, False, snap("InvType") & "-" & snap("INVNo"), snap("INVDate"))
      Call spSetTextColArray(frmPrint.sprPrint, 4, sprow, 1, mAlignRightCenter, False, snap("InvoiceQty"), Format(snap("InvAmount"), "0.00"), snap("RRNo"))
      
      If snap("RRNo") <> "" Then
        Call spSetText(frmPrint.sprPrint, 7, sprow, snap("RRDate"), mAlignLeftCenter)
      End If
      
      Call spSetTextColArray(frmPrint.sprPrint, 8, sprow, 1, mAlignRightCenter, False, IIf(snap("ReceivedQty") > 0, Format(snap("ReceivedQty"), "0.00"), ""), IIf(snap("AcceptedQty") > 0, Format(snap("AcceptedQty"), "0.00"), ""), IIf(snap("RejQty") > 0, Format(snap("RejQty"), "0.00"), ""), IIf(snap("AdjustedAmt") > 0, Format(snap("AdjustedAmt"), "0.00"), ""))
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess

      TotAdjAmt = TotAdjAmt + snap("AdjustedAmt")
      Cnt = Cnt + 1
      snap.MoveNext
    Loop
      
    If snap.EOF Then
      'C/B
      TaglblClosingBal.Caption = Val(TaglblOPBalance.Caption) - TotAdjAmt
      Call spSetText(frmPrint.sprPrint, 9, k, " : " & Format(Val(TaglblClosingBal.Caption), "0.00"), mAlignLeftCenter, True)
      
      'Total
      Call spSetCellBorder(frmPrint.sprPrint, 1, sprow, frmPrint.sprPrint.MaxCols, sprow, mBorderBottom)
      Call spSetCellBorder(frmPrint.sprPrint, 4, sprow, frmPrint.sprPrint.MaxCols, sprow, mBorderLeft)
      Call spSetText(frmPrint.sprPrint, 1, sprow, "Total", mAlignRightCenter, True, 4, 1)
      Call spSetText(frmPrint.sprPrint, 5, sprow, Format(TotAmt, "0.00"), mAlignRightCenter, True)
      Call spSetText(frmPrint.sprPrint, 11, sprow, Format(TotAdjAmt, "0.00"), mAlignRightCenter, True)
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess
    End If
      
    frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.DataRowCnt + 1
    frmPrint.Show
    
  totalpages = iPageNo
  
  sPrintCurrent_Print = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Private Sub RowProcess()
' On Error Resume Next
  
  If GridCount >= pheight Then
      GridCount = 1
      frmPrint.sprPrint.MaxRows = frmPrint.sprPrint.MaxRows + pheight + 1
      frmPrint.sprPrint.Row = frmPrint.sprPrint.Row + 1
sprow = sprow + 1
      frmPrint.sprPrint.RowPageBreak = 1
      iPageNo = iPageNo + 1
      Call Header
  End If
  
End Sub

Public Sub Header()
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
    Call spSetText(frmPrint.sprPrint, 1, sprow, "Page Number :" + str(iPageNo), mAlignLeftCenter, False, 2, 1)
         
    If chkPrintDate.Value = 1 Then
      Call spSetText(frmPrint.sprPrint, 9, sprow, "Printed Date : " & Format(ServerDate, "dd/MM/yyyy HH:MM:ss"), mAlignLeftCenter, False, frmPrint.sprPrint.MaxCols, 1)
    End If
    
    sprow = sprow + 1
GridCount = GridCount + 1

    Sql = " Select * from dbo.fn_SAL_Menu_List('" & Trim$(mnuFlg) & "')"
    Set rs = New Recordset
    If RSOpen(rs, Sql) = False Then GoTo ExitHere
    
    If rs.RecordCount > 0 Then
      Call spSetText(frmPrint.sprPrint, 1, sprow, "ERP Ref. No   :" & Space(1) & rs("ErpNo"), mAlignLeftCenter, False, 2, 1)
      Call spSetText(frmPrint.sprPrint, 9, sprow, "TS Ref. No   :" & Space(1) & rs("TsNo"), mAlignLeftCenter, False, 2, 1)
      sprow = sprow + 1
GridCount = GridCount + 1
    End If
      
    'Company Name
    frmPrint.sprPrint.Row = sprow
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.FontSize = 12
frmPrint.sprPrint.FontUnderline = True
    Call spSetText(frmPrint.sprPrint, 1, sprow, gCompName, mAlignCenterCenter, True, frmPrint.sprPrint.MaxCols, 1)
    sprow = sprow + 2
GridCount = GridCount + 2
        
    'Report Heading
    frmPrint.sprPrint.Row = sprow
frmPrint.sprPrint.Col = 1

    frmPrint.sprPrint.FontUnderline = True
frmPrint.sprPrint.FontSize = 11
    Call spSetText(frmPrint.sprPrint, 1, sprow, "LC STATEMENT [ " & Format(dtpFromDate, "dd/MM/yyyy") & " To " & Format(dtpTodate, "dd/MM/yyyy") & " ]", mAlignCenterCenter, True, frmPrint.sprPrint.MaxCols, 1)
    sprow = sprow + 3
GridCount = GridCount + 3
        
    'Setting Columns width
    Call spSetColWidth(frmPrint.sprPrint, 1, 6, 2, 8, 3, 9, 4, 10, 5, 10, 6, 8, 7, 9, 8, 10, 9, 10, 10, 10, 11, 10)
    
    If iPageNo = 1 Then
      '1st Row
      Call spSetTextColArray(frmPrint.sprPrint, 1, sprow, 1, mAlignLeftCenter, True, "LC No", " : " & Trim$(txtLCNo.Text))

      Call spSetTextColArray(frmPrint.sprPrint, 8, sprow, 1, mAlignLeftCenter, True, "Customer", " : " & Trim$(txtCustomerName.Text))

      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess
      
      '2nd Row
      Call spSetTextColArray(frmPrint.sprPrint, 1, sprow, 1, mAlignLeftCenter, True, "LC Date", " : " & Trim$(txtLCDate.Text))

      Call spSetTextColArray(frmPrint.sprPrint, 8, sprow, 1, mAlignLeftCenter, True, "LC Amount", " : " & Format(Val(TaglblLCAmount.Caption), "0.00"))

      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess
      
      '3rd Row
      Call spSetTextColArray(frmPrint.sprPrint, 1, sprow, 1, mAlignLeftCenter, True, "O/B", " : " & Format(Val(TaglblOPBalance.Caption), "0.00"))

      Call spSetTextColArray(frmPrint.sprPrint, 8, sprow, 1, mAlignLeftCenter, True, "C/B", " : " & Format(Val(TaglblClosingBal.Caption), "0.00"))
      k = sprow
      
      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess

      sprow = sprow + 1
GridCount = GridCount + 1
RowProcess
    End If
    
    Call spSetCellBorder(frmPrint.sprPrint, 1, sprow, frmPrint.sprPrint.MaxCols, sprow, mBorderAllSides)
    Call Spread_BkColour(frmPrint.sprPrint, &H8000000F, 1, sprow, frmPrint.sprPrint.MaxCols, sprow)
    
    'Header
    Call spSetTextColArray(frmPrint.sprPrint, 1, sprow, 1, mAlignCenterCenter, True, "S.No", "INV No", "INV Date", "INV Qty", "INV Amt", "RR No", "RR Date", "Recd Qty", "Acc Qty", "Rej Qty", "Adjusted Amt")
    sprow = sprow + 1
GridCount = GridCount + 1
RowProcess
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Public Sub sprint()

  Try
  
  frmPrint.cmdPPrint.Enabled = True
  frmPrint.cmdExcel.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub sView()
    
  Try
  
  frmPrint.cmdPPrint.Enabled = False
  frmPrint.cmdExcel.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub CmdClear_Click()
' On Error Resume Next
  
  txtLCNo.Text = ""
  txtLCDate.Text = ""
  txtCustomerName.Text = ""
  TaglblCustomerCode.Caption = ""
  TaglblLCAmount.Caption = ""
  
  fraInvDate.Enabled = False
  fraInvType.Enabled = False
  
  cmbType.ListIndex = 0
  CmbInvType.ListIndex = 0
  
  Call StepClear(Me)
  Call StepEnable(Me)
  
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub CmdStep_Click(ByVal Index As Integer)
  
  Try
  Me.MousePointer = vbHourglass
 
  If contrl_valid(txtLCNo, "Select the  LC No") = False Then
 GoTo ExitHere
End If
  
  fraInvDate.Enabled = True
  fraInvType.Enabled = True
  
  Call StepProceedNext(Me, Index)
  cmdview.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdView_Click()

  Try
  Me.MousePointer = vbHourglass
      
  If dtpTodate < dtpFromDate Then
    sShowMessage "To Date Should be greater than From Date"
    dtpTodate.SetFocus
    GoTo ExitHere
  End If
          
  If sPrintCurrent_Print = False Then
 GoTo ExitHere
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_Load()
  
  Try
  Me.MousePointer = vbHourglass

  Call CurrDate(Me)
  
  dtpFromDate.Value = Format(ServerDate, "dd/mm/yyyy")
  dtpTodate.Value = Format(dtpFromDate.Value, "dd/mm/yyyy")
  
  cmbType.ListIndex = 0
  CmbInvType.ListIndex = 0
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
    MDISales.Toolbar1.Visible = False
  End If
  
End Sub

Private Sub txtcustomername_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = " Execute SAL_Rpt_LCReport_Hlp_Customer"
  Sql = Sql & " @BranchID='" & gBranchID & "',"
  Sql = Sql & " @CompID='" & gCompID & "'"

  If FetchValue_New(False, db.ConnectionString, Sql, "01", "List of Customer's", "Customer") = True Then
    
    If Val(TaglblCustomerCode.Caption) <> Val(getvalue(1)) Then
      txtLCNo.Text = ""
txtLCDate.Text = ""
    End If
    
    TaglblCustomerCode.Caption = getvalue(1)
    txtCustomerName.Text = Trim$(getvalue(2))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtCustomerName_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  
  If KeyCode = vbKeyF2 Then
    Call txtcustomername_DblClick
  End If
  
End Sub

Private Sub txtLCNo_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = " Execute SAL_Rpt_LCReport_Hlp_LCNo"
  Sql = Sql & " @CustomerCode=" & Val(TaglblCustomerCode.Caption) & ","
  Sql = Sql & " @BranchID='" & gBranchID & "',"
  Sql = Sql & " @CompID='" & gCompID & "'"

  If FetchValue_New(False, db.ConnectionString, Sql, "1101", "List of LC No's", "LC No", "LC Date", "Customer") = True Then
    txtLCNo.Text = getvalue(1)
    txtLCDate.Text = getvalue(2)
    TaglblCustomerCode.Caption = getvalue(3)
    txtCustomerName.Text = Trim$(getvalue(4))
    TaglblLCAmount.Caption = Val(getvalue(5))
    
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtLCNo_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  
  If KeyCode = vbKeyF2 Then
    Call txtLCNo_DblClick
  End If
  
End Sub
