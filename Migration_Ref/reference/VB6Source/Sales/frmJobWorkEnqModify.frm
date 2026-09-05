VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJobWorkEnqModify 
   Caption         =   "JobWork Enquiry Modify"
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
   Begin VB.Frame fraSent 
      Caption         =   "Sent ItemCode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2160
      TabIndex        =   0
      Top             =   4980
      Visible         =   0   'False
      Width           =   7515
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Height          =   390
         Left            =   2790
         TabIndex        =   1
         Top             =   765
         Width           =   2175
      End
Begin AceSpread sprSent
         Height          =   525
         Left            =   75
         TabIndex        =   2
         Top             =   210
         Width           =   7380
         _Version        =   458752
         _ExtentX        =   13017
         _ExtentY        =   926
         _StockProps     =   64
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   1
         ScrollBars      =   0
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6375
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   3855
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5085
      TabIndex        =   8
      ToolTipText     =   "Click to Clear the Data"
      Top             =   3855
      Width           =   1300
   End
   Begin VB.Frame FraEnq 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8985
      Begin VB.TextBox txtEnqType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   28
         ToolTipText     =   "Jobwork Enquiry No"
         Top             =   330
         Width           =   960
      End
      Begin VB.TextBox txtRefNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3120
         MaxLength       =   15
         TabIndex        =   16
         ToolTipText     =   "Enter the Ref. No"
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox txtEnqNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   15
         ToolTipText     =   "Jobwork Enquiry No"
         Top             =   330
         Width           =   795
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         MaxLength       =   50
         TabIndex        =   14
         ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help."
         Top             =   915
         Width           =   3690
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3780
         MaxLength       =   50
         TabIndex        =   13
         ToolTipText     =   "Enter the Remarks"
         Top             =   915
         Width           =   4575
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   8385
         TabIndex        =   12
         ToolTipText     =   "Click to OK"
         Top             =   915
         Width           =   495
      End
      Begin VB.ComboBox cmbCustomerType 
         Height          =   330
         ItemData        =   "frmJobWorkEnqModify.frx":030A
         Left            =   5805
         List            =   "frmJobWorkEnqModify.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Select the Customer Type"
         Top             =   315
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker dtpEnqDate 
         Height          =   315
         Left            =   1875
         TabIndex        =   17
         ToolTipText     =   "Jobwork Enquiry Date"
         Top             =   330
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   65208321
         CurrentDate     =   37258
      End
      Begin MSComCtl2.DTPicker dtpRefDate 
         Height          =   315
         Left            =   4560
         TabIndex        =   18
         ToolTipText     =   "Enter the Ref. Date"
         Top             =   330
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   65208321
         CurrentDate     =   37258
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "Customer"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   26
         Top             =   735
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000000&
         Caption         =   "Enquiry Type"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   75
         TabIndex        =   25
         Top             =   135
         Width           =   945
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Remarks"
         Height          =   195
         Left            =   3765
         TabIndex        =   24
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Number"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1290
         TabIndex        =   23
         Top             =   135
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Enquiry Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1875
         TabIndex        =   22
         Top             =   150
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ref. No"
         Height          =   210
         Left            =   3120
         TabIndex        =   21
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ref. Date"
         Height          =   210
         Left            =   4560
         TabIndex        =   20
         Top             =   150
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Customer Type"
         Height          =   210
         Left            =   5805
         TabIndex        =   19
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   400
      Left            =   3795
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Click to Save the Data"
      Top             =   3855
      Width           =   1300
   End
   Begin VB.Frame fraRecd 
      Caption         =   "Reeived ItemCode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   2025
      TabIndex        =   3
      Top             =   4500
      Visible         =   0   'False
      Width           =   7755
      Begin VB.CommandButton cmdRCancel 
         Caption         =   "Cancel"
         Height          =   390
         Left            =   3660
         TabIndex        =   6
         Top             =   1710
         Width           =   2175
      End
      Begin VB.CommandButton CmdROk 
         Caption         =   "OK"
         Height          =   390
         Left            =   1500
         TabIndex        =   5
         Top             =   1710
         Width           =   2175
      End
Begin AceSpread sprRecd
         Height          =   1470
         Left            =   75
         TabIndex        =   4
         Top             =   210
         Width           =   7605
         _Version        =   458752
         _ExtentX        =   13414
         _ExtentY        =   2593
         _StockProps     =   64
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   1
         ScrollBars      =   2
      End
   End
Begin AceSpread sprItem
      Height          =   2400
      Left            =   0
      TabIndex        =   27
      ToolTipText     =   "Enter the Jobwork Details"
      Top             =   1335
      Width           =   11460
      _Version        =   458752
      _ExtentX        =   20214
      _ExtentY        =   4233
      _StockProps     =   64
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   18
      MaxRows         =   1
      ProcessTab      =   -1  'True
      SelectBlockOptions=   0
   End
End
Attribute VB_Name = "frmJobWorkEnqModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim CurrDate As Date

'Item Spread Column Numbers
Private Enum eItem
  mFirstCol = 1
  mItemType = 2
  mrecdItem = 3
  mRecdDesc = 4
  mRecdUOM = 5
  msentItem = 6
  mSentDesc = 7
  mSentUOM = 8
  mSentQty = 9
  mVolumeper = 10
  mRatePer = 11
  mRate = 12
  mAssRate = 13
  mTarDate = 14
  mGoodsValue = 15
  mAssGoodsValue = 16
  mRelation = 17
  mItemGroupKey = 18
End Enum

Private Function ItemGroupKey(ByVal sChildCode As String,ByVal  sParentCode As String) As Long
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try

  Sql = "select * from dbo.fn_ItemGroupKey('" & Trim$(sParentCode) & "','" & gBranchID & "','" & gCompID & "')"
  Sql = Sql & "where FromCode='" & Trim$(sChildCode) & "'"
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    
  If rs.EOF Then
    Sql = "select isnull(max(ItemGroupKey),0)as ItemGroupKey from ItemGroupKey_Dtl"
    If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
    ItemGroupKey = rs("ItemGroupKey") + 1 & 1 ' New
  Else
    ItemGroupKey = rs("ItemGroupKey") & 0 'OLD
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function Group_Child(ByVal fRow As Long,ByVal  ParentCode As String) As String
  Dim i As Integer
  Dim retVal As Variant
  Dim Row As Long
  Dim SRow As Long
  Dim ChildCode As String
  
  Try
  Me.MousePointer = vbHourglass
  
    SRow = sprItem.SearchCol(eItem.mFirstCol, 0, sprItem.MaxRows, Val(fRow), SearchFlagsValue)
    ChildCode = ""
    Do While SRow <> -1
      sprItem.GetText eItem.mrecdItem, SRow, retVal
      ChildCode = ChildCode + retVal
      SRow = sprItem.SearchCol(eItem.mFirstCol, SRow, sprItem.MaxRows, Trim$(fRow), SearchFlagsValue)
    Loop
  
  Group_Child = ChildCode
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Function

Private Function HelpDirectItem(ByVal fCol As Long,ByVal  fRow As Long) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim sFirstRow As Long
  
  Try
  HelpDirectItem = False
  
  If fCol = eItem.mrecdItem Then
    'Show 1 to 1 Item for the Vendor
    Sql = " Select a.ReceivedItemCode, c.ItemDescription as RecdDesc, c.PurchaseUOMDesc as RecdUOM,"
    Sql = Sql & " a.SentItemCode, b.ItemDescription as SentDesc, b.PurchaseUOMDesc as SentUOM, b.PurchaseUOMDecimal as SentNoOfDec,"
    Sql = Sql & " a.RatePer, a.Rate, a.ItemGroupKey from SAL_JOBWORK_ITEMMASTER a Left Outer Join Material_Master b"
    Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Left Outer Join Material_Master c"
    Sql = Sql & " On a.SentItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
    Sql = Sql & " where a.CustomerCode=" & Val(txtCustomer.Tag) & " and a.RelationType='1-1' and a.isActive='1'"
    Sql = Sql & " and a.ValidFrom <= '" & Format(dtpEnqDate, "yyyy-MM-dd") & "' and a.ValidTo >= '" & Format(dtpEnqDate, "yyyy-MM-dd") & "'"
    Sql = Sql & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and b.Status = 1 and c.Status = 1 "
    If FetchValue_New(False, db.ConnectionString, Sql, "1111", "List of Item", "Recd Item", "Description", "Sent Item", "Description") = True Then
        If fRow <= sprItem.DataRowCnt Then
          rCnt = fRow
        Else
          sprItem.MaxRows = sprItem.DataRowCnt + 1
          rCnt = sprItem.DataRowCnt + 1
        End If
        
        'Key Values First Set in Spread
        sprItem.SetText eItem.mItemGroupKey, rCnt, getvalue(10)
        sprItem.SetText eItem.mrecdItem, rCnt, getvalue(1)
        sprItem.SetText eItem.msentItem, rCnt, getvalue(4)
        
        If Spread_Clear_Duplicate(sprItem, rCnt, eItem.mItemGroupKey, eItem.mrecdItem, eItem.msentItem) = False Then
          sprItem.SetText eItem.mFirstCol, rCnt, rCnt
          
          sprItem.SetText eItem.mRecdDesc, rCnt, getvalue(2)
          sprItem.SetText eItem.mRecdUOM, rCnt, getvalue(3)
          
          sprItem.SetText eItem.mSentDesc, rCnt, getvalue(5)
          sprItem.SetText eItem.mSentUOM, rCnt, getvalue(6)
          Call Spread_Set_Decimal(sprItem, Val(getvalue(7)), eItem.mSentQty, rCnt)
          
          sprItem.SetText eItem.mRatePer, rCnt, getvalue(8)
          sprItem.SetText eItem.mRate, rCnt, getvalue(9)
          sprItem.SetText eItem.mAssRate, rCnt, getvalue(9)
          
          sprItem.SetText eItem.mRelation, rCnt, "1-1"
        End If
        sprItem.MaxRows = sprItem.DataRowCnt + 1    Else
      GoTo ExitHere
    End If
  Else
    'Show N to 1 Item for the Vendor
    
    Sql = "Select a.SentItemCode, b.ItemDescription as SentDesc, a.ReceivedItemCode, c.ItemDescription as RecdDesc, a.ItemGroupKey"
    Sql = Sql & " from SAL_JOBWORK_ITEMMASTER a Left Outer Join Material_Master b"
    Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Left Outer Join Material_Master c"
    Sql = Sql & " On a.SentItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
    Sql = Sql & " where a.CustomerCode=" & Val(txtCustomer.Tag) & " and a.RelationType='N-1' and a.isActive='1'"
    Sql = Sql & " and (a.ValidFrom <= '" & Format(dtpEnqDate, "yyyy-MM-dd") & "' and a.ValidTo >= '" & Format(dtpEnqDate, "yyyy-MM-dd") & "')"
    Sql = Sql & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and b.Status = 1 and c.Status = 1"
    Sql = Sql & " Order by a.ReceivedItemCode, a.SentItemCode, a.ItemGroupKey"
    If FetchValue_New(False, db.ConnectionString, Sql, "11110", "List of Item", "Sent Item", "Description", "Recd Item", "Description") = True Then
    
      Sql = "Select a.ReceivedItemCode, c.ItemDescription as RecdDesc, c.PurchaseUOMDesc as RecdUOM,"
      Sql = Sql & " a.SentItemCode, b.ItemDescription as SentDesc, b.PurchaseUOMDesc as SentUOM, b.PurchaseUOMDecimal as SentNoOfDec,"
      Sql = Sql & " a.RatePer, a.Rate, a.ItemGroupKey"
      Sql = Sql & " from SAL_JOBWORK_ITEMMASTER a Left Outer Join Material_Master b"
      Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
      Sql = Sql & " Left Outer Join Material_Master c"
      Sql = Sql & " On a.SentItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
      Sql = Sql & " where a.CustomerCode=" & Val(txtCustomer.Tag) & " and a.RelationType='N-1' and a.isActive='1'"
      Sql = Sql & " and a.ItemGroupKey=" & getvalue(5) & " and b.Status = 1 and c.Status = 1"
      Sql = Sql & " and a.ValidFrom <= '" & Format(dtpEnqDate, "yyyy-MM-dd") & "' and a.ValidTo >= '" & Format(dtpEnqDate, "yyyy-MM-dd") & "'"
      Sql = Sql & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
      If rs.EOF Then
 GoTo ExitHere
End If
        rs.MoveLast
rs.MoveFirst
        If fRow <= sprItem.DataRowCnt Then
          sFirstRow = fRow
          sprItem.MaxRows = rs.RecordCount + 1
        Else
          sFirstRow = sprItem.DataRowCnt + 1
          sprItem.MaxRows = sprItem.DataRowCnt + rs.RecordCount + 1
        End If
        For rCnt = sFirstRow To sprItem.MaxRows - 1
          'Key Values First Set
          sprItem.SetText eItem.mItemGroupKey, rCnt, rs("ItemGroupKey")
          sprItem.SetText eItem.mrecdItem, rCnt, rs("ReceivedItemCode")
          If sFirstRow = rCnt Then
            sprItem.SetText eItem.msentItem, rCnt, rs("SentItemCode")
          End If
          
          If Spread_Clear_Duplicate(sprItem, rCnt, eItem.mItemGroupKey, eItem.msentItem, eItem.mrecdItem) = True Then
            Exit For
          End If
        
          sprItem.SetText eItem.mFirstCol, rCnt, sFirstRow
          
          sprItem.SetText eItem.mRecdDesc, rCnt, rs("SentDesc")
          sprItem.SetText eItem.mRecdUOM, rCnt, rs("SentUOM")
          
          If sFirstRow = rCnt Then
            sprItem.SetText eItem.mSentDesc, rCnt, rs("RecdDesc")
            sprItem.SetText eItem.mSentUOM, rCnt, rs("RecdUOM")
            sprItem.SetText eItem.mRatePer, rCnt, rs("RatePer")
            sprItem.SetText eItem.mRate, rCnt, rs("Rate")
            sprItem.SetText eItem.mAssRate, rCnt, rs("Rate")
            Call Spread_Lock(sprItem, False, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
          Else
            Call Spread_Lock(sprItem, True, eItem.mSentQty, rCnt, eItem.mSentQty, rCnt)
          End If
          
          sprItem.SetText eItem.mRelation, rCnt, "N-1"
          rs.MoveNext
        Next rCnt
    Else
      GoTo ExitHere
    End If
  End If
  
  HelpDirectItem = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function Chkbefsave() As Boolean
  Dim i As Long
  Dim SRow As Long
  Dim retVal As Variant
  Dim Sql As String
  Dim Group As Long
  Dim rs As AceADODB.Recordset

  Try

  If sprItem.DataRowCnt = 0 Then
    sShowMessage "Enter atleast One Item"
    Spread_Set_Focus sprItem, 1, 2
    Chkbefsave = False
    GoTo ExitHere
  End If
  
    For i = 1 To sprItem.DataRowCnt
      sprItem.Row = i
sprItem.Col = eItem.mrecdItem
      If Trim$(sprItem.Text) = "" Then
        sShowMessage "Select the Item Code"
        Spread_Set_Focus sprItem, i, eItem.mrecdItem
        Chkbefsave = False
        GoTo ExitHere
      End If
    
      sprItem.Row = i
sprItem.Col = eItem.mFirstCol
SRow = Val(sprItem.Text)

      sprItem.Row = SRow
sprItem.Col = eItem.mSentQty
      If Val(sprItem.Text) = 0 Then
        sShowMessage "Select the Sent Qty"
        Spread_Set_Focus sprItem, SRow, eItem.mSentQty
        Chkbefsave = False
        GoTo ExitHere
      End If
      
      sprItem.Row = i
sprItem.Col = eItem.mFirstCol
SRow = Val(sprItem.Text)

      sprItem.Row = SRow
sprItem.Col = eItem.mRate
      If Val(sprItem.Text) = 0 Then
        sShowMessage "Select the Rate"
        Spread_Set_Focus sprItem, SRow, eItem.mRate
        Chkbefsave = False
        GoTo ExitHere
      End If
      
      sprItem.GetText eItem.mRecdUOM, i, retVal
      If Trim$(retVal) = "" Then
        sShowMessage "Select The RecdUOM"
        Spread_Set_Focus sprItem, i, eItem.mRecdUOM
        GoTo ExitHere
      End If
      
      sprItem.Row = i
sprItem.Col = eItem.mFirstCol
SRow = Val(sprItem.Text)

      sprItem.Row = SRow
sprItem.Col = eItem.mSentUOM
      If Trim$(sprItem.Text) = "" Then
        sShowMessage "Select The SentUOM"
        Spread_Set_Focus sprItem, i, eItem.mSentUOM
        GoTo ExitHere
      End If
      
      sprItem.GetText eItem.mRecdDesc, i, retVal
      If Trim$(retVal) = "" Then
        sShowMessage "Enter the Recd Decsription"
        Spread_Set_Focus sprItem, i, eItem.mRecdDesc
        GoTo ExitHere
      End If
      
      sprItem.Row = i
sprItem.Col = eItem.mFirstCol
SRow = Val(sprItem.Text)

      sprItem.Row = SRow
sprItem.Col = eItem.mSentDesc
      If Trim$(sprItem.Text) = "" Then
        sShowMessage "Enter the Sent Decsription"
        Spread_Set_Focus sprItem, i, eItem.mSentDesc
        GoTo ExitHere
      End If
      
      sprItem.Row = i
sprItem.Col = eItem.mFirstCol
SRow = Val(sprItem.Text)

      sprItem.Row = SRow
sprItem.Col = eItem.msentItem
      If Trim$(sprItem.Text) = "" Then
        sShowMessage "Enter the Sent ItemCode"
        Spread_Set_Focus sprItem, i, eItem.msentItem
        GoTo ExitHere
      End If
    Next i
  Chkbefsave = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmbCustomerType_Click()
  Try
  If cmbCustomerType = "Existing Customer" Then
    txtCustomer.Locked = True
    Label2.ForeColor = &HFF0000
    txtCustomer = ""
txtCustomer.Tag = ""
  ElseIf cmbCustomerType = "New Customer" Then
    txtCustomer.Locked = False
    Label2.ForeColor = &H80000012
    txtCustomer = ""
txtCustomer.Tag = ""
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdClear_Click()
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  FraEnq.Enabled = True
  sprItem.Enabled = False
  cmdSave.Enabled = False
  sprItem.MaxRows = 1
  cmbCustomerType.ListIndex = 0
  txtEnqNo.SetFocus
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  dtpEnqDate = CurrDate
  dtpRefDate = CurrDate
  sprItem.SetText eItem.mItemType, 1, "Existing Item"
  sprItem.SetText eItem.mFirstCol, 1, "1"
  fraRecd.Visible = False
  fraSent.Visible = False
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
  Dim i As Long
  Dim rs As AceADODB.Recordset
  
  Try
  
  If Trim$(txtEnqType) = "" Then
    sShowMessage "Select the Enquiry type"
    txtEnqNo.SetFocus
    GoTo ExitHere
  End If

  If contrl_valid(txtRefNo, "Enter the Customer Ref. No") = False Then
 GoTo ExitHere
End If
  If contrl_valid(txtCustomer, "Select the Customer") = False Then
 GoTo ExitHere
End If
  
  Sql = "Select *,(select NoofDecimals from INV_UOM_MASTER where SDesc=a.SentUOM and Branchid =a.Branchid"
  Sql = Sql & " and compid=a.Compid)NoofDec from SAL_VW_JobWorkEnquiry_Dtl a where EnquiryType='" & Trim$(txtEnqType) & "' and EnquiryNo=" & Val(txtEnqNo) & " and EnquiryDate='" & Format(dtpEnqDate, "yyyy-MM-dd") & "' "
  Sql = Sql & " and a.QuotationType='' and a.QuotationDate is null and BRanchid='" & Trim$(gBranchID) & "' and Compid='" & Trim$(gCompID) & "' "
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If Not rs.EOF Then
      For i = 1 To rs.RecordCount
        sprItem.MaxRows = rs.RecordCount + 1
        sprItem.SetText eItem.mItemType, i, IIf(Trim$(rs("ItemType")) = "N", "New Item", "Existing Item")
        Call sprItem_ComboSelChange(eItem.mItemType, i)
        Call Spread_Set_Decimal(sprItem, Val(rs("NoofDec")), eItem.mSentQty, i)
        Call Spread_Lock(sprItem, False, eItem.mSentQty, i)
        Call Spread_Lock(sprItem, False, eItem.mVolumeper, i)
        sprItem.SetText eItem.mrecdItem, i, rs("ReceivedItemCode")
        sprItem.SetText eItem.mRecdDesc, i, rs("RecdDescription")
        sprItem.SetText eItem.mRecdUOM, i, rs("RecdUom")
        sprItem.SetText eItem.msentItem, i, rs("SentItemCode")
        sprItem.SetText eItem.mSentDesc, i, rs("SentDescription")
        sprItem.SetText eItem.mSentUOM, i, rs("SentUOM")
        sprItem.SetText eItem.mSentQty, i, Val(rs("SendQty"))
        sprItem.SetText eItem.mVolumeper, i, Trim$(rs("VolumePer"))
        sprItem.SetText eItem.mRatePer, i, Val(rs("RatePer"))
        sprItem.SetText eItem.mRate, i, rs("Rate")
        sprItem.SetText eItem.mAssRate, i, rs("AssRate")
        Call sprItem_EditMode(eItem.mRate, i, 0, True)
        sprItem.SetText eItem.mTarDate, i, rs("TargetDate")
        sprItem.SetText eItem.mRelation, i, Trim$(rs("RelationType"))
        sprItem.SetText eItem.mItemGroupKey, i, Trim$(rs("ItemGroupKey"))
        rs.MoveNext
      Next i
  End If
  
  FraEnq.Enabled = False
  sprItem.Enabled = True
  Call Spread_Lock(sprItem, True, eItem.mrecdItem, 1, eItem.mRecdDesc)
  Call Spread_Lock(sprItem, False, eItem.mItemType, 1, eItem.mItemType)
  sprItem.SetFocus
  Spread_Set_Focus sprItem, eItem.mItemType, 1
  cmdSave.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

    Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdRCancel_Click()
' On Error Resume Next
  fraRecd.Visible = False
  fraSent.Visible = False
End Sub

Private Sub CmdROk_Click()
  Dim tmp As Variant
  Dim Row As Variant
  Dim firstCol As Integer
  Dim i As Long
  
  Try
  Me.MousePointer = vbHourglass

    For i = 1 To sprRecd.DataRowCnt
      sprRecd.GetText 1, i, tmp
      If Trim$(tmp) = "" Then
        sShowMessage "Enter the ItemCode"
        Spread_Set_Focus sprRecd, i, 1
        GoTo ExitHere
      End If
      sprRecd.GetText 2, i, tmp
      If Trim$(tmp) = "" Then
        sShowMessage "Enter the ItemDescription"
        Spread_Set_Focus sprRecd, i, 2
        GoTo ExitHere
      End If
    Next i
    
    fraRecd.Visible = False
    fraSent.Visible = True
    sprSent.ClearRange 1, 1, -2, -2, True
    
    sprRecd.GetText 3, 1, Row
    firstCol = Val(Row)
    sprItem.MaxRows = sprItem.DataRowCnt + sprRecd.DataRowCnt
    For i = 1 To sprRecd.DataRowCnt
      sprRecd.GetText 1, i, tmp
      sprItem.SetText eItem.mrecdItem, Row, tmp
      sprRecd.GetText 2, i, tmp
      sprItem.SetText eItem.mRecdDesc, Row, tmp
      sprItem.SetText eItem.mFirstCol, Row, firstCol
'      sprItem.SetText eItem.mItemType, Row, "New Item(N-1)"
      sprItem.SetText eItem.mRelation, Row, "N-1"
      Call Spread_Lock(sprItem, True, eItem.mItemType, CLng(Row), eItem.mGoodsValue, CLng(Row))
'      Call Spread_Lock(sprItem, False, eItem.mSentUOM, CLng(Row), eItem.mRecdUOM, CLng(Row))
      Row = Row + 1
    Next i
    sprSent.SetText 3, 1, firstCol
    Spread_Set_Focus sprSent, 1, 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
  Dim retVal As Variant
  Dim cmd As AceADODB.Command
  Dim i As Long
  Dim sFirstRow As Long
  Dim CurDate As Date
  Dim sRelation As String
  Dim MaxItemGroup As Integer
  Dim sChildCode As String
  Dim SRow As Long
  Dim ChkSave As Long
  Dim Sql As String

  Try
  Me.MousePointer = vbHourglass

  Set cmd = New AceADODB.Command
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  If Chkbefsave = False Then GoTo ExitHere

  CurDate = ServerDate
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  commClear cmd
  cmd.CommandText = "SAL_INS_JOBWORKENQUIRY_HDR"
  cmd.Parameters.Append cmd.CreateParameter("EnquiryType", adVarChar, adParamInput, 3, Trim$(txtEnqType))
  cmd.Parameters.Append cmd.CreateParameter("EnquiryNo", adInteger, adParamInput, , Val(Trim$(txtEnqNo)))
  cmd.Parameters.Append cmd.CreateParameter("EnquiryDate", adDBTimeStamp, adParamInput, , Format(dtpEnqDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adInteger, adParamInput, , Val(txtEnqNo.Tag))
  cmd.Parameters.Append cmd.CreateParameter("RevisionDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("ReferenceNo", adVarChar, adParamInput, 30, Trim$(txtRefNo))
  cmd.Parameters.Append cmd.CreateParameter("ReferenceDate", adDBTimeStamp, adParamInput, , Format(dtpRefDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("CustomerType", adChar, adParamInput, 1, Left$(Trim$(cmbCustomerType), 1))
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(txtCustomer.Tag))
  cmd.Parameters.Append cmd.CreateParameter("CustomerName", adVarChar, adParamInput, 50, Trim$(txtCustomer.Text))
  cmd.Parameters.Append cmd.CreateParameter("Remarks", adVarChar, adParamInput, 150, Trim$(txtremarks))
  Call AddCommonCmdParameters(cmd)
  cmd.Execute
   
    cmd.CommandText = "SAL_INS_JOBWORKENQUIRY_DTL"
    For i = 1 To sprItem.DataRowCnt
      sprItem.GetText eItem.mFirstCol, i, retVal
      sFirstRow = Val(retVal)
      If i = sFirstRow Then
        SRow = sprItem.SearchCol(eItem.mFirstCol, 0, sprItem.MaxRows, Trim$(sFirstRow), SearchFlagsValue)
        sChildCode = ""
        Do While SRow <> -1
          sprItem.GetText eItem.mrecdItem, SRow, retVal
          sChildCode = sChildCode + Trim$(retVal)
          SRow = sprItem.SearchCol(eItem.mFirstCol, SRow, sprItem.MaxRows, Trim$(sFirstRow), SearchFlagsValue)
        Loop
        sprItem.GetText eItem.msentItem, sFirstRow, retVal
        MaxItemGroup = Val(ItemGroupKey(sChildCode, Trim$(retVal)))
        ChkSave = Right(MaxItemGroup, 1)
        MaxItemGroup = Left(MaxItemGroup, Len(Trim$(MaxItemGroup)) - 1)
      End If
      
      If ChkSave = 1 Then 'New
        sprItem.GetText eItem.mrecdItem, i, retVal
        Sql = "Insert into ItemGroupKey_Dtl select " & Val(MaxItemGroup) & ",'" & Trim$(retVal) & "',"
        sprItem.GetText eItem.msentItem, sFirstRow, retVal
        Sql = Sql & " '" & Trim$(retVal) & "',"
        sprItem.GetText eItem.mRelation, i, retVal
        Sql = Sql & " '" & Trim$(retVal) & "',1,'" & gUserID & "','" & gBranchID & "','" & gCompID & "', '" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      End If
      
      commClear cmd
      cmd.Parameters.Append cmd.CreateParameter("EnquiryType", adVarChar, adParamInput, 3, Trim$(txtEnqType))
      cmd.Parameters.Append cmd.CreateParameter("EnquiryNo", adInteger, adParamInput, , Val(Trim$(txtEnqNo)))
      cmd.Parameters.Append cmd.CreateParameter("EnquiryDate", adDBTimeStamp, adParamInput, , Format(dtpEnqDate, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("RevisionNo", adInteger, adParamInput, , Val(txtEnqNo.Tag))
      cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(MaxItemGroup))
      sprItem.GetText eItem.mRelation, i, retVal
      sRelation = Trim$(retVal)
      If sRelation = "1-1" Or sRelation = "N-1" Then
        sprItem.GetText eItem.mrecdItem, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(retVal))
        sprItem.GetText eItem.mRecdDesc, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("RecdDescription", adVarChar, adParamInput, 50, Trim$(retVal))
        sprItem.GetText eItem.mRecdUOM, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("RecdUOM", adVarChar, adParamInput, 3, Trim$(retVal))

        sprItem.GetText eItem.mItemType, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("ItemType", adChar, adParamInput, 1, Left$(Trim$(retVal), 1))
        sprItem.GetText eItem.msentItem, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(retVal))
        sprItem.GetText eItem.mSentDesc, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("SentDescription", adVarChar, adParamInput, 50, Trim$(retVal))
        sprItem.GetText eItem.mSentUOM, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("SentUOM", adVarChar, adParamInput, 3, Trim$(retVal))

      Else
        sprItem.GetText eItem.mrecdItem, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("ReceivedItemCode", adVarChar, adParamInput, 20, Trim$(retVal))
        sprItem.GetText eItem.mRecdDesc, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("RecdDescription", adVarChar, adParamInput, 50, Trim$(retVal))
        sprItem.GetText eItem.mRecdUOM, i, retVal
        cmd.Parameters.Append cmd.CreateParameter("RecdUOM", adVarChar, adParamInput, 3, Trim$(retVal))
        
        sprItem.GetText eItem.mItemType, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("ItemType", adChar, adParamInput, 1, Left$(Trim$(retVal), 1))
        sprItem.GetText eItem.msentItem, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("SentItemCode", adVarChar, adParamInput, 20, Trim$(retVal))
        sprItem.GetText eItem.mSentDesc, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("SentDescription", adVarChar, adParamInput, 50, Trim$(retVal))
        sprItem.GetText eItem.mSentUOM, sFirstRow, retVal
        cmd.Parameters.Append cmd.CreateParameter("SentUOM", adVarChar, adParamInput, 3, Trim$(retVal))

      End If
      
      sprItem.GetText eItem.mSentQty, sFirstRow, retVal
      cmd.Parameters.Append cmd.CreateParameter("SentQty", adDouble, adParamInput, , Val(retVal))
      sprItem.GetText eItem.mVolumeper, sFirstRow, retVal
      cmd.Parameters.Append cmd.CreateParameter("VolumePer", adVarChar, adParamInput, 15, Trim$(retVal))
      sprItem.GetText eItem.mRatePer, sFirstRow, retVal
      cmd.Parameters.Append cmd.CreateParameter("RatePer", adDouble, adParamInput, , Val(retVal))
      sprItem.GetText eItem.mRate, sFirstRow, retVal
      cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Val(retVal))
      sprItem.GetText eItem.mAssRate, sFirstRow, retVal
      cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Val(retVal))
      sprItem.GetText eItem.mTarDate, sFirstRow, retVal
      cmd.Parameters.Append cmd.CreateParameter("TargetDate", adDBTimeStamp, adParamInput, , Format(retVal, "yyyy-MM-dd"))
      sprItem.GetText eItem.mRelation, i, retVal
      cmd.Parameters.Append cmd.CreateParameter("RelationType", adVarChar, adParamInput, 3, Trim$(retVal))
      cmd.Parameters.Append cmd.CreateParameter("RateEffDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("isActive", adBoolean, adParamInput, , 1)
      cmd.Parameters.Append cmd.CreateParameter("Status", adChar, adParamInput, 1, "N")
      cmd.Parameters.Append cmd.CreateParameter("StatusDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("EffStatus", adBoolean, adParamInput, , 1)
      cmd.Parameters.Append cmd.CreateParameter("EffDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("CreatedDate", adDBTimeStamp, adParamInput, , Format(CurDate, "yyyy-MM-dd"))
      Call AddCommonCmdParameters(cmd)
      cmd.Execute
    Next i
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Jobwork Enquiry Number :" & Trim$(txtEnqType) & Trim$(txtEnqNo) & "  Is Updated"
  cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  Set cmd = Nothing
End Try
End Sub

Private Sub Command1_Click()
  Dim tmp As Variant
  Dim Row As Long
  
  Try
  Me.MousePointer = vbHourglass

    sprSent.GetText 1, 1, tmp
    If Trim$(tmp) = "" Then
      sShowMessage "Enter the ItemCode"
      Spread_Set_Focus sprSent, 1, 1
      GoTo ExitHere
    End If
    sprSent.GetText 2, 1, tmp
    If Trim$(tmp) = "" Then
      sShowMessage "Enter the Description"
      Spread_Set_Focus sprSent, 1, 2
      GoTo ExitHere
    End If
    sprSent.GetText 3, 1, tmp
    Row = Val(tmp)
    sprSent.GetText 1, 1, tmp
    sprItem.SetText eItem.msentItem, Row, tmp
    sprSent.GetText 2, 1, tmp
    sprItem.SetText eItem.mSentDesc, Row, tmp
    Call Spread_Lock(sprItem, False, eItem.mSentQty, Row, eItem.mTarDate, Row)
  fraSent.Visible = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtpenqdate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)

  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
    SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtprefdate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyCode = 13 Then
    SendKeys "{tab}"
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  If TypeOf Me.ActiveControl Is fpSpread Then
 GoTo ExitHere
End If
  If KeyAscii = 13 Then
    SendKeys "{tab}"
  End If
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

'  Loadenqtype '''Loading all the enquiry types to combo
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  cmbCustomerType.ListIndex = 0
  dtpEnqDate = CurrDate
  dtpRefDate = CurrDate
  sprItem.Enabled = False
  cmdSave.Enabled = False
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub sprItem_ComboSelChange(sender as object, e as ComboSelChangeEvent)
  Dim retVal As Variant
  
  Try
  
  If Col = eItem.mItemType Then
    sprItem.GetText Col, Row, retVal
    If Trim$(retVal) = "New Item" Then
      Call Spread_Lock(sprItem, False, eItem.mrecdItem, Row, eItem.mRecdDesc, Row)
      Call Spread_Lock(sprItem, False, eItem.msentItem, Row, eItem.mRate, Row)
      Call Spread_Lock(sprItem, True, eItem.mSentUOM, Row, eItem.mSentUOM, Row)
      sprItem.ClearRange eItem.mrecdItem, Row, eItem.mAssGoodsValue, Row, True
      sprItem.ClearRange eItem.mFirstCol, Row, eItem.mFirstCol, Row, True
      sprItem.SetText eItem.mRatePer, Row, "1"
      sprItem.SetText eItem.mFirstCol, Row, Val(Row)
      sprItem.SetText eItem.mRelation, Row, "1-1"
      fraRecd.Visible = False
      fraSent.Visible = False
    ElseIf Trim$(retVal) = "Existing Item" Then
      Call Spread_Lock(sprItem, True, eItem.mrecdItem, Row, eItem.mRecdDesc, Row)
      Call Spread_Lock(sprItem, True, eItem.msentItem, Row, eItem.mRate, Row)
      Call Spread_Lock(sprItem, True, eItem.mSentUOM, Row, eItem.mSentUOM, Row)
      Call Spread_Lock(sprItem, False, eItem.mRatePer, Row, eItem.mRate, Row)
      sprItem.ClearRange eItem.mrecdItem, Row, eItem.mAssGoodsValue, Row, True
      sprItem.ClearRange eItem.mFirstCol, Row, eItem.mFirstCol, Row, True
      sprItem.SetText eItem.mFirstCol, Row, Val(Row)
      fraRecd.Visible = False
      fraSent.Visible = False
    Else
      Call Spread_Lock(sprItem, True, eItem.mrecdItem, Row, eItem.mRecdDesc, Row)
      Call Spread_Lock(sprItem, True, eItem.msentItem, Row, eItem.mRate, Row)
      Call Spread_Lock(sprItem, True, eItem.mSentUOM, Row, eItem.mSentUOM, Row)
      sprItem.ClearRange eItem.mrecdItem, Row, eItem.mAssGoodsValue, Row, True
      sprItem.ClearRange eItem.mFirstCol, Row, eItem.mFirstCol, Row, True
      sprItem.SetText eItem.mRatePer, Row, "1"
      sprItem.SetText eItem.mFirstCol, Row, Val(Row)
      sprItem.SetText eItem.mRelation, Row, "N-1"
      sprRecd.ClearRange 1, 1, -1, -1, True
      sprRecd.SetText 3, 1, Val(Row)
      sprRecd.MaxRows = 1
      fraRecd.Visible = True
      Spread_Set_Focus sprRecd, 1, 1
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub sprItem_DblClick(sender as object, e as DblClickEventArgs)
  Dim retVal As Variant
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  sprItem.GetText eItem.mItemType, Row, retVal
  If Trim$(retVal) = "Existing Item" Then
    If Col = eItem.mrecdItem Then    'Select Sent ItemCode means One To One / One To Many
      Call OneToOneItemHelp(Row)
    ElseIf Col = eItem.msentItem Then 'Select Recd ItemCode means Many to One
      Call ManyToOneItemHelp(Row)
    End If
  Else
    If Col = eItem.mRecdUOM Or Col = eItem.mSentUOM Then
      If Col = eItem.mSentUOM Then
        sprItem.GetText eItem.msentItem, Row, retVal
      Else
        sprItem.GetText eItem.mrecdItem, Row, retVal
      End If
      If Len(retVal) > 0 Then
        Sql = "select SDesc,NoofDecimals from INV_UOM_MASTER where Branchid ='" & gBranchID & "'"
        Sql = Sql & " and compid='" & gCompID & "'"
        If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of UOM's", "UOM") = True Then
          sprItem.SetText Col, Row, Trim$(getvalue(1))
        End If
      End If
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub OneToOneItemHelp(ByVal fRow As Long)
  Dim Sql As String
  Dim retVal As Variant
  Dim sParentCode As String

  Try

    sprItem.GetText eItem.mrecdItem, fRow, retVal
    If Trim$(retVal) = "" Then   'already saved itemcode should not be changed
      sprItem.GetText eItem.mFirstCol, fRow, retVal
      If Len(Trim$(retVal)) > 0 Then
        Sql = "Select distinct m.ItemCode, n.ItemDescription"
        Sql = Sql & " from (Select distinct a.ItemCode, (Select Count(*) from BOM_MPS_Master b where b.ItemCode=a.ItemCode and b.BranchID=a.BranchID and b.CompID=a.CompID) ProcessCount, "
        Sql = Sql & " (Select Count(*) from BOM_Structure_Dtl c where c.ParentCode=a.ItemCode and c.BranchID=a.BranchID and c.CompID=a.CompID) ChildCount"
        Sql = Sql & " from BOM_MPS_Master a where a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "') m "
        Sql = Sql & " Left outer Join Inv_Material_Master n On m.ItemCode=n.ItemCode"
        Sql = Sql & " where (m.ProcessCount+m.ChildCount) >= 2 and m.ChildCount <= 1"
        Sql = Sql & " and n.BranchID='" & gBranchID & "' and n.CompID='" & gCompID & "' and n.status = 1"
        If FetchValue_New(False, db.ConnectionString, Sql, "*", "List of Items", "Item Code", "Description") = True Then
          sParentCode = getvalue(1)
          
          'Show the Process Details & Child Item for the Selected Item
          Sql = "Select a1.StageCode, (Select top 1 b.OperationName from BOM_Operation_Master b where a1.OperationCode=b.OperationCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID) as StageDesc,"
          Sql = Sql & " (Select b.UOMDesc from Material_Master b where a1.StageCode=b.ItemCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID and b.Status = 1) as UOM, "
          Sql = Sql & " a1.ItemCode, (Select top 1 b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ItemCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ItemDesc"
          Sql = Sql & " from BOM_MPS_Master a1"
          Sql = Sql & " where a1.ItemCode='" & sParentCode & "' and a1.OperationSequence > 0"
          Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "'"
          Sql = Sql & " Union All"
          Sql = Sql & " Select a1.ChildCode, b1.ItemDescription as ChildDesc, b1.PurchaseUOMDesc as UOM,"
          Sql = Sql & " a1.ParentCode, (Select b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ParentCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ParentDesc"
          Sql = Sql & " from BOM_Structure_Dtl a1 Left Outer Join Material_Master b1 On b1.ItemCode=a1.ChildCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID"
          Sql = Sql & " where a1.ParentCode = (Select a2.ParentCode from BOM_Structure_Dtl a2 Where a2.ParentCode='" & sParentCode & "'"
          Sql = Sql & " and a2.BranchID='" & gBranchID & "' and a2.CompID='" & gCompID & "'"
          Sql = Sql & " Group by a2.ParentCode Having Count(*) = 1)"
          Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "'and b1.Status = 1 "
          If FetchValue_New(False, db.ConnectionString, Sql, "11011", "List of Process - Select Received Item", "Stage Code", "Description", "Item Code", "Description") = True Then
            sprItem.SetText eItem.mrecdItem, fRow, getvalue(1)
            sprItem.SetText eItem.mRecdDesc, fRow, getvalue(2)
            sprItem.SetText eItem.mRecdUOM, fRow, getvalue(3)
             Call Spread_Lock(sprItem, True, eItem.mItemType, fRow, eItem.mItemType, fRow)
            'Show the Process Details for that Item greater than the selected process
            Sql = "Select a1.StageCode, (Select top 1 b.OperationName from BOM_Operation_Master b where a1.OperationCode=b.OperationCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID) as StageDesc,"
            Sql = Sql & " a3.PurchaseUOMDesc as UOM, a3.PurchaseUOMDecimal as NoOfDec,"
            Sql = Sql & " a1.ItemCode, (Select top 1 b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ItemCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ItemDesc"
            Sql = Sql & " from BOM_MPS_Master a1 Left Outer Join Material_Master a3 On a1.StageCode=a3.ItemCode and a1.BranchID=a3.BranchID and a1.CompID=a3.CompID"
            Sql = Sql & " Where a1.ItemCode='" & sParentCode & "' and (a1.OperationSequence = 0 Or a1.OperationSequence > isNull((select top 1 OperationSequence from bom_mps_master a2 where stagecode='" & getvalue(1) & "'"
            Sql = Sql & " and a2.BranchID='" & gBranchID & "' and a2.CompID='" & gCompID & "'), -1))"
            Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "' and a3.Status = 1"
            If FetchValue_New(False, db.ConnectionString, Sql, "110011", "List of Process - Select Send Item", "Stage Code", "Description", "Item Code", "Description") = True Then
              sprItem.SetText eItem.msentItem, fRow, getvalue(1)
              If Spread_Clear_Duplicate(sprItem, fRow, eItem.mrecdItem, eItem.msentItem) = True Then
                sprItem.MaxRows = sprItem.DataRowCnt + 1
                sprItem.SetActiveCell eItem.mrecdItem, sprItem.ActiveRow
              Else
                sprItem.SetText eItem.mSentDesc, fRow, getvalue(2)
                sprItem.SetText eItem.mSentUOM, fRow, getvalue(3)
'                Spread_Decimal_Set sprItem, eItem.mRatePer, eItem.mRatePer, fRow, fRow, Val(getvalue(4))
                Call Spread_Set_Decimal(sprItem, Val(getvalue(4)), eItem.mRatePer, fRow)
                
                sprItem.SetText eItem.mFirstCol, fRow, fRow
                sprItem.SetText eItem.mRelation, fRow, "1-1"
                sprItem.SetText eItem.mRatePer, fRow, 1
              
                sprItem.MaxRows = sprItem.DataRowCnt + 1
                Call Spread_Lock(sprItem, False, eItem.mSentQty, fRow, eItem.mAssRate, fRow)
                sprItem.SetActiveCell eItem.mRatePer, sprItem.ActiveRow
              End If
            Else
              sprItem.ClearRange 1, fRow, -1, fRow, True
            End If    'Recd Item Help
          End If      'Sent Item Help
        End If        'ItemCode Help
      End If          'No Recd Item Found
    End If            'Already Saved Item
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub ManyToOneItemHelp(ByVal fRow As Long)
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim retVal As Variant
  Dim sMaxRevNo As Integer
  Dim sParentCode As String
  Dim sStageCode As String
  Dim sChildCode As String
  Dim DubLicate As Boolean
  Dim SRow As Long
  Dim i As Integer
  
  Try
  DubLicate = False
  
    sprItem.GetText eItem.mrecdItem, fRow, retVal
    If Trim$(retVal) = "" Then   'already saved itemcode should not be changed
      sprItem.GetText eItem.mFirstCol, fRow, retVal
      If Len(Trim$(retVal)) > 0 Then          'If Sent Item is not already selected
      
        Sql = "Select distinct isNull((Select TOP 1 n.ItemCode from BOM_MPS_Master n where n.StageCode=a.ParentCode and n.BranchID='" & gBranchID & "' and n.CompID='" & gCompID & "'), a.ParentCode) as ParentCode, "
        Sql = Sql & " b.ItemDescription, b.PurchaseUOMDesc as UOM, Max(a.RevisionNo) as RevNo"
        Sql = Sql & " from BOM_STRUCTURE_HDR a Left Outer Join MATERIAL_MASTER b On a.ParentCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
        Sql = Sql & " where a.status=1 and a.branchid='" & gBranchID & "' and a.compid='" & gCompID & "'"
        Sql = Sql & " group by a.ParentCode, b.ItemDescription, b.PurchaseUOMDesc, b.PurchaseUOMDecimal"
      
        If FetchValue_New(False, db.ConnectionString, Sql, "110", "List of Items", "Item Code", "Description") = True Then
          sParentCode = Trim$(getvalue(1))
          sMaxRevNo = Val(getvalue(4))
          
          'Show the Recd Item Process Structure for the selected Item
          Sql = "Select a1.StageCode, (Select b.OperationName from BOM_Operation_Master b where a1.OperationCode=b.OperationCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID) as StageDesc,"
          Sql = Sql & " b.PurchaseUOMDesc as UOM, b.PurchaseUOMDecimal as NoOfDec,"
          Sql = Sql & " a1.ItemCode, (Select b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ItemCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ItemDesc"
          Sql = Sql & " from BOM_MPS_Master a1 Left Outer Join Material_Master b On a1.StageCode=b.ItemCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID"
          Sql = Sql & " where a1.ItemCode='" & sParentCode & "' and b.Status = 1 "
          Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "'"
'          sql = sql & " union (Select distinct a2.ParentCode,b2.ItemDescription, b2.PurchaseUOMDesc as UOM, "
'          sql = sql & " Max(a.RevisionNo) as RevNo from BOM_STRUCTURE_HDR a Left Outer Join MATERIAL_MASTER b "
'          sql = sql & " On a2.ParentCode=b2.ItemCode and a2.BranchID=b2.BranchID and a2.CompID=b2.CompID"
'          sql = sql & " where a2.ParentCode='" & sParentCode & "' and a2.status=1 and a2.branchid='" & gBranchID & "' and a2.compid='" & gCompID & "'"
'          sql = sql & " group by a2.ParentCode, b2.ItemDescription, b2.PurchaseUOMDesc, b2.PurchaseUOMDecimal)"
          If FetchValue_New(False, db.ConnectionString, Sql, "110011", "List of Process - Select Send Item", "Stage Code", "Description", "Item Code", "Description") = True Then
            sprItem.SetText eItem.msentItem, fRow, getvalue(1)
            sprItem.SetText eItem.mSentDesc, fRow, CheckNull(getvalue(2))
            sprItem.SetText eItem.mSentUOM, fRow, CheckNull(getvalue(3))
            sprItem.SetText eItem.mRatePer, fRow, 1
            sStageCode = getvalue(1)
            sprItem.SetText eItem.mFirstCol, fRow, Val(fRow)
            
            Spread_Lock sprItem, False, eItem.mSentQty, fRow, eItem.mRate, fRow
            
            Call Spread_Set_Decimal(sprItem, Val(CheckNull(getvalue(4), 0)), eItem.mRatePer, fRow)
          
            'Fill the Child Items for the selected Item Structure Details in SentItem
            Sql = "Select b.ChildCode, c.ItemDescription as ChildDesc, c.UOMDesc as UOM"
            Sql = Sql & " from BOM_Structure_Dtl b Left Outer Join MATERIAL_MASTER c "
            Sql = Sql & " On b.ChildCode=c.ItemCode and b.BranchID=c.BranchID and b.CompID=c.CompID"
            Sql = Sql & " Where b.ParentCode=(Select d.StageCode from BOM_MPS_Master d "
            Sql = Sql & " where d.ItemCode='" & sParentCode & "' and d.OperationSequence=1 and d.BranchID=b.BranchID and d.CompID=b.CompID)"
            Sql = Sql & " and b.RevisionNo=" & sMaxRevNo & " and b.BranchID='" & gBranchID & "' and b.CompID='" & gCompID & "'"
            Sql = Sql & " Order by b.ChildCode"
            If FetchValue_New(True, db.ConnectionString, Sql, "110011", "Select Send Item", "Item Code", "Description") = True Then
              sprItem.MaxRows = sprItem.DataRowCnt + UBound(getvalue, 1)
              sChildCode = ""
              For i = 1 To UBound(getvalue, 1)
                sChildCode = sChildCode + getvalue(i, 1)
              Next i
              
                SRow = sprItem.SearchCol(eItem.msentItem, 0, sprItem.MaxRows, Trim$(sStageCode), SearchFlagsValue)
                Do While SRow <> -1
                  If Group_Child(SRow, sStageCode) = sChildCode Then
                    DubLicate = True
                    sprItem.DeleteRows fRow, 1
                    Exit Do
                  End If
                  SRow = sprItem.SearchCol(eItem.msentItem, SRow, sprItem.MaxRows, Trim$(sStageCode), SearchFlagsValue)
                Loop
              If DubLicate = False Then
                i = 1
                For rCnt = sprItem.DataRowCnt To UBound(getvalue, 1) + sprItem.DataRowCnt - 1
                  sprItem.SetText eItem.mFirstCol, rCnt, Val(fRow)
                  sprItem.SetText eItem.mrecdItem, rCnt, getvalue(i, 1)
                  sprItem.SetText eItem.mRecdDesc, rCnt, getvalue(i, 2)
                  sprItem.SetText eItem.mRecdUOM, rCnt, getvalue(i, 3)
                  sprItem.SetText eItem.mRelation, rCnt, "N-1"
'                  .SetText eItem.mItemType, rCnt, "Existing Item"
                  Call Spread_Lock(sprItem, True, eItem.mItemType, rCnt, eItem.mItemType, rCnt)
                  i = i + 1
                Next rCnt
              End If
              sprItem.MaxRows = sprItem.DataRowCnt + 1
            Else
              sprItem.DeleteRows fRow, 1
              sprItem.MaxRows = sprItem.DataRowCnt + 1
            End If    'Child Item Code
          End If      'RecdCode Process Help
        End If        'RecdCode Parent Help
      End If          'No Sent Item Found
    End If            'Already Saved Item
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub sprItem_EditMode(sender as object, e as EditModeEventArgs)
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 1 Then GoTo ExitHere
  If Val(SpVal(sprItem, Row, eItem.mSentQty)) > 0 Then
    sprItem.SetText eItem.mGoodsValue, Row, (Val(SpVal(sprItem, Row, eItem.mSentQty)) / Val(SpVal(sprItem, Row, eItem.mRatePer))) * Val(SpVal(sprItem, Row, eItem.mRate))
    sprItem.SetText eItem.mAssGoodsValue, Row, (Val(SpVal(sprItem, Row, eItem.mSentQty)) / Val(SpVal(sprItem, Row, eItem.mRatePer))) * Val(SpVal(sprItem, Row, eItem.mAssRate))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_KeyDown(sender as object, e as KeyDownEventArgs)
  Dim retVal As Variant
  Dim SRow As Integer
  Dim mRow As Integer
  Dim nRow As Integer
  Dim i As Integer
  
  Try
  Me.MousePointer = vbHourglass
  
    Select Case KeyCode
      Case vbKeyF2
        sprItem.GetText eItem.mItemType, sprItem.ActiveRow, retVal
        If Trim$(retVal) = "Existing Item" Then
          sprItem_DblClick sprItem.ActiveCol, sprItem.ActiveRow
        ElseIf sprItem.ActiveCol = eItem.mRecdUOM Or sprItem.ActiveCol = eItem.mSentUOM Then
          Call sprItem_DblClick(sprItem.ActiveCol, sprItem.ActiveRow)
        End If
      Case 46
        If sprItem.ActiveCol = eItem.mItemType And Shift = 1 Then
          sprItem.GetText eItem.mFirstCol, sprItem.ActiveRow, retVal
          If Val(retVal) <> 0 Then
            SRow = 0
            SRow = sprItem.SearchCol(eItem.mFirstCol, 0, sprItem.MaxRows, Trim$(retVal), SearchFlagsValue)
            mRow = 0
            Do While SRow <> -1
              sprItem.DeleteRows SRow, 1
mRow = mRow + 1
nRow = SRow
              SRow = sprItem.SearchCol(eItem.mFirstCol, 0, sprItem.MaxRows, Trim$(retVal), SearchFlagsValue)
            Loop
            For i = 1 To sprItem.DataRowCnt
              If i >= nRow Then
                sprItem.GetText eItem.mFirstCol, i, retVal
                sprItem.SetText eItem.mFirstCol, i, Val(retVal) - mRow
              End If
             Next i
              If sprItem.DataRowCnt = 0 Then
              sprItem.SetText eItem.mItemType, 1, "Existing Item"
              sprItem.SetText eItem.mFirstCol, 1, "1"
            End If
           
            sprItem.MaxRows = sprItem.DataRowCnt + 1
            
          End If
        End If
      Case 13
        sprItem.GetText eItem.mItemType, sprItem.ActiveRow, retVal
        If Trim$(retVal) = "New Item" And sprItem.ActiveCol = eItem.mGoodsValue Then
          sprItem.GetText eItem.mrecdItem, sprItem.ActiveRow, retVal
          If Trim$(retVal) <> "" Then
            sprItem.MaxRows = sprItem.DataRowCnt + 1
          End If
        End If
    End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprItem_LeaveCell(sender as object, e as LeaveCellEventArgs)
  Dim tmp As Variant
  
    sprItem.GetText eItem.mFirstCol, NewRow, tmp
    sprItem.GetText eItem.mItemType, Val(tmp), tmp
    If Trim$(tmp) = "Existing item" Then
      If NewCol = eItem.mrecdItem Or NewCol = eItem.msentItem Then
        Spread_Set_HeaderColor sprItem, anSelectOnly, anCol, NewCol, , NewCol
        Spread_Lock sprItem, True, NewCol, NewRow, NewCol, NewRow
      ElseIf NewCol = eItem.mRecdUOM Or NewCol = eItem.mSentUOM Then
        Spread_Set_HeaderColor sprItem, anEditable, anCol, NewCol, , NewCol
        Spread_Lock sprItem, True, NewCol, NewRow, NewCol, NewRow
      End If
    ElseIf Trim$(tmp) = "NEW ITEM" Or Trim$(tmp) = "NEW ITEM(N-1)" Then
      If NewCol = eItem.mRecdUOM Or NewCol = eItem.mSentUOM Then
        Spread_Set_HeaderColor sprItem, anSelectOnly, anCol, NewCol, , NewCol
        Spread_Lock sprItem, True, NewCol, NewRow, NewCol, NewRow
      ElseIf NewCol = eItem.mrecdItem Or NewCol = eItem.msentItem Then
        Spread_Set_HeaderColor sprItem, anEditable, anCol, NewCol, , NewCol
      End If
    End If
End Sub

Private Sub sprRecd_KeyDown(sender as object, e as KeyDownEventArgs)
  Dim tmp As Variant
  
  Try
  
  If KeyCode = 13 And sprRecd.ActiveCol = 2 Then
    sprRecd.MaxRows = sprRecd.DataRowCnt + 1
'    Spread_Set_Focus sprRecd, sprRecd.MaxRows, 1
  End If
  If KeyCode = 46 And sprRecd.ActiveCol = 1 Then
    sprRecd.GetText 3, 1, tmp
    sprRecd.DeleteRows sprRecd.ActiveRow, 1
    sprRecd.MaxRows = sprRecd.DataRowCnt + 1
    sprRecd.SetText 3, 1, tmp
'    Spread_Set_Focus sprRecd, sprRecd.MaxRows, 1
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub sprRecd_LeaveCell(sender as object, e as LeaveCellEventArgs)
' On Error Resume Next
  If Col = 1 Then
    Spread_Clear_Duplicate sprRecd, Row, 1
  End If
End Sub

Private Sub txtCustomer_Change()
' On Error Resume Next
  
  If txtCustomer = "" Then
 txtCustomer.Tag = ""
End If
End Sub

Private Sub txtcustomer_DblClick()
  Dim Sql As String

  Try
  
  Me.MousePointer = vbHourglass
  
  If cmbCustomerType = "New Customer" Then
 GoTo ExitHere
End If
  
  Sql = " select distinct (select b.PartyName from FAS_PARTY_MASTER b where b.PartyCode=a.CustomerCode"
  Sql = Sql + " and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.CustomerCode"
  Sql = Sql + " from SAL_JOBWORK_ITEMMASTER a where a.Branchid='" & gBranchID & "' "
  Sql = Sql + " and a.Compid='" & gCompID & "'"
  Sql = Sql & " order by PartyName"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Customers", "Customer Name") = True Then
     txtCustomer.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
     txtCustomer = IIf(IsNull(getvalue(1)), "", getvalue(1))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcustomer_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
  
  If KeyCode = vbKeyF2 Then
     txtcustomer_DblClick
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtCustomer_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  KeyAscii = ToUpper(KeyAscii)
End Sub

Private Sub txtEnqNo_DblClick()
  Dim Sql As String
  

  Try
  Me.MousePointer = vbHourglass

  Sql = "Select Distinct a.EnquiryType,a.EnquiryNo,a.EnquiryDate,a.Revisionno,ReferenceNo,ReferenceDate,CustomerType,CustomerCode,"
  Sql = Sql & " CustomerName,Remarks from SAL_VW_JobworkEnquiry_Hdr a Left outer join SAL_VW_JobworkEnquiry_Dtl b"
  Sql = Sql & " on a.EnquiryType=b.EnquiryType and a.EnquiryNo=b.EnquiryNo and a.EnquiryDate=b.EnquiryDate"
  Sql = Sql & " and a.Branchid=b.Branchid and a.Compid=b.Compid"
  Sql = Sql & " where b.QuotationType='' and QuotationDate is null and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "1111000010", "List of Enquirys", "Enquiry Type", "Enquiry No", "Enquiry Date", "Revisionno", "Customer Name") = True Then
    txtEnqType = getvalue(1)
    txtEnqNo = getvalue(2)
    dtpEnqDate = getvalue(3)
    txtEnqNo.Tag = Val(getvalue(4)) + 1 'Revno
    txtRefNo = getvalue(5)
    dtpRefDate = getvalue(6)
    cmbCustomerType.ListIndex = IIf(Trim$(getvalue(7)) = "E", 0, 1)
    txtCustomer.Tag = getvalue(8)
    txtCustomer = getvalue(9)
    txtremarks = getvalue(10)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtrefno_KeyPress(ByVal KeyAscii As Integer)
  
  Try
  Me.MousePointer = vbHourglass

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtremarks_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Design_sprSent(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprSent

End Sub

Private Sub Design_sprRecd(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprRecd


sprRecd.KeyDown +=  new EventHandler(sprRecd_KeyDown)
sprRecd.LeaveCell +=  new EventHandler(sprRecd_LeaveCell)

End Sub

Private Sub Design_sprItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprItem


sprItem.ComboSelChange +=  new EventHandler(sprItem_ComboSelChange)
sprItem.DblClick +=  new EventHandler(sprItem_DblClick)
sprItem.EditMode +=  new EventHandler(sprItem_EditMode)
sprItem.KeyDown +=  new EventHandler(sprItem_KeyDown)
sprItem.LeaveCell +=  new EventHandler(sprItem_LeaveCell)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprSent("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmJobWorkEnqModify.frm", sprSent, "")

	Call Design_sprRecd("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmJobWorkEnqModify.frm", sprRecd, "")

	Call Design_sprItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmJobWorkEnqModify.frm", sprItem, "")

End Sub
