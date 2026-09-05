VERSION 5.00
Begin VB.Form frmCustomerJobworkMaster 
   Caption         =   "Customer Jobwork Master"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
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
   ScaleHeight     =   6585
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.Frame fraVendor 
      Height          =   690
      Left            =   2655
      TabIndex        =   7
      Top             =   495
      Width           =   6675
      Begin VB.TextBox txtVendCode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Select the  Customer. Press F2 or Double Click for Help."
         Top             =   300
         Width           =   5865
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   330
         Left            =   5985
         TabIndex        =   1
         ToolTipText     =   "Click to OK"
         Top             =   285
         Width           =   585
      End
      Begin VB.Label lblVendCode 
         AutoSize        =   -1  'True
         Caption         =   "Customer"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   8
         Top             =   120
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   4740
      Width           =   1300
   End
Begin AceSpread sprItem
      Height          =   2610
      Left            =   390
      TabIndex        =   3
      ToolTipText     =   "Enter the Jobwork Item Details"
      Top             =   2010
      Width           =   11190
      _Version        =   458752
      _ExtentX        =   19738
      _ExtentY        =   4604
      _StockProps     =   64
      Enabled         =   0   'False
      ButtonDrawMode  =   4
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
      MaxCols         =   20
      MaxRows         =   1
      SelectBlockOptions=   0
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5342
      TabIndex        =   5
      ToolTipText     =   "Click to Clear the Data"
      Top             =   4740
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   4050
      TabIndex        =   4
      ToolTipText     =   "Click to Save the Data"
      Top             =   4740
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   4312
      TabIndex        =   9
      Top             =   1170
      Width           =   3360
      Begin VB.TextBox txtDefaultFormula 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Select the Defalut Formula. Press F2 or Double Click for Help."
         Top             =   300
         Width           =   3225
      End
      Begin VB.Label lblVendCode 
         AutoSize        =   -1  'True
         Caption         =   "Default Formula"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   120
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmCustomerJobworkMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'For Item Spread Column Numbers
Private Enum eItem
  mFirstCol = 1
  mrecdItem = 2
  mRecdDesc = 3
  mRecdUOM = 4
  msentItem = 5
  mSentDesc = 6
  mPartNo = 7
  mSentUOM = 8
  mRatePer = 9
  mRate = 10
  mAssRate = 11
  mValidFrom = 12
  mValidTo = 13
  mQCCheck = 14
  mFrmCode = 15
  mFrmName = 16
  mStatus = 17
  mRelation = 18
  mItemGroupKey = 19
  mRecdItemCheck = 20
End Enum

Private frowno As Long
Private fLevel_Ind  As Long
Private fPrev_Level_Ind As Long
Private fhighest_Level  As Long
Private reccount As Long

Private Function ItemGroupKey(ByVal sChildCode As String,ByVal  sParentCode As String) As Long
  Dim Sql As String
  Dim rs As Recordset
  
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

End Try
End Function

Private Function Group_Child(ByVal fRow As Long,ByVal  ParentCode As String) As String
  Dim i As Integer
  Dim retVal As Variant
  Dim Row As Long
  Dim sRow As Long
  Dim ChildCode As String
  
  Try
  Me.MousePointer = vbHourglass
  
    sRow = sprItem.SearchCol(eItem.mFirstCol, 0, sprItem.MaxRows, Val(fRow), SearchFlagsValue)
    ChildCode = ""
    Do While sRow <> -1
      sprItem.GetText eItem.mrecdItem, sRow, retVal
      ChildCode = ChildCode + retVal
      sRow = sprItem.SearchCol(eItem.mFirstCol, sRow, sprItem.MaxRows, Trim$(fRow), SearchFlagsValue)
    Loop
  
  Group_Child = ChildCode
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Function

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
  Dim sRow As Long
  Dim i As Integer
  
  Try
  DubLicate = False
  
  
    If fRow > sprItem.DataRowCnt Then     'already saved itemcode should not be changed
      sprItem.GetText eItem.mFirstCol, fRow, retVal
      If Len(Trim$(retVal)) = 0 Then          'If Sent Item is not already selected
      
        Sql = "Select distinct isNull((Select TOP 1 n.ItemCode from BOM_MPS_Master n where n.StageCode=a.ParentCode and n.BranchID='" & gBranchID & "' and n.CompID='" & gCompID & "'), a.ParentCode) as ParentCode, "
        Sql = Sql & " b.ItemDescription, b.PurchaseUOMDesc as UOM, Max(a.RevisionNo) as RevNo"
        Sql = Sql & " from BOM_STRUCTURE_HDR a Left Outer Join MATERIAL_MASTER b On a.ParentCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
        Sql = Sql & " where a.status=1 and b.Status = 1 and a.branchid='" & gBranchID & "' and a.compid='" & gCompID & "'"
        Sql = Sql & " group by a.ParentCode, b.ItemDescription, b.PurchaseUOMDesc, b.PurchaseUOMDecimal"
      
        If FetchValue_New(False, db.ConnectionString, Sql, "110", "List of Items", "Item Code", "Description") = True Then
          sParentCode = Trim$(getvalue(1))
          sMaxRevNo = Val(getvalue(4))
          
          'Show the Recd Item Process Structure for the selected Item
          Sql = "Select a1.StageCode, (Select b.OperationName from BOM_Operation_Master b where a1.OperationCode=b.OperationCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID) as StageDesc,"
          Sql = Sql & " b.PurchaseUOMDesc as UOM, b.PurchaseUOMDecimal as NoOfDec,"
          Sql = Sql & " a1.ItemCode, (Select b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ItemCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ItemDesc"
          Sql = Sql & " from BOM_MPS_Master a1 Left Outer Join Material_Master b On a1.StageCode=b.ItemCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID"
          Sql = Sql & " where a1.ItemCode='" & sParentCode & "'"
          Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "' and b.Status = 1 "
'          sql = sql & " union (Select distinct a2.ParentCode,b2.ItemDescription, b2.PurchaseUOMDesc as UOM, "
'          sql = sql & " Max(a.RevisionNo) as RevNo from BOM_STRUCTURE_HDR a Left Outer Join MATERIAL_MASTER b "
'          sql = sql & " On a2.ParentCode=b2.ItemCode and a2.BranchID=b2.BranchID and a2.CompID=b2.CompID"
'          sql = sql & " where a2.ParentCode='" & sParentCode & "' and a2.status=1 and a2.branchid='" & gBranchID & "' and a2.compid='" & gCompID & "'"
'          sql = sql & " group by a2.ParentCode, b2.ItemDescription, b2.PurchaseUOMDesc, b2.PurchaseUOMDecimal)"
          If FetchValue_New(False, db.ConnectionString, Sql, "110011", "List of Process - Select Send Item", "Stage Code", "Description", "Item Code", "Description") = True Then
            sprItem.SetText eItem.msentItem, fRow, getvalue(1)
            '.SetText eItem.mSentDesc, fRow, CheckNull(getvalue(2))
            sprItem.SetText eItem.mSentDesc, fRow, CheckNull(getvalue(6))
            sprItem.SetText eItem.mSentUOM, fRow, CheckNull(getvalue(3))
            sprItem.SetText eItem.mRatePer, fRow, 1
            sStageCode = getvalue(1)
            sprItem.SetText eItem.mFirstCol, fRow, Val(fRow)
            
            Spread_Lock sprItem, False, eItem.mPartNo, fRow, eItem.mStatus, fRow
            Spread_Lock sprItem, True, eItem.mSentUOM, fRow, eItem.mSentUOM, fRow
            
            'Fill the Child Items for the selected Item Structure Details in SentItem
            Sql = "Select b.ChildCode, c.ItemDescription as ChildDesc, c.UOMDesc as UOM"
            Sql = Sql & " from BOM_Structure_Dtl b Left Outer Join MATERIAL_MASTER c "
            Sql = Sql & " On b.ChildCode=c.ItemCode and b.BranchID=c.BranchID and b.CompID=c.CompID"
            Sql = Sql & " Where b.ParentCode=(Select d.StageCode from BOM_MPS_Master d "
            Sql = Sql & " where d.ItemCode='" & sParentCode & "' and d.OperationSequence=(Select Case When Max(OperationSequence) >=1 Then 1 Else 0 End"
            Sql = Sql & " From Bom_Mps_Master Where ItemCode = '" & sParentCode & "'"
            Sql = Sql & " and BranchId = '" & gBranchID & "' and CompID = '" & gCompID & "')"
            Sql = Sql & " and d.BranchID=b.BranchID and d.CompID=b.CompID)"
            Sql = Sql & " and b.RevisionNo=" & sMaxRevNo & " and c.Status = 1 and b.BranchID='" & gBranchID & "' and b.CompID='" & gCompID & "'"
            Sql = Sql & " Order by b.ChildCode"
            If FetchValue_New(True, db.ConnectionString, Sql, "110011", "Select Send Item", "Item Code", "Description") = True Then
              sprItem.MaxRows = sprItem.DataRowCnt + UBound(getvalue, 1)
              sChildCode = ""
              For i = 1 To UBound(getvalue, 1)
                sChildCode = sChildCode + getvalue(i, 1)
              Next i
              
                sRow = sprItem.SearchCol(eItem.msentItem, 0, sprItem.MaxRows, Trim$(sStageCode), SearchFlagsValue)
                Do While sRow <> -1
                  If Group_Child(sRow, sStageCode) = sChildCode Then
                    DubLicate = True
                    sprItem.DeleteRows fRow, 1
                    Exit Do
                  End If
                  sRow = sprItem.SearchCol(eItem.msentItem, sRow, sprItem.MaxRows, Trim$(sStageCode), SearchFlagsValue)
                Loop
              If DubLicate = False Then
                i = 1
                For rCnt = sprItem.DataRowCnt To UBound(getvalue, 1) + sprItem.DataRowCnt - 1
                  sprItem.SetText eItem.mFirstCol, rCnt, Val(fRow)
                  sprItem.SetText eItem.mrecdItem, rCnt, getvalue(i, 1)
                  sprItem.SetText eItem.mRecdDesc, rCnt, getvalue(i, 2)
                  sprItem.SetText eItem.mRecdUOM, rCnt, getvalue(i, 3)
                  sprItem.SetText eItem.mRelation, rCnt, "N-1"
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

Private Sub cmdClear_Click()
  Try
  Me.MousePointer = vbHourglass
  
  Form_Clear Me
  fraVendor.Enabled = True
  sprItem.Enabled = False
  cmdSave.Enabled = False
  sprItem.MaxRows = 1
  txtVendCode.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub cmdExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  Dim rCnt As Long
  Dim ValueRow As Integer
  Dim PrevItemGroup As Integer
  
  Try
  Me.MousePointer = vbHourglass

  If Trim$(txtVendCode) = "" Then
    sShowMessage "Select the Vendor Code"
    txtVendCode.SetFocus
    GoTo ExitHere
  End If
  
  Sql = "Select distinct a.FrmId, a.FrmLName from COM_FormulaMast_Hdr a Right Outer Join" 'a.FormulaCode,
  Sql = Sql & " FAS_Party_Master b On a.FrmId = b.CustomerFormulaCode and a.BranchID = b.BranchID "
  Sql = Sql + " and a.CompID = b.CompID where a.ModuleCode = '" & gModuleCode & "' and b.PartyCode=" & Val(txtVendCode.Tag) & ""
  Sql = Sql + " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  If Not snap.EOF Then
    txtDefaultFormula.Tag = snap("FrmId")
    txtDefaultFormula.Text = snap("FrmLName")
  Else
    sShowMessage "No Default Formula Found. So select the Default Formula & also open and update it in Party Master."
    Call txtDefaultFormula_DblClick
  End If
  
  Sql = "Select a.ItemGroupKey, a.SentItemCode, b.ItemDescription as SentDesc, b.PurchaseUOMDesc as SentUOM, a.ReceivedItemCode, c.ItemDescription as RecdDesc, c.PurchaseUOMDesc as RecdUOM,"
  Sql = Sql & " a.RelationType, a.PartNo, c.PurchaseUOMDecimal as NoOfDec, a.RatePer, a.rate, a.ValidFrom, a.ValidTo, a.QcCheck, a.isActive, a.FormulaCode,"
  Sql = Sql & " (select top 1 d.FrmLName from COM_FormulaMast_Hdr d where d.ModuleCode = '" & gModuleCode & "' and d.FrmId = a.FormulaCode "
  Sql = Sql & " and d.Branchid = a.Branchid and d.Compid = a.Compid ) FrmLName"
  Sql = Sql & " from SAL_Jobwork_ItemMaster a Left Outer Join Material_Master b"
  Sql = Sql & " On a.SentItemCode=b.ItemCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
  Sql = Sql & " Left Outer Join Material_Master c"
  Sql = Sql & " On a.ReceivedItemCode=c.ItemCode and a.BranchID=c.BranchID and a.CompID=c.CompID"
  Sql = Sql & " where a.CustomerCode=" & Val(txtVendCode.Tag)
  Sql = Sql & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and b.Status = 1 and c.Status = 1"
  Sql = Sql & " Order by a.ItemGroupKey, a.ReceivedItemCode"
  If RSOpen(snap, Sql) = False Then
    GoTo ExitHere
  End If
  
    If Not snap.EOF Then
      snap.MoveLast
snap.MoveFirst
      sprItem.MaxRows = snap.RecordCount + 1
      ValueRow = 0
PrevItemGroup = 0
      For rCnt = 1 To snap.RecordCount
        If PrevItemGroup <> snap("ItemGroupKey") Then
          PrevItemGroup = snap("ItemGroupKey")
          ValueRow = rCnt
        End If
        sprItem.SetText eItem.mFirstCol, rCnt, ValueRow
        
        sprItem.SetText eItem.mrecdItem, rCnt, snap("ReceivedItemCode")
        sprItem.SetText eItem.mRecdDesc, rCnt, CheckNull(snap("RecdDesc"))
        sprItem.SetText eItem.mRecdUOM, rCnt, CheckNull(snap("RecdUOM"))
        
        sprItem.SetText eItem.mRelation, rCnt, CheckNull(snap("RelationType"))
        sprItem.SetText eItem.mItemGroupKey, rCnt, CheckNull(snap("ItemGroupKey"))
        
        If snap("RelationType") = "1-1" Or (snap("RelationType") = "N-1" And ValueRow = rCnt) Then
          sprItem.SetText eItem.msentItem, rCnt, snap("SentItemCode")
          sprItem.SetText eItem.mSentDesc, rCnt, CheckNull(snap("SentDesc"))
          sprItem.SetText eItem.mSentUOM, rCnt, CheckNull(snap("SentUOM"))
          sprItem.SetText eItem.mPartNo, rCnt, CheckNull(snap("PartNo"))
          
          Call Spread_Set_Decimal(sprItem, CheckNull(snap("NoOfDec"), 0), eItem.mRatePer, rCnt)
          sprItem.SetText eItem.mRatePer, rCnt, CheckNull(snap("RatePer"))
          sprItem.SetText eItem.mRate, rCnt, CheckNull(snap("Rate"))
          sprItem.SetText eItem.mValidFrom, rCnt, IIf(CheckNull(snap("ValidFrom")) = "", "", Format(snap("ValidFrom"), "dd/MM/yyyy"))
          sprItem.SetText eItem.mValidTo, rCnt, IIf(CheckNull(snap("ValidTo")) = "", "", Format(snap("ValidTo"), "dd/MM/yyyy"))
          sprItem.SetText eItem.mQCCheck, rCnt, IIf(snap("QcCheck"), "Yes", "No")
          
          Spread_Lock sprItem, False, eItem.mPartNo, rCnt, eItem.mStatus, rCnt
          Spread_Lock sprItem, True, eItem.mSentUOM, rCnt, eItem.mSentUOM, rCnt
          
          If Val(txtDefaultFormula.Tag) <> snap("FormulaCode") Then
            If IsNull(snap("FrmLName")) Then
              sShowMessage "Formula Not found in the Master for the Item in Row " & rCnt & ". So default formula applied."
            Else
              sprItem.SetText eItem.mFrmCode, rCnt, CheckNull(snap("FormulaCode"))
              sprItem.SetText eItem.mFrmName, rCnt, CheckNull(snap("FrmLName"))
            End If
          End If
          sprItem.SetText eItem.mStatus, rCnt, IIf(snap("IsActive"), "Active", "Inactive")
        Else
          Call Spread_Lock(sprItem, True, eItem.mPartNo, rCnt, eItem.mStatus, rCnt)
        End If
        
        
        snap.MoveNext
      Next rCnt
      sprItem.Tag = snap.RecordCount
    Else
      sprItem.MaxRows = 1
      sprItem.Tag = 0
    End If
  snap.Close
  fraVendor.Enabled = False
  sprItem.Enabled = True
  cmdSave.Enabled = True
  sprItem.SetFocus
  sprItem.SetActiveCell 1, 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = 0
End Try
End Sub

Private Sub cmdSave_Click()
  Dim retVal As Variant
  Dim Sql As String
  Dim rCnt As Long
  Dim MaxItemGroup As Integer
  Dim sFirstRow As Long
  Dim Relation As String
  Dim sChildCode As String
  Dim sRow As Long
  Dim ChkSave As Integer
  Dim CurDate As Date

  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  CurDate = Format(ServerDate, "yyyy-MM-dd")
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
    For rCnt = 1 To sprItem.DataRowCnt
      sprItem.GetText eItem.mFirstCol, rCnt, retVal
       sFirstRow = Val(retVal)    'Row Where the values found for 1 to N, N to 1
      If rCnt = sFirstRow Then
        sRow = sprItem.SearchCol(eItem.mFirstCol, 0, sprItem.MaxRows, Trim$(sFirstRow), SearchFlagsValue)
        sChildCode = ""
        Do While sRow <> -1
          sprItem.GetText eItem.mrecdItem, sRow, retVal
          sChildCode = sChildCode + Trim$(retVal)
          sRow = sprItem.SearchCol(eItem.mFirstCol, sRow, sprItem.MaxRows, Trim$(sFirstRow), SearchFlagsValue)
        Loop
        sprItem.GetText eItem.msentItem, sFirstRow, retVal
        MaxItemGroup = Val(ItemGroupKey(sChildCode, Trim$(retVal)))
        ChkSave = Right(MaxItemGroup, 1)
        MaxItemGroup = Left(MaxItemGroup, Len(Trim$(MaxItemGroup)) - 1)
      End If
      
      If ChkSave = 1 Then 'New
        
        sprItem.GetText eItem.mrecdItem, rCnt, retVal
        Sql = "Insert into ItemGroupKey_Dtl select " & Val(MaxItemGroup) & ",'" & Trim$(retVal) & "',"
        sprItem.GetText eItem.msentItem, sFirstRow, retVal
        Sql = Sql & " '" & Trim$(retVal) & "',"
        sprItem.GetText eItem.mRelation, sFirstRow, retVal
        Sql = Sql & " '" & Trim$(retVal) & "',"
        Sql = Sql & " 1,'" & gUserID & "','" & gBranchID & "','" & gCompID & "','" & Format(CurDate, "yyyy-MM-dd") & "'"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      End If
      If rCnt > Val(sprItem.Tag) Then
        Sql = "Execute SAL_INS_JOBWORK_ITEMMASTER"
         Sql = Sql & " @ItemGroupKey=" & Val(MaxItemGroup) & ","
      Else
        sprItem.GetText eItem.mItemGroupKey, sFirstRow, retVal
        Sql = "Execute SAL_UP_JOBWORK_ITEMMASTER"
        Sql = Sql & " @ItemGroupKey=" & Val(retVal) & ","
      End If
      Sql = Sql & " @CustomerCode=" & txtVendCode.Tag & ","
      sprItem.GetText eItem.mRelation, rCnt, retVal
      Relation = Trim$(retVal)
      If Relation = "N-1" Then
        sprItem.GetText eItem.mrecdItem, rCnt, retVal
      Else
        sprItem.GetText eItem.mrecdItem, sFirstRow, retVal
      End If
      Sql = Sql & " @ReceivedItemCode='" & Trim$(retVal) & "',"
      
      If Relation = "1-N" Then
        sprItem.GetText eItem.msentItem, rCnt, retVal
      Else
        sprItem.GetText eItem.msentItem, sFirstRow, retVal
      End If
      Sql = Sql & " @SentItemCode='" & Trim$(retVal) & "',"
      
      Sql = Sql & " @RelationType='" & Relation & "',"
      
      sprItem.GetText eItem.mPartNo, sFirstRow, retVal
      Sql = Sql & " @PartNo='" + Trim$(retVal) + "',"
      
      sprItem.GetText eItem.mRatePer, sFirstRow, retVal
      Sql = Sql & " @RatePer=" & IIf(Val(retVal) = 0, 1, Val(retVal)) & ","
      sprItem.GetText eItem.mRate, sFirstRow, retVal
      Sql = Sql & " @Rate=" & Val(retVal) & ","
      sprItem.GetText eItem.mAssRate, sFirstRow, retVal
      Sql = Sql & " @AssessableRate=" & Val(retVal) & ","
      sprItem.GetText eItem.mValidFrom, sFirstRow, retVal
      Sql = Sql & " @ValidFrom='" & Format(retVal, "yyyy-MM-dd") & "',"
      sprItem.GetText eItem.mValidTo, sFirstRow, retVal
      Sql = Sql & " @ValidTo='" & Format(retVal, "yyyy-MM-dd") & "',"
      sprItem.GetText eItem.mQCCheck, sFirstRow, retVal
      Sql = Sql & " @QcCheck='" & IIf(Trim$(retVal) = "Yes", "1", "0") & "',"
      
      sprItem.GetText eItem.mFrmCode, sFirstRow, retVal
      If Val(retVal) <= 0 Then
        'Default Formula
        Sql = Sql & " @FormulaCode=" & Val(txtDefaultFormula.Tag) & ","
      Else
        'Specified Formula
        Sql = Sql & " @FormulaCode=" & Val(retVal) & ","
      End If
      
      sprItem.GetText eItem.mStatus, sFirstRow, retVal
      Sql = Sql & " @isActive='" & IIf(Trim$(retVal) = "Active", "1", "0") & "',"
      Sql = Sql & " @UserID='" & gUserID & "',"
      Sql = Sql & " @BranchID='" & gBranchID & "',"
      Sql = Sql & " @CompID='" & gCompID & "'"
      If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      
    Next rCnt
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Jobwork Item Details Saved"
  Call cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub Form_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  If TypeOf Me.ActiveControl Is fpSpread Then
    GoTo ExitHere
  Else
'    If KeyAscii = vbKeyReturn Then
'      SendKeys "{tab}"
'    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub EXP_REC_Indented(ByVal parent As String)
  'Indented explosion using Recursion - Ideal display order
  ReDim ChildArray(1 To 10000) As String
  Dim Curr_Child As String
  Dim Child_desc  As String
  Dim Curr_QTY As Double
  Dim nrec As Long
  Dim count As Integer
  Dim sChild As AceADODB.Recordset
  Dim LevelText As String
  Dim Sql As String
   
  Try
  
  fLevel_Ind = fLevel_Ind + 1
  Sql = " select a.ChildCode, a.QTY, b.ItemDescription, b.UOMDesc from BOM_STRUCTURE_DTL a, "
  Sql = Sql & " MATERIAL_MASTER b where a.ParentCode = '" & Trim(parent) & "' and a.Branchid = '" & gBranchID & "' and b.Status = 1"
  Sql = Sql & " and a.compid = '" & gCompID & "' and a.branchid=b.Branchid and a.compid=b.Compid"
  Sql = Sql & " and a.ChildCode=b.ItemCode and a.RevisionNo = (select max(RevisionNo) "
  Sql = Sql & " from BOM_STRUCTURE_HDR d where d.ParentCode = '" & (parent) & "' and Branchid = '" & gBranchID & "'"
  Sql = Sql & " and d.compid = '" & gCompID & "')"
  Set sChild = New AceADODB.Recordset
  If RSOpen(sChild, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not sChild.EOF Then
    sChild.MoveLast
sChild.MoveFirst
    nrec = sChild.RecordCount
  Else
    nrec = 0
    If fLevel_Ind = 1 Then
      MsgBox "No Child Items for Item " & parent, vbInformation + vbOKOnly, "BOM"
    End If
  End If
  Do While Not sChild.EOF
    Curr_Child = sChild(0)
    Curr_QTY = sChild(1)
    Child_desc = sChild(2)
    count = count + 1
    ChildArray(count) = Curr_Child
    LevelText = Space(3 * fLevel_Ind) & Format(fLevel_Ind)
    If fLevel_Ind > fPrev_Level_Ind And fLevel_Ind > fhighest_Level Then
      fhighest_Level = fLevel_Ind
    End If
    
    Sql = "insert into SAL_ReceivedItem_tmp values('" + Curr_Child + "', '" + Child_desc + "', '" + sChild(3) + "', '" & gUserID & "', '" & gBranchID & "','" & gCompID & "')"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
    'SendToPort LAlign(7, Trim(LevelText)) & Space(1) & LAlign(20, Curr_Child) & Space(1) & LAlign(30, Child_desc) & Space(1) & LAlign(5, sChild(3)) & Space(1) & RAlign(10, CStr(Curr_QTY))
    fPrev_Level_Ind = fLevel_Ind
    frowno = frowno + 1
    reccount = reccount + 1
    'Recursive Call
    Call EXP_REC_Indented(ChildArray(count))
    
    sChild.MoveNext
  Loop
  fLevel_Ind = fLevel_Ind - 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set sChild = Nothing
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub

  Try

'  Call Spread_Set_HeaderColor(sprItem, anSelectOnly, anCol, 1)  'ItemCode
'  Call Spread_Set_HeaderColor(sprItem, anSelectOnly, anCol, 11) 'Formula Name
   
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub sprItem_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  Dim tmp As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If Col = eItem.mrecdItem Then    'Select Sent ItemCode means One To One / One To Many
    Call OneToOneItemHelp(Row)
  
  ElseIf Col = eItem.msentItem Then 'Select Recd ItemCode means Many to One
    Call ManyToOneItemHelp(Row)
    
  ElseIf Col = eItem.mFrmName Then   'Formula Name
    sprItem.GetText eItem.mFirstCol, Row, tmp
    If Val(tmp) = Row Then
      Sql = "Select distinct a.FrmId, a.FrmLName from COM_FormulaMast_Hdr a where "
      Sql = Sql + " a.ModuleCode = '" & gModuleCode & "' and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
      If FetchValue_New(False, db.ConnectionString, Sql, "11", "List of Formula", "Formula Code", "Formula Name") = True Then
        sprItem.SetText eItem.mFrmCode, Row, getvalue(1)
        sprItem.SetText eItem.mFrmName, Row, getvalue(2)
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

    If fRow > Val(sprItem.Tag) Then    'already saved itemcode should not be changed
      sprItem.GetText eItem.mFirstCol, fRow, retVal
      If Len(Trim$(retVal)) = 0 Then
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
          Sql = "Select a1.StageCode, (Select b.OperationName from BOM_Operation_Master b where a1.OperationCode=b.OperationCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID) as StageDesc,"
          Sql = Sql & " (Select b.UOMDesc from Material_Master b where a1.StageCode=b.ItemCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID and b.Status = 1 ) as UOM, "
          Sql = Sql & " a1.ItemCode, (Select b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ItemCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ItemDesc, 0 as Priority"
          Sql = Sql & " from BOM_MPS_Master a1"
          Sql = Sql & " where a1.ItemCode='" & sParentCode & "' and a1.OperationSequence > 0"
          Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "'"
          Sql = Sql & " Union All"
          Sql = Sql & " Select a1.ChildCode, b1.ItemDescription as ChildDesc, b1.PurchaseUOMDesc as UOM,"
          Sql = Sql & " a1.ParentCode, (Select b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ParentCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ParentDesc, a1.Priority"
          Sql = Sql & " from BOM_Structure_Dtl a1 Left Outer Join Material_Master b1 On b1.ItemCode=a1.ChildCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID"
          Sql = Sql & " where a1.ParentCode=(Select a2.ParentCode from BOM_Structure_Dtl a2"
          Sql = Sql & " Where a2.ParentCode=isnull((Select a3.StageCode from BOM_MPS_Master a3 Where a3.ItemCode='" & sParentCode & "' and a3.OperationSequence=1"
          Sql = Sql & " and a3.BranchID=a2.BranchID and a3.CompID=a2.CompID and Priority=1),(Select a3.StageCode from BOM_MPS_Master a3 Where a3.ItemCode='" & sParentCode & "' and a3.OperationSequence=0"
          Sql = Sql & " and a3.BranchID=a2.BranchID and a3.CompID=a2.CompID and Priority=1))"
          Sql = Sql & " and a2.Explodable = 1 and a2.BranchID='" & gBranchID & "' and a2.CompID='" & gCompID & "'"
          Sql = Sql & " Group by a2.ParentCode Having Count(*) = 1)"
          Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "' and b1.Status = 1 "
          Sql = Sql & " Order by priority"
          If FetchValue_New(False, db.ConnectionString, Sql, "11011", "List of Process - Select Received Item", "Stage Code", "Description", "Item Code", "Description") = True Then
            sprItem.SetText eItem.mrecdItem, fRow, getvalue(1)
            sprItem.SetText eItem.mRecdDesc, fRow, getvalue(2)
            sprItem.SetText eItem.mRecdUOM, fRow, getvalue(3)
            
            'Show the Process Details for that Item greater than the selected process
            Sql = "Select a1.StageCode, (Select b.OperationName from BOM_Operation_Master b where a1.OperationCode=b.OperationCode and a1.BranchID=b.BranchID and a1.CompID=b.CompID) as StageDesc,"
            Sql = Sql & " a3.PurchaseUOMDesc as UOM, a3.PurchaseUOMDecimal as NoOfDec,"
            Sql = Sql & " a1.ItemCode, (Select b1.ItemDescription from Inv_Material_Master b1 where b1.ItemCode=a1.ItemCode and b1.BranchID=a1.BranchID and b1.CompID=a1.CompID and b1.status = 1) as ItemDesc"
            Sql = Sql & " from BOM_MPS_Master a1 Left Outer Join Material_Master a3 On a1.StageCode=a3.ItemCode and a1.BranchID=a3.BranchID and a1.CompID=a3.CompID"
            Sql = Sql & " Where a1.ItemCode='" & sParentCode & "' and a3.Status = 1 and (a1.OperationSequence = 0 Or a1.OperationSequence > isNull((select OperationSequence from bom_mps_master a2 where stagecode='" & getvalue(1) & "'"
            Sql = Sql & " and a2.BranchID='" & gBranchID & "' and a2.CompID='" & gCompID & "'), -1))"
            Sql = Sql & " and a1.BranchID='" & gBranchID & "' and a1.CompID='" & gCompID & "'"
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
                sprItem.SetText eItem.mStatus, fRow, "Active"
              
                sprItem.MaxRows = sprItem.DataRowCnt + 1
                Call Spread_Lock(sprItem, False, eItem.mPartNo, fRow, sprItem.MaxCols, fRow)
                Call Spread_Lock(sprItem, True, eItem.mFrmName, fRow, eItem.mFrmName, fRow)
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

Private Sub sprItem_KeyUp(sender as object, e as KeyUpEventArgs)

  Try
  Me.MousePointer = vbHourglass
  
    Select Case KeyCode
      Case vbKeyF2
        Call sprItem_DblClick(sprItem.ActiveCol, sprItem.ActiveRow)
      Case vbKeyDelete
        If sprItem.ActiveCol = eItem.mFrmName Then
          sprItem.ClearRange eItem.mFrmCode, sprItem.ActiveRow, eItem.mFrmName, sprItem.ActiveRow, True
        Else
          If sprItem.ActiveRow > Val(sprItem.Tag) And sprItem.ActiveCol = eItem.mrecdItem Then
            sprItem.DeleteRows sprItem.ActiveRow, 1
            sprItem.MaxRows = sprItem.DataRowCnt + 1
          End If
        End If
    End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtDefaultFormula_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = "Select distinct a.FrmId, a.FrmLName from COM_FormulaMast_Hdr a where "
  Sql = Sql + " a.ModuleCode = '" & gModuleCode & "' and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, Sql, "11", "List of Formula", "Formula Code", "Formula Name") = True Then
    txtDefaultFormula.Text = getvalue(2)
    txtDefaultFormula.Tag = getvalue(1)
  Else
    If Val(txtDefaultFormula.Tag) <= 0 Then
      sShowMessage "Default Formula must be selected."
      txtDefaultFormula.SetFocus
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtDefaultFormula_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
    Call txtDefaultFormula_DblClick
  End If
End Sub

Private Sub txtVendCode_DblClick()
  Dim Sql As String

  Try
  Me.MousePointer = vbHourglass

  Sql = "select PartyName, PartyCode from FAS_PARTY_MASTER where AcType IN ('D','Y') "
  Sql = Sql & " and BranchID='" & gBranchID & "' and compid='" & gCompID & "' and EffStatus = 1"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Vendors", "Vendor Name") = True Then
    txtVendCode.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
    txtVendCode = IIf(IsNull(getvalue(1)), "", getvalue(1))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub txtVendCode_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)

  Try
  Me.MousePointer = vbHourglass

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Function ValidateData() As Boolean
  Dim retVal As Variant
  Dim chkval As Variant
  Dim rCnt As Long
  Dim CurDate As Date
  
  Try
  ValidateData = False
  
  If Trim(txtDefaultFormula) = "" Then
    sShowMessage "Select the Formula"
    txtDefaultFormula.SetFocus
    ValidateData = False
Exit Function
  End If

  CurDate = ServerDate

    For rCnt = 1 To sprItem.DataRowCnt
      sprItem.GetText eItem.mFirstCol, rCnt, retVal
      If Val(retVal) = rCnt Then
        sprItem.GetText eItem.mRatePer, rCnt, retVal        'Rate Per
        If Val(retVal) = 0 Then
          sShowMessage "Enter the Rate Per Qty"
          sprItem.SetActiveCell eItem.mRatePer, rCnt
          sprItem.SetFocus
          GoTo ExitHere
        End If
        
        sprItem.GetText eItem.mRate, rCnt, retVal        'Rate
        If Val(retVal) = 0 Then
          sShowMessage "Enter the Rate"
          sprItem.SetActiveCell eItem.mRate, rCnt
          sprItem.SetFocus
          GoTo ExitHere
        End If
        
        sprItem.GetText eItem.mValidFrom, rCnt, retVal      'Valid From
        If Not IsDate(retVal) Then
          sShowMessage "Enter ValidFrom Date"
          sprItem.SetActiveCell eItem.mValidFrom, rCnt
          sprItem.SetFocus
          GoTo ExitHere
        End If
        
        sprItem.GetText eItem.mValidTo, rCnt, chkval      'Valid To
        If Not IsDate(chkval) Then
          sShowMessage "Enter ValidTo Date"
          sprItem.SetActiveCell eItem.mValidTo, rCnt
          sprItem.SetFocus
          GoTo ExitHere
        End If
         
         sprItem.GetText eItem.mQCCheck, rCnt, chkval      'Qc Check
        If Trim$(chkval) = "" Then
          sShowMessage "Select QC Check"
          sprItem.SetActiveCell eItem.mQCCheck, rCnt
          sprItem.SetFocus
          GoTo ExitHere
        End If
        
        If rCnt > Val(sprItem.Tag) Then
          If IsDate(retVal) And IsDate(chkval) Then
            If DateValue(retVal) > DateValue(chkval) Then
               sShowMessage "ValidTo Date should be greater than ValidFrom Date"
               sprItem.SetActiveCell eItem.mValidTo, rCnt
               sprItem.SetFocus
               GoTo ExitHere
            ElseIf Format(chkval, "yyyyMMdd") < Format(CurDate, "yyyyMMdd") Then
               sShowMessage "ValidTo Date should be greater than Current Date"
               sprItem.SetActiveCell eItem.mValidTo, rCnt
               sprItem.SetFocus
               GoTo ExitHere
            End If
          End If
        End If
      End If
    Next rCnt
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_sprItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprItem


sprItem.DblClick +=  new EventHandler(sprItem_DblClick)
sprItem.KeyUp +=  new EventHandler(sprItem_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\CustomerJobworkMaster.frm", sprItem, "")

End Sub
