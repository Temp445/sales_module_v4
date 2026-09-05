VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSalesDeliveryChallan 
   Caption         =   "Delivery Challan"
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
Begin AceSpread spItem
      Height          =   1905
      Left            =   150
      TabIndex        =   7
      ToolTipText     =   "Select the Item Code / Jobwork Item. Press F2 or Double Click for Help."
      Top             =   930
      Width           =   11565
      _Version        =   458752
      _ExtentX        =   20399
      _ExtentY        =   3360
      _StockProps     =   64
      Enabled         =   0   'False
      ColsFrozen      =   2
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   17
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   480
      Left            =   150
      MaxLength       =   150
      MultiLine       =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Enter the Remarks"
      Top             =   3045
      Width           =   11565
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   400
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   3615
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5282
      TabIndex        =   10
      ToolTipText     =   "Click to Clear the Data"
      Top             =   3615
      Width           =   1300
   End
   Begin VB.Frame fraCustomer 
      Height          =   720
      Left            =   4785
      TabIndex        =   16
      Top             =   135
      Width           =   5385
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   4695
         TabIndex        =   6
         ToolTipText     =   "Click to OK"
         Top             =   300
         Width           =   525
      End
      Begin VB.TextBox txtcustcd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   75
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         ToolTipText     =   "Select the Customer. Press F2 or Double Click for Help."
         Top             =   315
         Width           =   4560
      End
      Begin VB.Label lblcustcode 
         AutoSize        =   -1  'True
         Caption         =   "* Customer"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   75
         TabIndex        =   17
         Top             =   135
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3990
      TabIndex        =   9
      ToolTipText     =   "Click to Save the Data"
      Top             =   3615
      Width           =   1300
   End
   Begin VB.Frame fraDC 
      Height          =   720
      Left            =   150
      TabIndex        =   13
      Top             =   135
      Width           =   2940
      Begin VB.ComboBox cmbDCType 
         Height          =   330
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select the DC Type"
         Top             =   315
         Width           =   765
      End
      Begin VB.TextBox txtDCNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   795
      End
      Begin MSComCtl2.DTPicker dtpDCDate 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         ToolTipText     =   "DC Date"
         Top             =   315
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   141623297
         CurrentDate     =   37250
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DC No"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1215
         TabIndex        =   19
         Top             =   135
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "DC Type"
         Height          =   210
         Left            =   105
         TabIndex        =   15
         Top             =   135
         Width           =   615
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   1680
         TabIndex        =   14
         Top             =   135
         Width           =   330
      End
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source of DC"
      Height          =   720
      Left            =   3105
      TabIndex        =   12
      ToolTipText     =   "Select the DC Source"
      Top             =   135
      Width           =   1680
      Begin VB.OptionButton optorder 
         Caption         =   "&Order"
         Height          =   210
         Left            =   870
         TabIndex        =   4
         Top             =   323
         Width           =   735
      End
      Begin VB.OptionButton optdirect 
         Caption         =   "&Direct"
         Height          =   210
         Left            =   75
         TabIndex        =   3
         Top             =   323
         Value           =   -1  'True
         Width           =   720
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   150
      TabIndex        =   18
      Top             =   2865
      Width           =   630
   End
End
Attribute VB_Name = "frmSalesDeliveryChallan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Enum eItem
  mItemCode = 1
  mItemDesc = 2
  mWhsCode = 3
  mWhsName = 4
  mRackCode = 5
  mStockQty = 6
  mDCQty = 7
  mRate = 8
  mARate = 9
  mGV = 10
  mAGV = 11
  mOType = 12
  mONo = 13
  mODate = 14
  mBalQty = 15
  mSchType = 16
  mGroupKey = 17
End Enum

Private Function WareHouseBalQty(ByVal Row As Long,ByVal  ItemCode As String) As String
  Dim retVal1 As Variant
  Dim retVal As Variant
  Dim ResCode As String
  Dim sRow As Long
  Dim wRow As Long
  Dim sQty As Double
  Dim WareCode As String
  Dim alrCode(50) As String
  Dim sflag As Boolean
  Dim i As Integer
  Dim j As Integer

  Try
  j = 1
sflag = True
  sRow = spItem.SearchCol(eItem.mItemCode, 0, spItem.MaxRows, Trim$(ItemCode), SearchFlagsValue)
  sQty = 0
WareCode = ""
ResCode = ""
  Do While sRow <> -1
   sQty = 0
    spItem.GetText eItem.mWhsCode, sRow, retVal
    spItem.GetText eItem.mItemCode, sRow, retVal1
    sflag = True
    For i = 1 To j
      If alrCode(i) = Trim$(retVal) Then
       sflag = False
       Exit For
      End If
    Next i
    If sflag = True Then
      wRow = spItem.SearchCol(eItem.mWhsCode, 0, spItem.MaxRows, Trim$(retVal), SearchFlagsValue)
      Do While wRow <> -1
        spItem.GetText eItem.mItemCode, wRow, retVal
        If Trim$(retVal1) = Trim$(retVal) Then
          spItem.GetText eItem.mDCQty, wRow, retVal
          sQty = sQty + Val(retVal)
          spItem.GetText eItem.mWhsCode, wRow, retVal
          alrCode(j) = Trim$(retVal)
          j = j + 1
        End If
        spItem.GetText eItem.mWhsCode, wRow, retVal
        wRow = spItem.SearchCol(eItem.mWhsCode, wRow, spItem.MaxRows, Trim$(retVal), SearchFlagsValue)
      Loop
      spItem.GetText eItem.mWhsCode, sRow, retVal
      WareCode = WareCode + "'" + Trim$(retVal) + "'" + " " + "Then" + " " + CStr(Val(sQty)) + " when "
    End If
    sRow = spItem.SearchCol(eItem.mItemCode, sRow, spItem.MaxRows, Trim$(ItemCode), SearchFlagsValue)
  Loop
  
  If WareCode <> "" Then
    WareCode = "case b.WareHouseCode when " & "" & Left(WareCode, Len(Trim$(WareCode)) - 4) & "Else" & " 0" & " End"
  End If
  WareHouseBalQty = IIf(WareCode = "", 0, WareCode)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Direct()
  Dim Sql As String
  Dim i As Long

  Try

  Sql = Sql + " select a.ItemCode,c.ItemDescription,a.PartNo,c.UOMDecimal,a.Rate,"
  Sql = Sql + " a.AssessableRate,0 as ItemGroupKey from SAL_Customer_ItemMaster a left outer join"
  Sql = Sql + " MATERIAL_MASTER c on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid"
  Sql = Sql + " and a.Compid=C.Compid where a.CustomerCode=" & Val(txtcustcd.Tag) & ""
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "' and c.Status = 1"
  Sql = Sql + " and a.ValidTo >=convert(datetime,'" + Format(ServerDate, "yyyy-MM-dd") + "',111)"
  Sql = Sql + " union (select distinct a.SentItemCode,c.ItemDescription,'',c.UOMDecimal,"
  Sql = Sql + " a.Rate,a.AssessableRate,a.ItemGroupKey from SAL_JOBWORK_ITEMMASTER a left outer join MATERIAL_MASTER c "
  Sql = Sql + " on a.SentItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=C.Compid"
  Sql = Sql + " where a.CustomerCode=" & Val(txtcustcd.Tag) & ""
  Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "' and c.Status = 1"
  Sql = Sql + " and a.ValidTo >=convert(datetime,'" + Format(ServerDate, "yyyy-MM-dd") + "',111))"
  If FetchValue_New(True, db.ConnectionString, Sql, "111000", "List of Item's", "Item Code", "Item Description", "Part No") = True Then
    spItem.MaxRows = UBound(getvalue, 1)
    For i = 1 To UBound(getvalue, 1)
      spItem.SetText 1, i, Trim(getvalue(i, 1))
      spItem.SetText 2, i, getvalue(i, 2)
      Call Spread_Set_Decimal(spItem, Val(getvalue(i, 4)), 6, i, 7)
      Call Spread_Set_Decimal(spItem, Val(getvalue(i, 4)), 15, i)
      spItem.SetText 8, i, getvalue(i, 5) 'Rate
      spItem.SetText 9, i, getvalue(i, 6) 'Assessable Rate
      spItem.SetText eItem.mGroupKey, i, getvalue(i, 7) 'Item GroupKey
    Next i
    fraDC.Enabled = False
    fraSource.Enabled = False
    fraCustomer.Enabled = False
    spItem.Enabled = True
    txtRemarks.Enabled = True
    cmdSave.Enabled = True
    spItem.SetFocus
    Spread_Set_Focus spItem, 1, 4
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function ValidateData() As Boolean
  Dim i As Long
  Dim tmp As Variant
  
  Try
  ValidateData = False
  
   For i = 1 To spItem.DataRowCnt
    spItem.GetText eItem.mWhsName, i, tmp
    If Trim$(tmp) = "" Then
      sShowMessage "Please select the Warehouse"
      spItem.SetFocus
      Spread_Set_Focus spItem, i, eItem.mWhsName
      GoTo ExitHere
    End If
    
    spItem.GetText eItem.mDCQty, i, tmp
    If Val(tmp) = 0 Then
      sShowMessage "Please enter the DC Qty"
      spItem.SetFocus
      Spread_Set_Focus spItem, i, eItem.mDCQty
      GoTo ExitHere
    End If
  Next i
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

'Private Sub LoadDCType()
'Dim InvTypeRec As ADODB.Recordset
'Dim Sql As String
'
'  On Error GoTo ErrHandler
'
'  Set InvTypeRec = New ADODB.Recordset
'  Sql = "select DocumentType from COM_DOCUMENTTYPE_MASTER where PertainingTo='D'"
'  Sql = Sql + " and ModuleName='SAL' and Branchid ='" & gBranchID & "' and Compid = '" & gCompID & "'"
'  If RSOpen(InvTypeRec, Sql) = False Then GoTo ExitHere
'  If Not InvTypeRec.EOF Then
'     Do Until InvTypeRec.EOF
'        cmbDCType.AddItem InvTypeRec("DocumentType")
'      InvTypeRec.MoveNext
'     Loop
'     cmbDCType.ListIndex = 0
'     txtDCNo = GetNo("SAL_DC_HDR", "DCType", "DCNo", Trim$(cmbDCType), "D", "SAL")
'  End If
'  InvTypeRec.Close
'ExitHere:
'  On Error GoTo 0
'  Set InvTypeRec = Nothing
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'  End Sub

Private Sub cmbdctype_Click()
  Try
  Me.MousePointer = vbHourglass
  
  Call SetDocNo
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub SetDocNo()
  Try

  txtDCNo = GetDocNo("SAL_DC_HDR", "DCType", "DCNo", Trim$(cmbDCType), "D")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub cmdClear_Click()
  Try
  
  Me.MousePointer = vbHourglass

  Form_Clear Me
  Call LoadDocType(mSalDC)
  fraDC.Enabled = True
  fraSource.Enabled = True
  fraCustomer.Enabled = True
  spItem.Enabled = False
  txtRemarks.Enabled = False
  cmdSave.Enabled = False
  optdirect.Value = True
'  LoadDCType 'Loading Invoice types in the combo
  cmbDCType.SetFocus
  CurrDate Me
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

Private Sub CmdOk_Click()
  Try
  Me.MousePointer = vbHourglass

  If Trim$(txtcustcd) = "" Then
    sShowMessage "Please enter the Customer"
    txtcustcd.SetFocus
    GoTo ExitHere
  End If
  
  If optorder Then
 Orders 'list the Orders
End If
  If optdirect Then
 Direct 'list the Direct
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim snap1 As AceADODB.Recordset
  Dim i As Long
  Dim tmp As Variant
  Dim Stime As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  Call SetDocNo
  
  Set cmd = New AceADODB.Command
  Set snap1 = New AceADODB.Recordset
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  commClear cmd
  
  cmd.CommandText = "SAL_INS_DC_HDR"
  cmd.Parameters.Append cmd.CreateParameter("DCType", adVarChar, adParamInput, 3, Trim$(cmbDCType))
  cmd.Parameters.Append cmd.CreateParameter("DCNo", adInteger, adParamInput, , Val(Trim$(txtDCNo)))
  cmd.Parameters.Append cmd.CreateParameter("DCDate", adDBTimeStamp, adParamInput, , Format(dtpDCDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("DCSource", adChar, adParamInput, 1, IIf(optdirect, "D", "O"))
  cmd.Parameters.Append cmd.CreateParameter("CustomerCode", adInteger, adParamInput, , Val(Trim$(txtcustcd.Tag)))
  cmd.Parameters.Append cmd.CreateParameter("Remarks", adVarChar, adParamInput, 150, Trim$(txtRemarks))
  Call AddCommonCmdParameters(cmd)
  cmd.Execute

    For i = 1 To spItem.DataRowCnt
      commClear cmd
      cmd.CommandText = "SAL_INS_DC_DTL"
      cmd.Parameters.Append cmd.CreateParameter("DCType", adVarChar, adParamInput, 3, Trim$(cmbDCType))
      cmd.Parameters.Append cmd.CreateParameter("DCNo", adInteger, adParamInput, , Val(Trim$(txtDCNo)))
      cmd.Parameters.Append cmd.CreateParameter("DCDate", adDBTimeStamp, adParamInput, , Format(dtpDCDate, "yyyy-MM-dd"))
      spItem.GetText eItem.mGroupKey, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
      spItem.GetText eItem.mItemCode, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemCode", adChar, adParamInput, 20, Trim$(tmp))
      spItem.GetText eItem.mDCQty, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("Qty", adDouble, adParamInput, , Val(tmp))
      spItem.GetText eItem.mRate, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("Rate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
      spItem.GetText eItem.mARate, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("AssessableRate", adDouble, adParamInput, , Format(Val(tmp), "0.00"))
      spItem.GetText eItem.mOType, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("OrderType", adVarChar, adParamInput, 3, Trim$(tmp))
      spItem.GetText eItem.mONo, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("OrderNo", adInteger, adParamInput, , Val(tmp))
      spItem.GetText eItem.mODate, i, tmp
      If Trim$(tmp) <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Format(Trim$(tmp), "yyyy-MM-dd"))
      Else
        cmd.Parameters.Append cmd.CreateParameter("OrderDate", adDBTimeStamp, adParamInput, , Null)
      End If
      Call AddCommonCmdParameters(cmd)
      cmd.Execute

      Stime = Format(ServerDate, "HH:mm:ss")
      
      commClear cmd
      cmd.CommandText = "INV_INS_MATERIALTRAN_DTL"
      cmd.Parameters.Append cmd.CreateParameter("DOCType", adVarChar, adParamInput, 3, Trim$(cmbDCType))
      cmd.Parameters.Append cmd.CreateParameter("DOCNo", adVarChar, adParamInput, 15, Trim$(txtDCNo))
      cmd.Parameters.Append cmd.CreateParameter("DOCDate", adDBTimeStamp, adParamInput, , Format(dtpDCDate, "yyyy-mm-dd"))
      cmd.Parameters.Append cmd.CreateParameter("TransactionDate", adDBTimeStamp, adParamInput, , Format(ServerDate, "yyyy-MM-dd"))
      cmd.Parameters.Append cmd.CreateParameter("TransactionTime", adDBTimeStamp, adParamInput, , Stime)
      spItem.GetText eItem.mItemCode, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("TransactionItem", adVarChar, adParamInput, 20, Trim$(tmp))
      spItem.GetText eItem.mDCQty, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("TransactionQty", adDouble, adParamInput, , Val(tmp))
      cmd.Parameters.Append cmd.CreateParameter("Norms", adDouble, adParamInput, , 1)
      spItem.GetText eItem.mWhsCode, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("FromWarehouseCode", adVarChar, adParamInput, 6, Trim$(tmp))
      spItem.GetText eItem.mRackCode, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("FromRackBinCode", adVarChar, adParamInput, 6, Trim$(tmp))
      cmd.Parameters.Append cmd.CreateParameter("ToWarehouseCode", adVarChar, adParamInput, 6, Trim(txtcustcd.Tag))
      cmd.Parameters.Append cmd.CreateParameter("ToRackBinCode", adVarChar, adParamInput, 6, "")
      spItem.GetText eItem.mGroupKey, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ItemGroupKey", adInteger, adParamInput, , Val(tmp))
      cmd.Parameters.Append cmd.CreateParameter("FromItem", adVarChar, adParamInput, 20, "")
      spItem.GetText eItem.mItemCode, i, tmp
      cmd.Parameters.Append cmd.CreateParameter("ToItem", adVarChar, adParamInput, 20, Trim$(tmp))
      cmd.Parameters.Append cmd.CreateParameter("Justification", adVarChar, adParamInput, 4, "")
      Call AddCommonCmdParameters(cmd)
      cmd.Execute
    Next i
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  sShowMessage "Record(s) Saved"
  cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
  Set snap1 = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub dtpdcdate_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
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

  CurrDate Me
'  LoadDCType 'Loading Invoice types in the combo
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

Private Sub spItem_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  Dim tmp As Variant
  Dim retVal As Variant
  Dim fCurDate As Date
  
  Try
  Me.MousePointer = vbHourglass
  
  'Get ServerDate
  fCurDate = ServerDate
  
  If Row <> 1 Then
    spItem.GetText eItem.mDCQty, IIf(Row > 1, Row - 1, Row), retVal
  Else
    retVal = 1
  End If
  spItem.GetText eItem.mStockQty, Row, tmp
   
  If Col = eItem.mWhsName And Row > 0 And Val(retVal) > 0 And Val(tmp) = 0 Then

    spItem.GetText 1, Row, tmp
    
    spItem.GetText eItem.mItemCode, Row, retVal
    retVal = WareHouseBalQty(Row, CStr(retVal))
    
    Sql = " Select b.WarehouseName,a.WarehouseCode, a.RackBinCode, (a.CBQty-" & Trim$(retVal) & ") CBQty"
    Sql = Sql & " From dbo.fn_INV_AsOnDate_Stock_V2('" & Format(fCurDate, "yyyy-MM-dd HH:mm:ss") & "',  '" & Format(fCurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & Format(fCurDate, "yyyy-MM-dd HH:mm:ss") & "', '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a,"
    'Sql = Sql + " From INV_STOCK_MASTER a, "
    Sql = Sql + " INV_WAREHOUSE_MASTER b "
    Sql = Sql + " Where b.WarehouseStatus = 1 "
    Sql = Sql + " and a.ItemCode = '" + Trim$(tmp) + "' and (a.CBQty-" & Trim$(retVal) & ") >0"
    Sql = Sql + " and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
    Sql = Sql + " and a.WarehouseCode=b.WarehouseCode and a.Branchid=b.Branchid "
    Sql = Sql + " and a.Compid=b.Compid"
    If FetchValue_New(False, db.ConnectionString, Sql, "1011", "List of Stock Details", "Warehouse Name", "Rack & Bin", "Stock Qty") = True Then
      spItem.SetText eItem.mWhsCode, Row, getvalue(2)
      spItem.SetText eItem.mWhsName, Row, getvalue(1)
      spItem.SetText eItem.mRackCode, Row, getvalue(3)
      spItem.SetText eItem.mStockQty, Row, getvalue(4)
      Call Spread_Lock(spItem, False, eItem.mDCQty, Row)
      spItem.GetText 16, Row, tmp
      If Trim$(tmp) = "S" And optorder Then
        spItem.GetText 15, Row, tmp
'        Spread_Max_Set spItem, eItem.mDCQty, eItem.mDCQty, Row, IIf(Val(getvalue(4)) > Val(tmp), Val(tmp), Val(getvalue(4)))
        Call Spread_Set_Max(spItem, IIf(Val(getvalue(4)) > Val(tmp), Val(tmp), Val(getvalue(4))), eItem.mDCQty, Row)
      Else
'        Spread_Max_Set spItem, eItem.mDCQty, eItem.mDCQty, Row, Val(getvalue(4))
        Call Spread_Set_Max(spItem, Val(getvalue(4)), eItem.mDCQty, Row)
      End If
      Spread_Set_Focus spItem, Row, eItem.mDCQty
    End If
      Spread_Lock spItem, True, eItem.mDCQty, 0, eItem.mDCQty, Row - 1
   End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spItem_EditMode(sender as object, e as EditModeEventArgs)
  Dim tmp As Variant
  Dim tmp1 As Variant
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 1 Then
 GoTo ExitHere
End If
  
  spItem.GetText eItem.mDCQty, Row, tmp  'Qty
  spItem.GetText eItem.mRate, Row, tmp1 'Rate
  spItem.SetText eItem.mGV, Row, Format(Val(tmp) * Val(tmp1), "0.00") 'GV
  spItem.GetText eItem.mARate, Row, tmp1 'ARate
  spItem.SetText eItem.mAGV, Row, Format(Val(tmp) * Val(tmp1), "0.00") 'AGV
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spItem_KeyDown(sender as object, e as KeyDownEventArgs)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 spItem_DblClick spItem.ActiveCol, spItem.ActiveRow
End If
End Sub

Private Sub txtcustcd_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = " Select distinct (Select b.PartyName From FAS_PARTY_MASTER b Where b.PartyCode=a.CustomerCode "
  Sql = Sql & " and b.Branchid=a.Branchid and b.Compid=a.Compid) PartyName,a.CustomerCode From SAL_Customer_ItemMaster a"
  Sql = Sql & " Where a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
  Sql = Sql & " Union (select distinct (Select b1.PartyName From FAS_PARTY_MASTER b1 Where b1.PartyCode=a1.CustomerCode "
  Sql = Sql & " and b1.Branchid=a1.Branchid and b1.Compid=a1.Compid) PartyName,a1.CustomerCode From SAL_Jobwork_ItemMaster a1"
  Sql = Sql & " Where a1.Branchid='" & gBranchID & "' and a1.Compid='" & gCompID & "')"
  If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Customers", "Customer") = True Then
    txtcustcd.Tag = IIf(IsNull(getvalue(2)), "", getvalue(2))
    txtcustcd = IIf(IsNull(getvalue(1)), "", getvalue(1))
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtcustcd_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
     txtcustcd_DblClick
  End If
End Sub

Private Sub Orders()
  Dim Sql As String
  Dim i As Long
  Dim CurDate As Date

  Try
  
  CurDate = ServerDate

  Sql = "Select a.OrderType,a.OrderNo,a.OrderDate from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where a.CustomerCode=" & Val(txtcustcd.Tag) & ""
  Sql = Sql & " and a.Status='A' and a.Branchid='" & gBranchID & "' and a.compid='" & gCompID & "'"
  Sql = Sql & " and a.ValidTo >=convert(datetime,'" + Format(ServerDate, "yyyy-MM-dd") + "',111)"
  Sql = Sql & " union (select a.OrderType,a.OrderNo,a.OrderDate "
  Sql = Sql & " from dbo.fn_SAL_JobWorkOrder_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a where a.CustomerCode=" & Val(txtcustcd.Tag) & " "
  Sql = Sql & " and a.Status='A' and a.Branchid='" & gBranchID & "'"
  Sql = Sql & " and a.compid='" & gCompID & "' and a.ValidTo >=convert(datetime,'" + Format(ServerDate, "yyyy-MM-dd") + "',111))"
  Sql = Sql & " order by a.OrderDate,a.OrderType,a.OrderNo"
  If FetchValue_New(False, db.ConnectionString, Sql, "111", "List of Order", "Order Type", "Order No", "Order Date", "Customer") = True Then
    Sql = "select x.ItemCode,x.ItemDescription,x.UOMDesc,x.UOMDecimal,x.OrderQty,x.Rate,"
    Sql = Sql & " x.AssessableRate,x.Status,x.OrderType,x.OrderNo,x.OrderDate,"
    Sql = Sql & " (x.OrderQty-IsNull(x.DCQty, 0)) BalQty,x.OrderingType,0 as ItemGroupKey from (select a.ItemCode,"
    Sql = Sql & " c.ItemDescription,c.UOMDesc,c.UOMDecimal,a.OrderQty,a.Rate,a.AssessableRate,"
    Sql = Sql & " a.Status,a.OrderType,a.OrderNo,a.OrderDate,(select sum(d.DCQty) from"
    Sql = Sql & " SAL_DC_DTL d where d.ItemCode=a.ItemCode and d.Ordertype=a.OrderType"
    Sql = Sql & " and d.OrderNo=a.OrderNo and d.OrderNo=a.OrderNo and d.OrderDate=a.OrderDate "
    Sql = Sql & " and d.Branchid=a.Branchid and d.Compid=a.Compid) DCQty,"
    Sql = Sql & " a.OrderingType from dbo.fn_SAL_SalesOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a left outer join MATERIAL_MASTER c "
    Sql = Sql & " on a.ItemCode=c.ItemCode and a.Branchid=c.Branchid and a.Compid=C.Compid "
    Sql = Sql & " where a.OrderType='" & Trim$(getvalue(1)) & "' and a.OrderNo=" & Val(getvalue(2)) & ""
    Sql = Sql & " and a.OrderDate=convert(datetime,'" + Format(getvalue(3), "yyyy-MM-dd") + "',111)"
    Sql = Sql & " and c.Status = 1 and a.Branchid='" & gBranchID & "' and a.Compid='" & gCompID & "'"
    Sql = Sql & " and a.Status='A') x"
    Sql = Sql & " union (select x1.SentItemCode,x1.ItemDescription,x1.UOMDesc,x1.UOMDecimal,"
    Sql = Sql & " x1.OrderQty, x1.Rate, x1.AssRate, x1.Status, x1.OrderType, x1.OrderNo,"
    Sql = Sql & " x1.OrderDate,(x1.OrderQty-IsNull(x1.DCQty, 0)) BalQty, x1.OrderingType,x1.ItemGroupKey from"
    Sql = Sql & " (select a1.ItemGroupKey,a1.SentItemCode, c1.ItemDescription,c1.UOMDesc,c1.UOMDecimal,"
    Sql = Sql & " a1.OrderQty, a1.Rate, a1.AssRate, a1.Status,a1.OrderType,a1.OrderNo,a1.OrderDate,"
    Sql = Sql & " (select sum(d1.DCQty) from SAL_DC_DTL d1 where d1.ItemCode=a1.SentItemCode "
    Sql = Sql & " and d1.Ordertype=a1.OrderType and d1.OrderNo=a1.OrderNo and d1.OrderNo=a1.OrderNo "
    Sql = Sql & " and d1.OrderDate=a1.OrderDate and d1.Branchid=a1.Branchid and d1.Compid=a1.Compid) DCQty, "
    Sql = Sql & " a1.OrderingType From dbo.fn_SAL_JobWorkOrder('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a1 "
    Sql = Sql & " Left outer join MATERIAL_MASTER c1 on a1.SentItemCode=c1.ItemCode and a1.Branchid=c1.Branchid and a1.Compid=c1.Compid"
    Sql = Sql & " Where a1.OrderType='" & Trim$(getvalue(1)) & "' and a1.OrderNo=" & Val(getvalue(2)) & " "
    Sql = Sql & " and a1.OrderDate=convert(datetime,'" & Format(getvalue(3), "yyyy-MM-dd") & "',111) and c1.Status = 1"
    Sql = Sql & " and a1.Branchid='" & gBranchID & "' and a1.Compid='" & gCompID & "'"
    Sql = Sql & " and a1.Status='A' and a1.Rate <> 0) x1)"
    If FetchValue_New(True, db.ConnectionString, Sql, "1100000000000", "List of Item's", "Item Code", "Item Description", "Received Item", "Received Item Description") = True Then
      spItem.MaxRows = UBound(getvalue, 1)
      For i = 1 To UBound(getvalue, 1)
        spItem.SetText eItem.mItemCode, i, Trim(getvalue(i, 1))  'item code
        spItem.SetText eItem.mItemDesc, i, getvalue(i, 2) 'item desc
        Call Spread_Set_Decimal(spItem, Val(getvalue(i, 4)), eItem.mStockQty, i, eItem.mDCQty)
        Call Spread_Set_Decimal(spItem, Val(getvalue(i, 4)), eItem.mBalQty, i)
        spItem.SetText eItem.mRate, i, getvalue(i, 6)  'Rate
        spItem.SetText eItem.mARate, i, getvalue(i, 7) 'Assessable Rate
        spItem.SetText eItem.mOType, i, getvalue(i, 9)  'Order Type
        spItem.SetText eItem.mONo, i, getvalue(i, 10) 'Order No
        spItem.SetText eItem.mODate, i, getvalue(i, 11)  'Order Date
        spItem.SetText eItem.mBalQty, i, getvalue(i, 12) ' Bal Qty
        spItem.SetText eItem.mSchType, i, getvalue(i, 13) ' Schedule Type
        spItem.SetText eItem.mGroupKey, i, getvalue(i, 14) ' Item GroupKey
      Next i
      fraDC.Enabled = False
      fraSource.Enabled = False
      fraCustomer.Enabled = False
      spItem.Enabled = True
      txtRemarks.Enabled = True
      cmdSave.Enabled = True
      spItem.SetFocus
      Spread_Set_Focus spItem, 1, 4
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Design_spItem(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spItem


spItem.DblClick +=  new EventHandler(spItem_DblClick)
spItem.EditMode +=  new EventHandler(spItem_EditMode)
spItem.KeyDown +=  new EventHandler(spItem_KeyDown)

End Sub

Private Sub InitializeSpreads()

	Call Design_spItem("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\SalesDeliveryChallan.frm", spItem, "")

End Sub
