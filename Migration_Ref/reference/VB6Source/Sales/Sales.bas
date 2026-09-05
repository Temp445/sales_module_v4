Attribute VB_Name = "mdlSales"
Option Explicit
Option Compare Text

''For Access Database
'  Public WS As Workspace
'  Public MDB As Database

'For saving ModuleCode in Formula Tables
  Public gModuleCode As Integer
'  Public FrmFlg As String
'  Public RptFlg As String
'  Public RptFrmFlg As Form
'  Public mnuFlg As String
'  Public DocFlg As String
'  Public Cmdflag As String

  Public CustProdFlag As String
  Public Pen_Enquiry As Long
  Public mDocSeq(7) As Byte
  Public mDocCat(5) As Byte
  Public SumColrFlag As String
  Public TmpChart As MSChart
  Public TmpSpread As AceSpread
  Public GphPrint As Byte

  Public pageno As Long 'For Abstract Report generate PageNo
'  Public GridCount As Integer 'For Abstract Report generate cellCount
'  Public TotalPages As Integer 'For Abstract Report Total pages

  Public Enum eSalesCallType
    mExportOrder = 1
    mSalesOrder = 2
  End Enum

  Public Enum eParamID
    mQxR = 4
    mQxAR = 5
    mGV = 7
    mAGV = 8
    mIGST = 9
    mCGST = 10
    mSGST = 11
  End Enum
  
  Public Enum eFormList
    mSalEnquiry = 0
    mJWEnquiry = 1
    mSalQuotation = 2
    mJWQuotation = 3
    mSalOrder = 4
    mSalOrderExport = 5
    mJWOrder = 6
    mSalOrderAmendt = 7
    mJWOrderAmendt = 8
    mSchEntry = 9
    mSalPlanCreate = 10
    mSalPlanChange = 11
    mWeeklyPlan = 12
    mSalDC = 13
    mSalInv = 14
    mSalInvExport = 15
    mInvChange = 16
    mInventScheEntry = 17
    mContractReview = 18
    mDownstreamDocument = 19
  End Enum

Public Function gFinStart() As Date
' On Error Resume Next
  gFinStart = CDate("01-" & Left$(gFinyear, 7))
End Function

Public Sub sButtonsAction()
  
  Try

  MDISales.Toolbar1.Buttons(1).Enabled = True
  MDISales.Toolbar1.Buttons(2).Enabled = True
  MDISales.Toolbar1.Buttons(3).Enabled = True
  MDISales.Toolbar1.Buttons(4).Enabled = True
  MDISales.Toolbar1.Buttons(5).Enabled = True
  MDISales.Toolbar1.Buttons(6).Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Function AlphaNumChk(ByVal KeyAscii As Integer,ByVal  ParamArray ExcludeCharacters()) As Integer
  Dim CharCnt As Integer
  
' On Error Resume Next
  AlphaNumChk = 0
  Select Case KeyAscii
    Case 8, 13, 32 To 127     'BkSpace, Enter, Space to Delete
      For CharCnt = LBound(ExcludeCharacters) To UBound(ExcludeCharacters)
        If KeyAscii = Asc(ExcludeCharacters(CharCnt)) Then
          Exit Function
        End If
      Next CharCnt
  End Select
  AlphaNumChk = KeyAscii
End Function

Public Sub CurrDate(FormName)
  Dim MyCtrl As Control
  Dim CurrDate As Date
  
  CurrDate = Format(ServerDate, "dd/MM/yyyy")
  
  For Each MyCtrl In FormName.Controls
    If TypeOf MyCtrl Is DTPicker Then
      MyCtrl.Value = CurrDate
    End If
  Next
End Sub
Public Sub Form_Clear(FormName)
  Dim MyCtrl As Control

  For Each MyCtrl In FormName.Controls
    If TypeOf MyCtrl Is TextBox Then
      MyCtrl.Text = ""
MyCtrl.Tag = ""
    ElseIf TypeOf MyCtrl Is DTPicker Then
      If MyCtrl.CheckBox = True Then
        MyCtrl.Value = Null
      Else
        MyCtrl.Value = Format(ServerDate, "dd/MM/yyyy")
      End If
    ElseIf TypeOf MyCtrl Is fpSpread Then
      MyCtrl.Row = 1
MyCtrl.Row2 = MyCtrl.MaxRows

      MyCtrl.Col = 1
MyCtrl.Col2 = MyCtrl.MaxCols
      MyCtrl.BlockMode = True
      MyCtrl.Action = 12
      MyCtrl.BlockMode = False
'    ElseIf TypeOf MyCtrl Is MSFlexGrid Then
'      MyCtrl.Clear
    ElseIf TypeOf MyCtrl Is CheckBox Then
      MyCtrl.Value = False
    ElseIf TypeOf MyCtrl Is ListBox Or TypeOf MyCtrl Is ComboBox Then
      MyCtrl.ListIndex = -1
    End If
  Next
End Sub
Public Function LoadDocType(ByVal fFormFrom As eFormList) As Boolean
  Dim rsDocType As AceADODB.Recordset
  Dim Sql As String
  Dim fForm As Form
  Dim fDocTypeCombo As ComboBox
  Dim fPertainingTo As String
  Dim fModuleName As String
  
  Try
  LoadDocType = False
  
  fModuleName = "SAL"
   
  Select Case fFormFrom
    Case eFormList.mSalEnquiry
      Set fForm = frmSalesEnquiry
      Set fDocTypeCombo = fForm.cmbEnqType
      fPertainingTo = "E"
    Case eFormList.mJWEnquiry
      Set fForm = frmJobworkEnquiry
      Set fDocTypeCombo = fForm.cmbEnqType
      fPertainingTo = "N"
    Case eFormList.mSalQuotation
      Set fForm = frmSalesQuotation
      Set fDocTypeCombo = fForm.cmbQuoType
      fPertainingTo = "Q"
    Case eFormList.mJWQuotation
      Set fForm = frmJobworkQuotation
      Set fDocTypeCombo = fForm.cmbQuoType
      fPertainingTo = "U"
    Case eFormList.mSalOrder
      Set fForm = frmSalesOrder
      Set fDocTypeCombo = fForm.cmbordtype
      fPertainingTo = "O"
    Case eFormList.mSalOrderExport
      Set fForm = frmSalesOrder
      Set fDocTypeCombo = fForm.cmbordtype
      fPertainingTo = "O"
    Case eFormList.mJWOrder
      Set fForm = frmJobworkOrder
      Set fDocTypeCombo = fForm.cmbordtype
      fPertainingTo = "J"
    Case eFormList.mSalOrderAmendt
      Set fForm = frmSalesOrderAmendment
      Set fDocTypeCombo = fForm.cmbamdtype
      fPertainingTo = "A"
    Case eFormList.mJWOrderAmendt
      Set fForm = frmJobworkOrderAmendment
      Set fDocTypeCombo = fForm.cmbamdtype
      fPertainingTo = "M"
'    Case eFormList.mSchEntry
'      Set fForm = frmSchduleEntry
'      Set fDocTypeCombo = fForm.CmbSchtype
'      fPertainingTo = "S"
    Case eFormList.mSalPlanCreate
      Set fForm = frmSalesPlan
      Set fDocTypeCombo = fForm.CmbPlanType
      fPertainingTo = "YP"
    Case eFormList.mSalPlanChange
      Set fForm = frmSalesPlan
      Set fDocTypeCombo = fForm.CmbPlanType
      fPertainingTo = "YP"
'    Case eFormList.mWeeklyPlan
'      Set fForm = frmWeeklyPlan
'      Set fDocTypeCombo = fForm.cmbWeekType
'      fPertainingTo = "W"
    Case eFormList.mSalDC
      Set fForm = frmSalesDeliveryChallan
      Set fDocTypeCombo = fForm.cmbDCType
      fPertainingTo = "D"
    Case eFormList.mSalInv
      Set fForm = frmSalesInvoice
      Set fDocTypeCombo = fForm.CmbInvType
      fPertainingTo = "I"
    Case eFormList.mSalInvExport
      Set fForm = frmSalesInvoice
      Set fDocTypeCombo = fForm.CmbInvType
      fPertainingTo = "I"
    Case eFormList.mInvChange
      Set fForm = frmInvoiceChange
      Set fDocTypeCombo = fForm.CmbInvType
      fPertainingTo = "I"
    Case eFormList.mInventScheEntry
      Set fForm = frmInventorySchduleEntry
      Set fDocTypeCombo = fForm.CmbSchtype
      fPertainingTo = "S"
    Case eFormList.mContractReview
      Set fForm = frmContractReview
      Set fDocTypeCombo = fForm.cmbDocType
      fPertainingTo = "CR"
    Case eFormList.mDownstreamDocument
      Set fForm = frmDownstreamDocument
      Set fDocTypeCombo = fForm.cmbDocType
      fPertainingTo = "DD"
  End Select
  
  Sql = "select DocumentType from COM_DOCUMENTTYPE_MASTER where PertainingTo='" & fPertainingTo & "'"
  Sql = Sql & " and ModuleName='" & fModuleName & "' and BranchID='" & gBranchID & "' and CompID = '" & gCompID & "'"
  Set rsDocType = New AceADODB.Recordset
  If RSOpen(rsDocType, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsDocType.EOF Then
    fDocTypeCombo.Clear
    Do Until rsDocType.EOF
      fDocTypeCombo.AddItem rsDocType("DocumentType")
      rsDocType.MoveNext
    Loop
    fDocTypeCombo.ListIndex = 0
  Else
    sShowMessage "Please Create Document Type for " & UCase(frmDocumentTypeMaster.PerToCode_Name(fPertainingTo)) & " in Document Type Master."
    GoTo ExitHere
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsDocType = Nothing
End Try
End Function

Public Function LastDate(ByVal CurDate As Date) As Date
  Dim aDate As Date
  
  aDate = DateAdd("m", 1, CurDate)
  aDate = Format(aDate, "yyyy-MM-01")
  LastDate = DateAdd("d", -1, aDate)
  
End Function

Public Function NumValChk(ByVal KeyAscii As Integer) As Integer
  If Not (KeyAscii > 47 And KeyAscii < 59 Or KeyAscii = 8) Then
    NumValChk = 0
    Exit Function
  End If
  NumValChk = KeyAscii
End Function

Public Function SaveTrim(ByVal StrInput As String,ByVal  Length As Integer) As String
  Try
  SaveTrim = Left$(Trim$(StrInput), Length)
  Exit Function
Catch ex As Exception
  Call ErrMsgShow(ex)
  SaveTrim = ""
  Exit Function
End Function

Public Function ComboLoad(ByVal flg As String, ByRef cmb As ComboBox) As Boolean
  Dim Sql As String
  Dim snap As New AceADODB.Recordset
  
  ComboLoad = False
  Sql = "select Code,description from TERMSMASTER where typeflg =  '" + flg + "' AND status =1"
  Sql = Sql & " and Branchid ='" & gBranchID & "' and Compid = '" & gCompID & "'"
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  If Not snap.EOF Then
    cmb.Clear
    Do While Not snap.EOF
      cmb.AddItem snap("description")
      cmb.ItemData(cmb.NewIndex) = snap("Code")
      snap.MoveNext
    Loop
    cmb.ListIndex = 0
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function
Public Function LoadComboBox(ByRef fCombo As ComboBox, ByVal flg As String) As Boolean
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  
  LoadComboBox = False
  Sql = "select Code, Description from TERMSMASTER where typeflg='" & flg & "' and Status=1"
  Sql = Sql & " and Branchid='" & gBranchID & "' and Compid='" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then GoTo ExitHere
  
    fCombo.Clear
    If Not snap.EOF Then
      Do While Not snap.EOF
        fCombo.AddItem snap("description")
        fCombo.ItemData(fCombo.NewIndex) = snap("Code")
        snap.MoveNext
      Loop
      fCombo.ListIndex = 0
    End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Public Function NumChkSpl(ByVal KeyAscii As Integer) As Integer
  If Not ((KeyAscii > 46 And KeyAscii < 58) Or KeyAscii = 45 Or KeyAscii = 8 Or (KeyAscii > 41 And KeyAscii < 45) Or KeyAscii = 38 Or KeyAscii = 32) Then
    NumChkSpl = 0
    Exit Function
  End If
  NumChkSpl = KeyAscii
End Function

Public Function SingleValRet(ByVal SqlStr As String) As String
  Dim rs As New AceADODB.Recordset
  Try
  SingleValRet = ""
  If RSOpen(rs, SqlStr) = False Then GoTo ExitHere
  If Not rs.EOF Then
    SingleValRet = IIf(IsNull(rs(0)), "", rs(0))
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

 Sub Main()
  frmWait.Show
  DoEvents
  
  Call LoginSecurity
  
  If DBConnect = True Then
  
'  'For Access Database
'    Set WS = DBEngine.Workspaces(0)
'    Set MDB = WS.OpenDatabase(App.Path & "\SALDOC.MDB")
    
    Load MDISales
    MDISales.Show
    
    pPaper = vbPRPSA4
    pheight = 75
    iPageWidth = 80
  Else
    MsgBox "Please Contact System Administrator", vbOKOnly + vbCritical, "Program Exiting..."
    End         'Close the Application
  End If
  Unload frmWait
End Sub

Public Function FixString(ByVal iVal As String,ByVal  ReqStrLen As Integer,ByVal  FillSingleChar As String) As String
  Dim str As String
  Try
  FixString = ""
  str = String(ReqStrLen, Left$(FillSingleChar, 1))
  Mid(str, (ReqStrLen - Len(iVal) + 1)) = iVal
  FixString = str
  Exit Function
Catch ex As Exception
  Call ErrMsgShow(ex)
  Exit Function
End Function

Public Function TranSub(ByVal VTSNo As Integer,ByVal  DocType As String,ByVal  DocNo As Integer,ByVal  DocDate As Date,ByVal  VrRoute As String,ByVal  LrCode As Integer,ByVal  amt As Double,ByVal  DrCr As String,ByVal  compcode As String) As Boolean
  Dim cmd As AceADODB.Command
  Dim snap As AceADODB.Recordset
  Dim Sql As String
  
  Try

  TranSub = False
  
  Set snap = New AceADODB.Recordset
  Set cmd = New AceADODB.Command
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  
  Sql = "select DocNo from FAS_VoucherTranSub where DocNo = " & DocNo & ""
  Sql = Sql + " and VrRoute= '" + VrRoute + "' and LrCode = " & LrCode & ""
  Sql = Sql + " and CompCode = '" + compcode + "'"
  If snap.State = adStateOpen Then
 snap.Close
End If
  snap.Open Sql, db, adOpenStatic, adLockReadOnly
  commClear cmd
  If snap.EOF Then
    cmd.CommandText = "FAS_Ins_VoucherTranSub"
  Else
    cmd.CommandText = "FAS_Up_VoucherTranSub"
  End If
  cmd.Parameters.Append cmd.CreateParameter("DocType", adVarChar, adParamInput, 3, DocType)
  cmd.Parameters.Append cmd.CreateParameter("DocNo", adInteger, adParamInput, , DocNo)
  cmd.Parameters.Append cmd.CreateParameter("DocDate", adDBTimeStamp, adParamInput, , Format(DocDate, "yyyy-MM-dd"))
  cmd.Parameters.Append cmd.CreateParameter("VTSNo", adInteger, adParamInput, , VTSNo)
  cmd.Parameters.Append cmd.CreateParameter("VrRoute", adVarChar, adParamInput, 42, VrRoute)
  cmd.Parameters.Append cmd.CreateParameter("LrCode", adInteger, adParamInput, , LrCode)
  If DrCr = "C" Then
    cmd.Parameters.Append cmd.CreateParameter("DrAmt", adCurrency, adParamInput, , 0)
    cmd.Parameters.Append cmd.CreateParameter("CrAmt", adCurrency, adParamInput, , amt)
  Else
    cmd.Parameters.Append cmd.CreateParameter("DrAmt", adCurrency, adParamInput, , amt)
    cmd.Parameters.Append cmd.CreateParameter("CrAmt", adCurrency, adParamInput, , 0)
  End If
  cmd.Parameters.Append cmd.CreateParameter("CompCode", adChar, adParamInput, 12, compcode)
  cmd.Parameters.Append cmd.CreateParameter("UserId", adVarChar, adParamInput, 20, gUserID)
  cmd.Parameters.Append cmd.CreateParameter("TranStamp", adDBTimeStamp, adParamInput, , Format(DocDate, "yyyy-MM-dd HH:mm:ss"))
  cmd.Execute
  TranSub = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
End Try
End Function

Public Function Check_ListValue(ByVal lst As ListBox)
  Dim count As Integer
  Dim i As Long
  count = 0
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) = True Then count = 1
Exit For
  Next
  If count = 0 Then
    MsgBox "Select atleast One Code", vbInformation + vbOKOnly, "PAYROLL"
    lst.SetFocus
    Check_ListValue = False
    Screen.MousePointer = 0
    Exit Function
  End If
  Check_ListValue = True
End Function

Public Function GetMaxNo(ByVal TblName As String,ByVal  FldName As String,Optional ByVal cmpCode As String = "",Optional ByVal ReqCond As String = "") As String
  Dim Sql As String
  Dim adodbrs As AceADODB.Recordset
  Dim MaxNo As String
  
  Try
  
  GetMaxNo = 0
  
  If cmpCode = "" Then
    GoTo ExitHere
  End If
  
  Sql = "Select Max(" & Trim$(FldName) & ") from " & Trim$(TblName)
  If Len(Trim$(cmpCode)) > 0 And Len(Trim$(ReqCond)) > 0 Then
    Sql = Sql & " Where CompCode='" & Trim$(cmpCode) & "' and " & Trim$(ReqCond)
  ElseIf Len(Trim$(cmpCode)) = 0 And Len(Trim$(ReqCond)) > 0 Then
    Sql = Sql & " Where " & Trim$(ReqCond)
  ElseIf Len(Trim$(cmpCode)) > 0 And Len(Trim$(ReqCond)) = 0 Then
    Sql = Sql & " Where CompCode='" & Trim$(cmpCode) & "'"
  End If
  Set adodbrs = New AceADODB.Recordset
  If RSOpen(adodbrs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If adodbrs.EOF Or IsNull(adodbrs(0)) Then
    MaxNo = "0"
  Else
    MaxNo = adodbrs(0)
  End If
  
  GetMaxNo = MaxNo
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set adodbrs = Nothing
End Try
End Function

Public Function Find_Value(ByVal tabname As String,ByVal  FName As String,ByVal  ParamArray Fldlist()) As String
  Dim tmpSQL As String
  Dim TmpRec As Recordset
  Dim k As Integer
  
  Try

  Find_Value = ""
  tmpSQL = "Select " + FName + " from " + tabname + " Where "
  For k = 1 To UBound(Fldlist) Step 2
    tmpSQL = tmpSQL + "" + Fldlist(k - 1) + " = '" + IIf(IsNull(Fldlist(k)), "", Trim(Fldlist(k))) + "' and "
  Next
  tmpSQL = Left(tmpSQL, Len(Trim(tmpSQL)) - 3)
  tmpSQL = tmpSQL + " and Branchid='" & gBranchID & "' and compid='" & gCompID & "'"
  Set TmpRec = New AceADODB.Recordset
  If RSOpen(TmpRec, tmpSQL) = False Then Exit Function
  TmpRec.ActiveConnection = Nothing
  If Not TmpRec.EOF Then Find_Value = IIf(IsNull(TmpRec(0)), "", TmpRec(0)) Else Find_Value = " "
  TmpRec.Close
  Set TmpRec = Nothing
  Exit Function
Catch ex As Exception
  sShowMessage Err.Number & " " & Err.Description
End Function

Public Sub PrintPackingList(ByVal OrdType As String,ByVal  ORDNo As Integer,ByVal  OrdDate As Date,ByVal  CustomerCode As Integer)
  Dim rs As AceADODB.Recordset
  Dim rt As AceADODB.Recordset
  Dim rd As AceADODB.Recordset
  Dim Sql As String
  Dim rCnt As Long
  Dim Qty As Double
  Dim NetWeight As Double
  Dim CompName As String
  Dim CurDate As Date
  
  Try
   
   CurDate = ServerDate
   
  iPageWidth = 80
  Printer.Orientation = vbPRORPortrait
  Printer.Width = 10 * 1440
  Printer.Height = 12 * 1440
  Printer.PrintQuality = vbPRPQDraft

  Sql = " select CompanyName, (Address1+', '+Address2) as Add1, City+'-'+Pincode as Add2, Convert(Varchar, Phone) as Add3 "
  Sql = Sql & "from company_master where compid = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  Dumrow 1
  'Printer.FontName = "Courier New"
  Printer.FontName = "Draft 10cpi"
  Printer.FontSize = 8
  Printer.FontBold = True
  
  Sql = "select isnull(RRNo,0) as RefNo,isNull(RRdate,0) as RefDate from SalesInvoice where"
  Sql = Sql & " InvoiceType='" & OrdType & "' and InvoiceNo=" & ORDNo & " and InvoiceDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  
  Set rd = New AceADODB.Recordset
  If RSOpen(rd, Sql) = False Then
    GoTo ExitHere
  End If
  
  
  Printer.Print Tab(2); CAlign(79, "PackingList");
  Dumrow 2
  Printer.Print Tab(2); Left$("Exporter", 38); Tab(41); Left$("Packing List No & Date", 40);
  
  Dumrow 1
  Printer.Print Tab(2); Left$(rs("CompanyName"), 38); Tab(41); Left$(CStr(ORDNo) + "/" + CStr(OrdDate), 40)
  CompName = rs("CompanyName")
  Printer.Print Tab(2); Left$(rs("Add1"), 38)
  Printer.Print Tab(2); Left$(rs("Add2"), 38); Tab(41); Left$("Buyer's Order No. & Date", 40)
  Printer.Print Tab(2); Left$(rs("Add3"), 38); Tab(41); Left$(CStr(rd("RefNo")) + "/" + CStr(rd("RefDate")), 40)
  Printer.Print Tab(41); Left$("Other Reference(s)", 40)
  
  Sql = "select Type, Name, Address1, Address2, Address3,"
  Sql = Sql & " City, PinCode, Phone, Email,OrderFlag,Country"
  Sql = Sql & " from Sal_PackingList_Address where Type='Consignee' and"
  Sql = Sql & " OrderType='" & OrdType & "' and OrderNo=" & ORDNo & " and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
 
  Sql = "select Type, Name, Address1, Address2, Address3,"
  Sql = Sql & " City, PinCode, Phone, Email,OrderFlag "
  Sql = Sql & " from Sal_PackingList_Address where Type='Buyer' and"
  Sql = Sql & " OrderType='" & OrdType & "' and OrderNo=" & ORDNo & " and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  
  If RSOpen(rd, Sql) = False Then
    GoTo ExitHere
  End If
  
  Printer.Print String(iPageWidth, "-")
  
  Printer.Print Tab(2); Left$("CONSIGNEE", 38); Tab(41); Left$("BUYER (if Other than Consignee)", 40)
  Printer.Print Tab(2); String(38, "-"); Tab(41); String(40, "-")
  Printer.Print Tab(2); Left$(rs("Name"), 38); Tab(41); Left$(rd("Name"), 40)
  Printer.Print Tab(2); Left$(rs("Address1") + rs("Address2"), 38); Tab(41); Left$(rd("Address1") + rd("Address2"), 40)
  Printer.Print Tab(2); Left$(rs("Address3"), 38); Tab(41); Left$(rd("Address3"), 40)
  Printer.Print Tab(2); Left$(rs("City") + "-" + rs("PinCode"), 38); Tab(41); Left$(rd("City") + "-" + rd("PinCode"), 40)
  Printer.Print Tab(2);
  
  Printer.Print String(iPageWidth, "-")
  
  Sql = "select CarriageName, CarriageNo, DischargePort, ReceivingPlace, LoadingPort,"
  Sql = Sql & " DestinationPlace, Reference, Buyer, Consignee ,OrderFlag "
  Sql = Sql & " from Sal_PackingList where OrderType='" & OrdType & "' and OrderNo=" & ORDNo & " and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  Set rt = New AceADODB.Recordset
  If RSOpen(rt, Sql) = False Then
    GoTo ExitHere
  End If
 
  Printer.Print Tab(2); Left$("Pre-Carriage", 18); Tab(21); Left$("Place of receipt", 19); Tab(41); Left$("Country of Origin", 19); Tab(61); Left$("Country of Final", 20)
  Printer.Print Tab(2); Left$("By", 18); Tab(21); Left$("by Pre-Carrier", 19); Tab(41); Left$("Of Goods", 19); Tab(61); Left$("Destination", 20)
  Printer.Print Tab(2); String(18, "-"); Tab(21); String(19, "-"); Tab(41); String(19, "-"); Tab(61); String(20, "-")
  
  
  Sql = " select c.OrderType,c.OrderNo,c.OrderDate,b.DeliveryDesc,b.PaymentDesc from sal_Invoice_dtl c"
  Sql = Sql & " left outer join ("
  Sql = Sql & " select a.OrderType,a.OrderNo,a.OrderDate,isNull((select b3.Description from TERMSMASTER b3"
  Sql = Sql & " where b3.TypeFlg='Delivery Instruction' and b3.Code=a.DeliveryCode"
  Sql = Sql & " and b3.Branchid=a.Branchid and b3.Compid=a.Compid), '') DeliveryDesc,"
  Sql = Sql & " isNull((select b4.Description from TERMSMASTER b4"
  Sql = Sql & " where b4.TypeFlg='Payment Terms' and b4.Code=a.PaymentCode"
  Sql = Sql & " and b4.Branchid=a.Branchid and b4.Compid=a.Compid), '') PaymentDesc"
  Sql = Sql & " from dbo.fn_SAL_Order_Hdr('" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "' ,'" & Format(CurDate, "yyyy-MM-dd") & "','" & gBranchID & "','" & gCompID & "','" & gFinyear & "') a"
  Sql = Sql & " ) b on b.OrderType=c.OrderType and b.OrderNo=c.OrderNo and b.OrderDate=c.OrderDate "
  Sql = Sql & " where c.InvoiceType='" & OrdType & "' and c.InvoiceNo=" & ORDNo & " and c.InvoiceDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and c.Compid='" & gCompID & "' and c.Branchid='" & gBranchID & "' and  c.finyear='" & gFinyear & "'"
  If RSOpen(rd, Sql) = False Then
    GoTo ExitHere
  End If
  
  
  Printer.Print Tab(2); Left$(rt("CarriageName"), 18); Tab(21); Left$(rt("ReceivingPlace"), 19); Tab(41); Left$("India", 19); Tab(61); Left$(CheckNull(rs("Country")), 20)
  Printer.Print
  Printer.Print Tab(2); Left$("Vessel/Flight No", 18); Tab(21); Left$("Port of Loading", 19)
  Printer.Print Tab(2); String(18, "-"); Tab(21); String(19, "-")
  Printer.Print Tab(2); Left$(rt("CarriageNo"), 18); Tab(21); Left$(rt("LoadingPort"), 19)
  Printer.Print
  Printer.Print Tab(2); Left$("Port Of Discharge", 18); Tab(21); Left$("Final Destination", 19); Tab(41); Left$("Terms of Delivery and Payment", 40)
  Printer.Print Tab(2); String(18, "-"); Tab(21); String(19, "-"); Tab(41); String(40, "-")
  Printer.Print Tab(2); Left$(rt("DischargePort"), 18); Tab(21); Left$(rs("City") + "," + rs("Pincode"), 19); Tab(41); Left$(rd("DeliveryDesc"), 40)
  Printer.Print Tab(41); Left$(rd("PaymentDesc"), 40)
  Printer.Print
  Printer.Print
  
  Sql = "select ContainerNo, PackageKind, GoodsDescription,"
  Sql = Sql & "Quantity, NetWeight, Length, Width, Height,OrderFlag from Sal_PackingList_Goods"
  Sql = Sql & " where OrderType='" & OrdType & "' and OrderNo='" & ORDNo & "' and OrderDate='" & Format(OrdDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and Compid='" & gCompID & "' and Branchid='" & gBranchID & "' and  finyear='" & gFinyear & "'"
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  
  
  Printer.Print String(iPageWidth, "-")
  Printer.Print Tab(2); Left$("Container", 9); Tab(12); Left$("Kind", 12); Tab(25); Left$("Description of Goods", 29); Tab(55); Left$("Quantity", 9); Tab(65); Left$("Nt.Wt", 9); Tab(75); Left$("Gr.Wt", 5)
  Printer.Print Tab(2); Left("No", 9); Tab(12); Left$("Of Package", 12); Tab(55); Left$("in Pcs", 9); Tab(65); Left$("in KGS", 9); Tab(75); Left$("in KGS", 5)
  
  Printer.Print String(iPageWidth, "-")
  Dumrow 1
  If Not rs.EOF Then
    For rCnt = 1 To rs.RecordCount
      Printer.Print Tab(2); Left$(rs("ContainerNo"), 9); Tab(12); Left$(rs("PackageKind"), 12); Tab(25); Left$(rs("GoodsDescription"), 29); Tab(55); Left$(rs("Quantity"), 9); Tab(65); Left$(rs("Netweight"), 9)
      Qty = Qty + rs("Quantity")
      NetWeight = NetWeight + rs("NetWeight")
      rs.MoveNext
    Next rCnt
  End If
  
  
  If rs.RecordCount < 5 Then
     Dumrow 5 - rs.RecordCount
  End If
  
  Printer.Print String(iPageWidth, "-")
  Printer.Print Tab(25); "TOTAL"; Tab(55); Left$(Qty, 9); Tab(65); Left$(NetWeight, 9)
  
  Printer.Print String(iPageWidth, "-")
  Printer.Print Tab(2); "Total Pack"
  
  Dumrow 1
  Printer.Print Tab(14); Left$("Length", 10); Tab(25); Left$("Width", 10); Tab(35); Left$("Height", 10)
  Printer.Print Tab(14); String(10, "-"); Tab(25); String(10, "-"); Tab(35); String(10, "-")
  Dumrow 1
  rs.MoveFirst
  
  Do While Not rs.EOF
    Printer.Print Tab(2); Left$(rs("ContainerNo"), 11); Tab(14); Left$(CStr(rs("Length")) & """", 10); Tab(25); Left$(CStr(rs("Width")) & """", 10); Tab(35); Left$(CStr(rs("Height")) & """", 10)
    rs.MoveNext
  Loop
  
  Dumrow 4
  
  Printer.Print Tab(40); CAlign(30, "Signature & Date")
  Printer.Print Tab(40); CAlign(30, CompName)
  
  Dumrow 5
  
  Printer.Print Tab(50); CAlign(30, "")           'Hard code the Director name according to Company.
  Printer.Print Tab(50); CAlign(30, "Director")
  
  Printer.EndDoc
  sShowMessage "Print Over"
Catch ex As Exception
  Call ErrMsgShow(ex)
  Printer.KillDoc

Finally

  Set rs = Nothing
  Set rd = Nothing
  Set rt = Nothing
End Try
End Sub

Public Function NumValFloatTextBox(ByRef TxtBox As TextBox,ByVal  KeyAscii As Integer,Optional ByVal MaxVal As Double,Optional ByVal NoOfDecimals As Integer,Optional ByVal MinVal As Double) As Integer
  Dim dotPos As Integer
Dim    CurStr As String
Dim  NewStr As String

  
  Try
  
  NumValFloatTextBox = 0
  
  If MaxVal < MinVal Then
    MsgBox "Minimum Parameter exceeds Maximum Parameter.", vbOKOnly + vbInformation, "Error while calling Function"
    GoTo ExitHere
  ElseIf NoOfDecimals < 0 Then
    MsgBox "Decimal Point Required should be greater than or equal to zero.", vbOKOnly + vbInformation, "Error while calling Function"
    GoTo ExitHere
  End If
  
    If Trim(TxtBox.Text) <> "" Then
      Select Case KeyAscii
        Case -1, 8, 45, 46, 48 To 57
          'continue Processing
        Case Else
          GoTo ExitHere
      End Select
      
      CurStr = Trim(TxtBox.Text)
      
      If TxtBox.SelLength = 0 Then      'No Selections made
        Select Case KeyAscii
          Case -1                 'Delete Key
            If TxtBox.SelStart = Len(CurStr) Then GoTo ExitHere
            NewStr = Left$(CurStr, TxtBox.SelStart) & Mid$(CurStr, TxtBox.SelStart + 2)
          Case 8                  'BackSpace Key
            If TxtBox.SelStart = 0 Then GoTo ExitHere
            NewStr = Left$(CurStr, TxtBox.SelStart - 1) & Mid$(CurStr, TxtBox.SelStart + 1)
          Case 45                 'Subtract Key
            If MinVal >= 0 Or TxtBox.SelStart > 0 Or Val(CurStr) = 0 Then GoTo ExitHere
            NewStr = "-" & CurStr
          Case 46                 'Decimal Point
            If NoOfDecimals = 0 Then GoTo ExitHere
            If InStr(1, CurStr, ".", vbTextCompare) > 0 Then GoTo ExitHere
            If Len(Mid$(CurStr, TxtBox.SelStart + 1)) > NoOfDecimals Then GoTo ExitHere
            NewStr = Left$(CurStr, TxtBox.SelStart) & Chr(KeyAscii) & Mid$(CurStr, TxtBox.SelStart + 1)
          Case Else               '0 to 9
            If KeyAscii = 48 And TxtBox.SelStart = 0 Then GoTo ExitHere
            dotPos = InStr(1, CurStr, ".", vbTextCompare)
            If dotPos <> 0 And TxtBox.SelStart >= dotPos Then
              If Len(Mid$(CurStr, dotPos + 1)) >= NoOfDecimals Then GoTo ExitHere
            End If
            NewStr = Left$(CurStr, TxtBox.SelStart) & Chr(KeyAscii) & Mid$(CurStr, TxtBox.SelStart + 1)
        End Select
      Else                        'If Selections Made
        Select Case KeyAscii
          Case -1, 8               'Delete Key & BackSpace Key
            NewStr = Left$(CurStr, TxtBox.SelStart) & Mid$(CurStr, TxtBox.SelStart + TxtBox.SelLength + 1)
          Case 45                 'Subtract Key
            If MinVal >= 0 Or TxtBox.SelStart > 0 Or Val(CurStr) = 0 Or Left$(CurStr, 1) = "-" Then GoTo ExitHere
            NewStr = "-" & CurStr
          Case 46                 'Decimal Point
            If NoOfDecimals = 0 Then GoTo ExitHere
            If InStr(1, TxtBox.SelText, ".", vbTextCompare) > 0 Then
              If Len(Mid$(CurStr, TxtBox.SelStart + TxtBox.SelLength + 1)) >= NoOfDecimals Then GoTo ExitHere
            Else
              If InStr(1, CurStr, ".", vbTextCompare) > 0 Then GoTo ExitHere
              If Len(Mid$(CurStr, TxtBox.SelStart + TxtBox.SelLength + 1)) >= NoOfDecimals Then GoTo ExitHere
            End If
            NewStr = Left$(CurStr, TxtBox.SelStart) & Chr(KeyAscii) & Mid$(CurStr, TxtBox.SelStart + TxtBox.SelLength + 1)
          Case Else               '0 to 9
            If KeyAscii = 48 And TxtBox.SelStart = 0 Then GoTo ExitHere
            dotPos = InStr(1, CurStr, ".", vbTextCompare)
            If dotPos > 0 And dotPos <= TxtBox.SelStart And InStr(1, TxtBox.SelText, ".", vbTextCompare) = 0 Then
              If Len(Mid$(CurStr, dotPos + 1, TxtBox.SelStart - dotPos) & Mid$(CurStr, TxtBox.SelStart + TxtBox.SelLength + 1)) >= NoOfDecimals Then GoTo ExitHere
            End If
            NewStr = Left$(CurStr, TxtBox.SelStart) & Chr(KeyAscii) & Mid$(CurStr, TxtBox.SelStart + TxtBox.SelLength + 1)
        End Select
      End If
      
      If Val(NewStr) > MaxVal Then GoTo ExitHere
      If Val(NewStr) < MinVal Then GoTo ExitHere
    End If
  
Catch ex As Exception
  MsgBox "Error Ocurred while verifying the input for valid float value." & Chr(13) & Err.Number & " : " & Err.Description, vbOKOnly + vbInformation, "Error in Function"

Finally

End Try
End Function

Public Function NumValFloatValidate(ByVal TxtBox As TextBox,Optional ByVal Max As Double,Optional ByVal NoOfDecimals As Integer,Optional ByVal Min As Double) As Boolean
Dim    CurValue As Double
Dim  ChrCnt As Integer

Dim    DecOccur As Boolean
Dim  CurChr As String

  
  Try
  
  NumValFloatValidate = False
    TxtBox.Text = Trim(TxtBox.Text)
CurValue = Val(TxtBox.Text)

    If TxtBox.Text = "" Then NumValFloatValidate = True
GoTo ExitHere
    
    DecOccur = False
    
    For ChrCnt = 1 To Len(TxtBox.Text)
      CurChr = Mid$(TxtBox.Text, ChrCnt, 1)
      Select Case Asc(CurChr)
        Case 45, 46, 48 To 57
          If CurChr = "-" Then
            If (Min >= 0) Or (ChrCnt > 1) Then GoTo ExitHere
          ElseIf CurChr = "." Then
            If DecOccur = True Then GoTo ExitHere Else DecOccur = True
          End If
        Case Else
          GoTo ExitHere
      End Select
    Next ChrCnt
    
    ChrCnt = InStr(1, TxtBox.Text, ".", vbTextCompare)
    If ChrCnt <> 0 Then If (Len(TxtBox.Text) - ChrCnt) > NoOfDecimals Then GoTo ExitHere
    If CurValue > Max Or CurValue < Min Then GoTo ExitHere
    
Catch ex As Exception
  MsgBox Err.Number & " : " & Err.Description

Finally

End Try
End Function

'Public Function StepProceedNext(frm As Form, Index As Integer) As Boolean
'
'  On Error GoTo ErrHandler
'  StepProceedNext = False
'
'  With frm
'    If .fraStep.UBound > Index Then
'      .fraStep(Index + 1).Enabled = True
'      .CmdStep(Index + 1).Enabled = True
'      .fraStep(Index).Enabled = False
'      .CmdStep(Index).Enabled = False
'    Else
'      .cmdView.Enabled = True
'      .fraStep(Index).Enabled = False
'      .CmdStep(Index).Enabled = False
'    End If
'  End With
'
'  StepProceedNext = True
'ExitHere:
'  On Error GoTo 0
'  Exit Function
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Function
'
'
'
Public Sub Spread_CellVerAlign_Set(ByRef mSpread As AceSpread,ByVal  mCellAlign As FPSpreadADO.TypeVAlignConstants,ByVal  mCol1 As Long,ByVal  mRow1 As Long,Optional ByVal mCol2 As Long = -2,Optional ByVal mRow2 As Long = -2)
' This Procedure is used to set the CellType & Alignment
  Try

    If mCol2 = -2 Then
      mCol2 = mCol1
    End If
    
    If mRow2 = -2 Then
      mRow2 = mRow1
    End If
    
    mSpread.Col = mCol1
mSpread.Col2 = mCol2

    mSpread.Row = mRow1
mSpread.Row2 = mRow2
    mSpread.BlockMode = True
      mSpread.TypeVAlign = mCellAlign
    mSpread.BlockMode = False
    
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub
