Attribute VB_Name = "mdlCommonFunctions"

Option Explicit
Option Compare Text

Public Enum eRoundOffOptions
  'Anand.S on 04/01/2005
  'required for RoundingOff Function
  anNone = 0
  anNearest = 1
  anUpward = 2
  anDownward = 3
End Enum

Public Enum eOprModeColorConstants
  'Anand.S on 17/01/2005
  'required for Spread Functions
  anDisplayOnly = 192       'Brown Color
  anSelectOnly = 16711680   'Blue Color
  anEditable = 0            'Black Color
  anSelectedit = 30464      'Dark Green Color
  anZoom = &HC000C0         'Pink Color
End Enum

Private gErrNo As Long         'Temporary Storage for Error Number
Private gErrDesc As String     'Temporary Storage for Error Description
    
'For saving ModuleCode in Formula Tables
  Public FrmFlg As String
  Public RptFlg As String
  Public RptFrmFlg As Form
  Public mnuFlg As String
  Public DocFlg As String
  Public Cmdflag As String
  Public GridCount As Integer
  Public totalpages As Integer
  
Public Enum eErpandTs
  mErp = 1
  mTs = 2
End Enum

Public Enum eBranch
  anBranchID = 1
  anBranchName = 2
End Enum

Public Enum eMsgBoxIcon
  mCritical = 16
  mQuestion = 32
  mExclamation = 48
  mInformation = 64
End Enum

Public Enum eEStatus
  mENew = 1
  mEAmendment = 2
  mEModify = 3
  mEAuthorisation = 4
  mEHold = 5
  mECompleted = 6
  mEShortClose = 7
  mECancel = 8
  mEActive = 9
  mEInactive = 10
  mEDelete = 11
End Enum

'Save Skip Message type
Public Enum eSaveSkipMsgType
  mNone = 0
  mCustom = 1
  mNoChanges = 2
  mHoldDelDoc = 3
End Enum

Public Function SetServerDate(Optional ByVal dtpAsOnDate As DTPicker,Optional ByVal ReturnDateFormatString As String = "") As Date
  Dim sdate As AceADODB.Recordset
  Dim sql As String

  Try

  'SetServerDate = Null
  
  sql = " Execute COM_Get_ServerDate"
  sql = sql & " @BranchID = '" & gBranchID & "',"
  sql = sql & " @CompID = '" & gCompID & "',"
  sql = sql & " @Finyear = '" & gFinyear & "'"
  Set sdate = New AceADODB.Recordset
  If RSOpen(sdate, sql) = False Then
    GoTo ExitHere
  End If
  
  If IsDate(sdate(0)) = False Then      'FinStart > "2000-01-01"
    'If (sDate(0) < FinStart Or Format(sDate(0), "dd/MM/yyyy") > FinEnd) Then
    sShowMessage "Date not Set Properly. Contact System Administrator.", mCritical
    GoTo ExitHere
    'End If
  End If
  
  If Not (dtpAsOnDate Is Nothing) Then
    dtpAsOnDate = Format(sdate(0), "dd/MM/yyyy HH:mm:ss")
  End If
  
  If Trim$(ReturnDateFormatString) = "" Then
    SetServerDate = sdate(0)
  Else
    SetServerDate = Format(sdate(0), ReturnDateFormatString)
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set sdate = Nothing
End Try
End Function

Public Function ReturnCommonNamedParametersSqlForSave(ByVal TranStamp As Date,Optional ByVal IsMaster As Boolean = False,Optional ByVal fEffStatus As String,Optional ByVal fEffDate As Date,Optional ByVal fIsTranLock As Boolean,Optional ByVal ReqUserID As Boolean = True,Optional ByVal Reqfinyear As Boolean = True,Optional ByVal PrevFinyear As String = "") As String
  Dim sql As String
  
  Try
  
  'For Master Forms
  If IsMaster = True Then
    sql = sql & " @EffStatus='" & Trim$(fEffStatus) & "',"
    sql = sql & " @EffDate='" & Format(Trim$(fEffDate), "yyyy-MMM-dd HH:mm:ss") & "',"
    sql = sql & " @isTranLock='" & IIf(fIsTranLock = True, 1, 0) & "',"
  End If
  
  If ReqUserID = True Then
    sql = sql & " @UserID='" & gUserID & "',"
  End If
  
  sql = sql & " @BranchID='" & gBranchID & "',"
  sql = sql & " @CompID='" & gCompID & "',"
  
   'For Transaction Forms
  If Reqfinyear = True Then
    If PrevFinyear = "" Then
      sql = sql & " @FinYear='" & gFinyear & "',"
    Else
      sql = sql & " @FinYear='" & PrevFinyear & "',"
    End If
  End If
  
  sql = sql & " @Transtamp='" & Format(TranStamp, "yyyy-MM-dd HH:mm:ss") & "'"
  
  ReturnCommonNamedParametersSqlForSave = sql
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function NoOfDecimal(ByVal sval As Double) As Integer
  Dim X As Double
Dim    DC As Integer
Dim  dcStr As String

  
  NoOfDecimal = 0
  X = Val(sval)
  DC = InStr(X, ".")
  If DC > 0 Then
    NoOfDecimal = IIf(Len(Mid(sval, DC + 1)) > 5, 5, Len(Mid(sval, DC + 1)))
  End If
  
End Function

Public Function AddCommonCmdParameters(ByRef cmd As AceADODB.Command,Optional ByVal ReqUserID As Boolean = True,Optional ByVal Reqfinyear As Boolean = True) As Boolean


  
  Try
  AddCommonCmdParameters = False
  
  If ReqUserID = True Then
    cmd.Parameters.Append cmd.CreateParameter("UserID", adVarChar, adParamInput, 10, gUserID)

  End If
  
  cmd.Parameters.Append cmd.CreateParameter("BranchID", adVarChar, adParamInput, 3, gBranchID)

  cmd.Parameters.Append cmd.CreateParameter("CompID", adVarChar, adParamInput, 3, gCompID)

  
  If Reqfinyear = True Then
    cmd.Parameters.Append cmd.CreateParameter("finyear ", adVarChar, adParamInput, 20, gFinyear)

  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function AddCommonCmdParameters_New(ByRef cmd As AceADODB.Command,Optional ByVal ReqUserID As Boolean = True,Optional ByVal Reqfinyear As Boolean = True,Optional ByVal PrevFinyear As String = "") As Boolean
  Dim CurDate As Date
  
  Try
  AddCommonCmdParameters_New = False
  
  If ReqUserID = True Then
    cmd.Parameters.Append cmd.CreateParameter("UserID", adVarChar, adParamInput, 10, gUserID)
  End If
  
  cmd.Parameters.Append cmd.CreateParameter("BranchID", adVarChar, adParamInput, 3, gBranchID)
  cmd.Parameters.Append cmd.CreateParameter("CompID", adVarChar, adParamInput, 3, gCompID)
  
  If Reqfinyear = True Then
    If PrevFinyear = "" Then
      cmd.Parameters.Append cmd.CreateParameter("finyear ", adVarChar, adParamInput, 20, gFinyear)
    Else
      cmd.Parameters.Append cmd.CreateParameter("finyear ", adVarChar, adParamInput, 20, PrevFinyear)
    End If
  End If
  
  CurDate = ServerDate
  cmd.Parameters.Append cmd.CreateParameter("TranStamp", adDBTimeStamp, adParamInput, 3, CurDate)

  AddCommonCmdParameters_New = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function AddCommonField(ByRef CurDate As Date, ByRef EffDate As Date, ByRef Effstatus As String,Optional ByVal ReqUserID As Boolean = True,Optional ByVal Reqfinyear As Boolean = True) As String

  
  Dim sql As String
  
  Try
  AddCommonField = ""
  
  sql = "@RevNo=0,"
  sql = sql & " @RevDate='" & Format(CurDate, "yyyy-MM-dd") & "',"
  sql = sql & " @EffStatus='" & Left(Trim$(Effstatus), 1) & "',"           'SetEffStatus(IIf(Trim$(Effstatus) = "Authorized", eEStatus.mEAuthorisation, IIf(Trim$(Effstatus) = "Hold", eEStatus.mEHold, IIf(Trim$(Effstatus) = "New", eEStatus.mENew, IIf(Trim$(Effstatus) = "Modify", eEStatus.mEModify, IIf(Trim$(Effstatus) = "Amendment", eEStatus.mEAmendment, IIf(Trim$(Effstatus) = "Cancel", eEStatus.mECancel, IIf(Trim$(Effstatus) = "Completed", eEStatus.mECompleted, eEStatus.mEDelete)))))))) & "',"
  sql = sql & " @EffDate='" & Format(EffDate, "yyyy-MM-dd") & "',"
  sql = sql & " @Transtamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
  If ReqUserID Then sql = sql & " @UserID='" & gUserID & "',"
  sql = sql & " @BranchID='" & gBranchID & "',"
  sql = sql & " @CompID='" & gCompID & "',"
  If Reqfinyear Then sql = sql & " @Finyear='" & gFinyear & "'"

  AddCommonField = sql
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Common_AssignGlobalVariables() As Boolean
  Dim rs As AceADODB.Recordset
  Dim sql As String
  
  Try
  Common_AssignGlobalVariables = False
  
  'Default Parameter For Print
  gPrHeight = ePaper.mPaper80Col
  pPaper = vbPRPSA4   'pPaper = vbPRPSFanfoldStdGerman
  pheight = gPrHeight
  iPageWidth = 80
  Orient = 1
  
  'Default Parameter For SpreadFormat
  spHeight = 54
  spOrient = 1
  gSpHeight = spHeight
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function AssignGlobalVariables() As Boolean
  Dim rs As AceADODB.Recordset
  Dim sql As String
  
  Try
  
  AssignGlobalVariables = False
  
  If Trim$(gCompName) = "" Then
    sql = " Select CompanyName, ReportId from Company_Master Where BranchID = '00' and CompID = '" & gCompID & "'"
    If RSOpen(rs, sql) = False Then
      GoTo ExitHere
    End If
    If Not rs.EOF Then
      gCompName = rs("CompanyName")
      gReportID = rs("ReportId")
    End If
  End If
  
  If gBranchID <> "00" Then
    sql = " Select CompanyName, ReportId from Company_Master Where BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "'"
    If RSOpen(rs, sql) = False Then
      GoTo ExitHere
    End If
    
    If Not rs.EOF Then
      gBranchName = rs("CompanyName")
      gReportID = rs("ReportId")
    End If
  End If
  
  '  If gBranchName <> "" Then gBranchName = "-" & gBranchName
  
  gCompName = gCompName & IIf(Trim$(gBranchName) = "", "", " - " & gBranchName)
    
  FinStart = CDate(Mid(gFinyear, 4, 4) & "-" & Left$(gFinyear, 2) & "-01")
  FinEnd = CDate(Right$(gFinyear, 4) & "-" & Mid(gFinyear, 9, 2) & "-31")
  
  gFasCompCode = gCompID & gBranchID '& Mid(gFinyear, 4, 4) & Right(gFinyear, 4)
  
  'Default Parameter For Print
  gPrHeight = ePaper.mPaper80Col
  pPaper = vbPRPSA4   'pPaper = vbPRPSFanfoldStdGerman
  pheight = gPrHeight
  iPageWidth = 80
  Orient = 1
  
  'Default Parameter For SpreadFormat
  spHeight = 54
  spOrient = 1
  gSpHeight = spHeight
  
  'GST Applicable
  IsGSTRequired = IIf(Val(GetOptionValue(oIsGSTRequired)) = "1", True, False)
  
  ' For Parameterwise Multiple No of Reading enter in Inspection Master
  IsReqIncomingInsMultipleReadingForSingleParameter = IIf(Val(GetOptionValue(eOptions.oIsReqIncomingInsMultipleReadingForSingleParameter)) = "1", True, False)
  IsReqFinalInsMultipleReadingForSingleParameter = IIf(Val(GetOptionValue(eOptions.oIsReqFinalInsMultipleReadingForSingleParameter)) = "1", True, False)
  IsReqPatrolInsMultipleReadingForSingleParameter = IIf(Val(GetOptionValue(eOptions.oIsReqPatrolInsMultipleReadingForSingleParameter)) = "1", True, False)
  IsReqReInsMultipleReadingForSingleParameter = IIf(Val(GetOptionValue(eOptions.oIsReqReInsMultipleReadingForSingleParameter)) = "1", True, False)

  gGSTStartDate = "01/07/2017"
  
  AssignGlobalVariables = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function BranchLoad(ByVal frmName As Form) As Boolean
  Dim sql As String
  Dim rs As AceADODB.Recordset
  
  Dim rCnt As Integer
  
  Try
  BranchLoad = False
  
  sql = " Select BranchID, CompanyName From Company_Master Where IsBranchAvail = 0 and BranchID <> '" & gBranchID & "' and CompID = '" & gCompID & "'"

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  If Not rs.EOF Then
      frmName.sprBranchDet.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        frmName.sprBranchDet.SetText eBranch.anBranchID, rCnt, rs("BranchID")
        frmName.sprBranchDet.SetText eBranch.anBranchName, rCnt, rs("CompanyName")
        rs.MoveNext
      Next rCnt
    frmName.fraBranchDetails.Visible = True
  Else
    frmName.fraBranchDetails.Visible = False
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function CAR_Save(ByVal cRefType As String,ByVal  cRefNo As Integer,ByVal  cRefDate As Date,ByVal  cItemcode As String,ByVal  cMachinecode As String,ByVal  cRaized As String,ByVal  mParamcode As Integer,ByVal  cParameter As String,ByVal  cFromTovalue As String,ByVal  cDefectValue As String,ByVal  cDeviation As Boolean,ByVal  cDueDate As Date) As Boolean
  Dim sql As String
  Dim CurrDate As Date
  Dim rs As AceADODB.Recordset
  Dim sDocno As Integer
  Dim sDocType As String

  Try
  CAR_Save = False
  
  'Get ServerDate
  CurrDate = ServerDate

  sql = "Select Max(DocNo) As DocNo From QC_CAR_Hdr Where RefType = '" & Trim$(cRefType) & "'"
  sql = sql & " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "' and Finyear = '" & gFinyear & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then GoTo ExitHere
  If Not rs.EOF Then
    sDocno = CheckNull(rs("DocNo"), 0) + 1
  Else
    'Continue
  End If
  Select Case Trim$(cRefType)
    Case "CCR"
      sDocType = "CC"
    Case "PIR"
      sDocType = "PC"
    Case "RIR"
      sDocType = "IC"
    Case "RPN"
      sDocType = "RPC"
    Case "SIR"
      sDocType = "SC"
    Case "FIN"
      sDocType = "FC"
    Case "IA"
      sDocType = "IAC"
    Case Else
      'Continue
  End Select

  'Insert - CAR Data
  sql = " Execute QC_Ins_CARHDR_comm"
  sql = sql + " @DocType = '" & Trim$(sDocType) & "',"
  sql = sql + " @DocNo = '" & Val(sDocno) & "',"
  sql = sql + " @DocDate = '" & Format(CurrDate, "yyyy-MM-dd") & "',  "
  sql = sql + " @RefType = '" & Trim$(cRefType) & "',"
  sql = sql + " @RefNo = '" & Val(cRefNo) & "',"
  sql = sql + " @RefDate = '" & Format(cRefDate, "yyyy-MM-dd") & "',"
  sql = sql + " @ItemCode = '" & Trim$(cItemcode) & "',"
  sql = sql + " @MachineCode = '" & Trim$(cMachinecode) & "',"
  sql = sql + " @RaisedBy = '" & Trim$(cRaized) & "',"
  sql = sql + " @CARBase = 0,"
  sql = sql + " @isNew = 0,"
  sql = sql + " @mParamCode = '" & Val(mParamcode) & "',"
  sql = sql + " @Parameter = '" & Trim$(cParameter) & "',"
  sql = sql + " @FromValueToValue = '" & Trim$(cFromTovalue) & "',"
  sql = sql + " @DefectValue = '" & Trim$(cDefectValue) & "',"
  If Trim$(cDeviation) = True Then
    sql = sql + " @DeviationStatus =  1,"
    sql = sql + " @DueDate = '" & Format(cDueDate, "yyyy-MM-dd HH:mm:ss") & "',"
  Else
    sql = sql + " @DeviationStatus = 0,"
    sql = sql + " @DueDate = Null,"
  End If

  sql = sql + " @ApprovedBy = '',"
  sql = sql + " @ApprovedDate = Null,"
  sql = sql + " @CauseEffectTargetDate = Null,"
  sql = sql + " @WhyWhyTargetDate = Null,"
  sql = sql + " @SimulationProblem = 0 ,"
  sql = sql + " @SimulationReason = '' ,"
  sql = sql + " @CARAdequacy = 0 ,"
  sql = sql + " @CARAdequacyDate = Null,"
  sql = sql + " @ReasonForNonAcceptance = '' ,"
  sql = sql + " @ReasonForNonAcceptanceDate = Null,"
  sql = sql + " @HorizondalDeployment = 0 ,"
  sql = sql + " @HorizondalDeploymentReason = '' ,"
  sql = sql + " @CARImplementation = '' ,"
  sql = sql + " @CARImplementationDate = Null,"
  sql = sql + " @CARVerified = 0,"
  sql = sql + " @CARVerifiedDate = Null,"
  sql = sql + " @CARClosed = 0,"
  sql = sql + " @CARClosedDate = Null,"
  sql = sql + " @Revisionno = 0,"
  sql = sql + " @UserID = '" & gUserID & "',"
  sql = sql + " @BranchID = '" & gBranchID & "',"
  sql = sql + " @CompID = '" & gCompID & "',"
  sql = sql + " @Finyear = '" & gFinyear & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere

  CAR_Save = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function CBinary(ByVal Degree As Long,ByVal  Minuts As Long,ByVal  Second As Long)
  'RAJAPRIYAN
  CBinary = Degree + (Minuts / 60) + (Second / 3600)
End Function

Public Function CDegree(ByVal Binary As Double) As String
  Dim X As Double
  Dim Y As Double
  Dim Z As Double
  'RAJAPRIYAN
  X = Int(Binary) 'Degree
  Y = Binary - X
  Y = Int(Y * 60) 'Minuts
  Z = Binary - (X + (Y / 3600))
  Z = Int(Z * 60) 'Second
  
  CDegree = X & Chr(176) & Format(Y, "00") & "'" & Format(Z, "00") & "''"
  
End Function

Public Function CheckAlphaNumeric(ByVal KeyAscii As Integer,ByVal  ParamArray ExcludeCharacters()) As Integer

  Dim CharCnt As Integer
  
' On Error Resume Next
  CheckAlphaNumeric = 0
  
  Select Case KeyAscii
    Case vbKeyBack, vbKeyReturn, vbKeySpace To vbKeyDelete     'BkSpace, Enter, Space to Delete
      For CharCnt = LBound(ExcludeCharacters) To UBound(ExcludeCharacters)
        If KeyAscii = Asc(ExcludeCharacters(CharCnt)) Then
          Exit Function
        End If
      Next CharCnt
  End Select
  
  CheckAlphaNumeric = KeyAscii
End Function

Public Function CheckApplicationBuildDate(ByVal MdlCode As eMdlCode) As Boolean
  

'  Dim Sql As String
'  Dim rs As ADODB.Recordset
'  Dim AppBuildDate As Date
'  Dim CurDate As Date
'
'  On Error GoTo ErrHandler
'  CheckApplicationBuildDate = False
'
'  Sql = App.Path & "\" & App.EXEName & ".exe"
'  If Dir$(Sql) <> "" Then
'    AppBuildDate = FileCreated(Sql)
'
'    CurDate = ServerDate
'    Sql = "Select OptionValue from COM_ERP_Options a"
'    Sql = Sql & " Where a.OptionDesc='LatestBuildDate' and a.ModuleCode='" & MdlCode & "' "
'    Sql = Sql & " and CreatedDate=(Select Max(CreatedDate) from COM_ERP_Options b"
'    Sql = Sql & " Where b.OptionDesc=a.OptionDesc and b.ModuleCode=a.ModuleCode)"
'
''    Sql = "Select OptionValue from dbo.fn_COM_ERP_Options('" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MMM-dd") & "', '" ''            & gBranchID & "', '" & gCompID & "')"

''    Sql = Sql & " Where ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
'    Set rs = New ADODB.Recordset
'    If RSOpen(rs, Sql) = False Then
'      GoTo ExitHere
'    End If
'
'    If rs.EOF Then
'      Sql = "Insert into COM_ERP_Options"
'      Sql = Sql & " Select " & GetMax("COM_ERP_Options", "Code", "ModuleCode", MdlCode) + 1 & ", " & MdlCode & ", "
'      Sql = Sql & " 'LatestBuildDate', '" & Format(AppBuildDate, "yyyy-MMM-dd HH:mm:ss") & "', 'NA', 'NA', 'DATETIME', '1', '"
'      Sql = Sql & Format(CurDate, "yyyy-MMM-dd") & "', '" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "', '"
'      Sql = Sql & gUserID & "', '" & gBranchID & "', '" & gCompID & "', '" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "'"
'      db.Execute Sql
'    Else
'      If AppBuildDate > CDate(rs("OptionValue")) Then   'Exe Date is Latest than stored
'        Sql = " Insert into COM_ERP_Options"
'        Sql = Sql & " Select " & GetMax("COM_ERP_Options", "Code", "ModuleCode", MdlCode, "OptionDesc", "LatestBuildDate") & ", " & MdlCode & ", "
'        Sql = Sql & " 'LatestBuildDate', '" & Format(AppBuildDate, "yyyy-MMM-dd HH:mm:ss") & "', 'NA', 'NA', 'DATETIME', '1', '"
'        Sql = Sql & Format(CurDate, "yyyy-MMM-dd") & "', '" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "', '"
'        Sql = Sql & gUserID & "', '" & gBranchID & "', '" & gCompID & "', '" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "'"
'        db.Execute Sql
'      ElseIf AppBuildDate < CDate(rs("OptionValue")) Then 'Exe Date is old
'        sShowMessage "Executable is Older than the Latest Available. Contact System Administrator"
'        GoTo ExitHere
'      Else
'        'Allow the Program Execute
'      End If
'    End If
'  End If
'
  CheckApplicationBuildDate = True
'
'ExitHere:
'  On Error GoTo 0
'  Set rs = Nothing
'  Exit Function
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
End Function

Public Function CheckNull(ByVal InputVal As Variant,Optional ByVal DefaultVal As Variant = "",Optional ByVal TrimString As Boolean = True) As String


  
' On Error Resume Next
  
  CheckNull = DefaultVal
  
  If Not IsNull(InputVal) Then
    If TrimString = True Then
      CheckNull = Trim$(InputVal)
    Else
      CheckNull = InputVal
    End If
  End If
  
End Function

Public Sub ClearForm(ByVal FormName As Form,ByVal  ParamArray ExclCntrlName() As Variant)
  Dim MyCtrl As Control
  Dim cCnt As Long
  Dim SkipCtrl As Boolean
  Dim CurDate As Date

  Try
  
  CurDate = ServerDate
  For Each MyCtrl In FormName.Controls
    SkipCtrl = False
    For cCnt = 0 To UBound(ExclCntrlName)
      If MyCtrl.Name = ExclCntrlName(cCnt) Then
        SkipCtrl = True
        Exit For
      End If
    Next cCnt
    
    If SkipCtrl = False Then
      If TypeOf MyCtrl Is TextBox Then
        'TextBox Text & Tag Clear
        MyCtrl.Tag = ""
        MyCtrl.Text = ""
      ElseIf TypeOf MyCtrl Is DTPicker Then
        'DatePicker Current Date Set
        MyCtrl.Value = Format(CurDate, "dd/mm/yyyy")
      ElseIf TypeOf MyCtrl Is fpSpread Then
        'Text Clear
        MyCtrl.ClearRange 1, 1, MyCtrl.MaxCols, MyCtrl.MaxRows, True
        'CellTag Clear
        MyCtrl.Col = 1
MyCtrl.Row = 1

        MyCtrl.Col2 = MyCtrl.MaxCols
MyCtrl.Row2 = MyCtrl.MaxRows
        MyCtrl.BlockMode = True
        MyCtrl.CellTag = ""
        MyCtrl.BlockMode = False
      ElseIf TypeOf MyCtrl Is CheckBox Then
        'CheckBox Clear
        MyCtrl.Value = False
      ElseIf TypeOf MyCtrl Is ListBox Or TypeOf MyCtrl Is ComboBox Then
        'ListBox / ComboBox Clear
        MyCtrl.ListIndex = -1
      End If
    End If
  Next
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Sub ClearTagLabel(ByVal FormName As Form)
  Dim MyCtrl As Control

  Try
  
  For Each MyCtrl In FormName.Controls
    If TypeOf MyCtrl Is Label Then
      If Left(MyCtrl.Name, 6) = "Taglbl" Then
        MyCtrl.Caption = ""
      Else
        'Continue
      End If
    Else
      'Continue
    End If
  Next
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Function commClear(ByVal comm As AceADODB.Command)
  Dim k As Integer
  
' On Error Resume Next
  For k = comm.Parameters.Count - 1 To 0 Step -1
    comm.Parameters.Delete k
  Next k
End Function

Public Function contrl_valid(ByRef Cntrl As Control, ByVal Msg As String) As Boolean

' On Error Resume Next
  contrl_valid = False
  
  If TypeOf Cntrl Is TextBox Then
    If Trim$(Cntrl.Text) = "" Then
      Cntrl.Text = ""
      Cntrl.SetFocus
      MsgBox Msg, vbInformation, App.Title
      Exit Function
    End If
  End If
  
  If TypeOf Cntrl Is ComboBox Then
    If Trim$(Cntrl.Text) = "" Then
      Cntrl.SetFocus
      MsgBox Msg, vbInformation, App.Title
      Exit Function
    End If
  End If
  
'  If TypeOf Cntrl Is MaskEdBox Then
'    If Trim$(Cntrl.ClipText) = "" Then
'      Cntrl.SetFocus
'      MsgBox msg, vbInformation, App.Title
'      Exit Function
'    End If
'  End If
  
  contrl_valid = True
End Function

Public Function CreateUserKey(Optional ByVal fDeleteDataUserKey As String = "",Optional ByVal fTableName As String = "COM_SelectList") As String
  Dim sql As String
' On Error Resume Next
  
  'Sets a random starting point for random number generation
  Randomize
  
  If Trim$(fDeleteDataUserKey) <> "" Then
    sql = "Delete From " & fTableName & " Where UserKey='" & Trim$(fDeleteDataUserKey) & "' and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    If DBExecuteSql(db, sql) = False Then Exit Function
  Else
    'Continue
  End If
  
  CreateUserKey = Format(ServerDate, "yyyyMMddHHmmss") & Left$(CStr(10 ^ 10 * Rnd), 10) & "[" & gUserID & "]"

End Function

'Public Function CreateUserKey() As String
'  On Error Resume Next
' 'Sets a random starting point for random number generation
'  Randomize
'  CreateUserKey = Format(ServerDate, '    "yyyyMMddHHmmss") & Left$(CStr(10 ^ 10 * Rnd), 10) & "[" & gUserID & "]"

'End Function

Public Function RemoveSpecialCharacters(ByVal InpStr As String) As String

  Try
  
  ' "%", "_", "[", "]", "^", "*", "?"
  
  InpStr = Replace(InpStr, "%", "-")
  InpStr = Replace(InpStr, "_", "-")
  InpStr = Replace(InpStr, "[", "-")
  InpStr = Replace(InpStr, "]", "-")
  InpStr = Replace(InpStr, "^", "-")
  InpStr = Replace(InpStr, "*", "-")
  InpStr = Replace(InpStr, "?", "-")
  InpStr = Replace(InpStr, "/", "-")
  InpStr = Replace(InpStr, "\", "-")
  InpStr = Replace(InpStr, ":", "-")
  InpStr = Replace(InpStr, "<", "-")
  InpStr = Replace(InpStr, ">", "-")
  InpStr = Replace(InpStr, "|", "-")
  InpStr = Replace(InpStr, Chr(13), " ")
  InpStr = Replace(InpStr, "  ", " ")
  InpStr = Replace(InpStr, vbLf, "")

  RemoveSpecialCharacters = InpStr
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function DBConnect() As Boolean
  Dim dname As String
  
  Try
  DBConnect = False
  
  dname = App.Path & "\ACE.ini"
  
  If Dir(dname) <> "" Then
    DBConnect = DBConnectThruIni
  Else
    DBConnect = DBConnectThruPrompt
  End If
  
Catch ex As Exception
  Call ErrMsgShow("Error - while trying to set up Database Connection.")
  Set db = Nothing

Finally

End Try
End Function

Public Function Dumrow(ByVal rownos As Integer,Optional ByVal ToFile As Boolean = False)
  ' function to print dummy rows while using the print method
  Dim rCnt As Integer

' On Error Resume Next
  
  If ToFile Then
    For rCnt = 1 To rownos
      sSendToPort
    Next rCnt
'    iLineNo = iLineNo + 1
  Else
    For rCnt = 1 To rownos
      Printer.Print
    Next rCnt
  End If
End Function

Public Sub ErrMsgShow(Optional ByVal errMsg As String = "",Optional ByVal ShowErrDescAlso As Boolean = True)
  Dim MsgStr As String
  
  If gErrNo = 0 Then
    gErrNo = Err.Number
    gErrDesc = Err.Description
  End If
  
  'If any Open Transaction found then it has to be RollBacked.
  If gTransactionCount > 0 Then
    If DBRollBackTransaction() = False Then
 Exit Sub
End If
  End If
  
  If Trim$(errMsg) <> "" Then
    MsgStr = Trim$(errMsg) & Chr(13)
    If ShowErrDescAlso = True Then
      MsgStr = MsgStr & "[" & gErrNo & "] " & gErrDesc
    End If
  Else
    MsgStr = "[" & gErrNo & "] " & gErrDesc
  End If
  
  MsgBox MsgStr, vbCritical, App.Title
  gErrNo = 0
End Sub

Public Function FAS_Posting_Validate_After_Save(ByVal VoucherNo As Integer,ByVal  VoucherDate As Date,ByVal  VoucherRoute As String,Optional ByVal PayMode As String = "None") As Boolean
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim fVTNo As Integer
  Dim fRoute As String
  Dim fVrDate As Date
  Dim rPos As Long
  Dim retVal As Variant
  
  Try
  FAS_Posting_Validate_After_Save = False
  
  'Proper Voucher No Validation
  If VoucherNo < 1 Then
    sShowMessage "Voucher No. not generated Properly."
    GoTo ExitHere
  End If
  
  fVTNo = VoucherNo
  fRoute = Trim$(VoucherRoute)
  fVrDate = VoucherDate
  
  'Voucher Date Validation
'  Sql = "Select DocDate from FAS_VoucherTranMast a"
'  Sql = Sql & " where a.docNo=" & fVTNo & " and a.VrRoute='" & fRoute & "' and a.CompCode='" & gFasCompCode & "'"
'  Set rs = New ADODB.Recordset
'  If RSOpen(rs, Sql) = False Then
'    GoTo ExitHere
'  End If
'
'  If rs.EOF Then
'    sShowMessage "Voucher Not Saved Properly."
'    GoTo ExitHere
'  Else
'    If IsNull(rs("DocDate")) Then
'      sShowMessage "Voucher Date Not Saved Properly."
'      GoTo ExitHere
'    End If
'  End If
  
'  If Format(rs("DocDate"), "yyyyMMdd") < Format(FinStart, "yyyyMMdd") Then
'    sShowMessage "Voucher Date is less than Financial year Start."
'    GoTo ExitHere
'  End If
  
  'Debit Credit Equality Validation
  sql = " Select Sum(CrAmt) as SubCrAmt, Sum(DrAmt) as SubDrAmt, isNull((Select Amount from FAS_VoucherTranMast"
  sql = sql & "  where docNo=a.docNo  and DocDate = a.DocDate and VrRoute=a.VrRoute and CompCode=a.CompCode), 0) as Amount"
  sql = sql & "  from FAS_VoucherTranSub a where a.docNo=" & fVTNo & " and a.DocDate='" & Format(fVrDate, "yyyy-MMM-dd") & "' and a.VrRoute='" & fRoute & "'"
  sql = sql & " and a.CompCode='" & gFasCompCode & "' group by a.docNo, a.DocDate, a.VrRoute, a.CompCode"
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    sShowMessage "Unable to Save Data Properly."
    GoTo ExitHere
  ElseIf rs("SubCrAmt") <> rs("SubDrAmt") Then
    sShowMessage "Debit & Credit Amount not Matched"
    GoTo ExitHere
  End If
  
  'All LrCode properly Updated Validation
  sql = "Select LrCode from FAS_VoucherTranSub a "
  sql = sql & " where a.DocNo=" & fVTNo & " and a.DocDate='" & Format(fVrDate, "yyyy-MMM-dd") & "' and a.VrRoute='" & fRoute & "' and LrCode<=0"
  sql = sql & " and a.CompCode='" & gFasCompCode & "'"
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    sShowMessage "Unable to save Accounts Name Properly in Detail Table."
    GoTo ExitHere
  End If
  
  'Cheque Entry Validation
  If PayMode = "Cheque" Then
    sql = "Select Sum(a.ChqAmount) as ChqAmt from FAS_VoucherTranPay a"
    sql = sql & " where a.DocNo=" & fVTNo & " and a.DocDate='" & Format(fVrDate, "yyyy-MMM-dd") & "' and a.VrRoute='" & fRoute & "' and a.CompCode='" & gFasCompCode & "'"
    If RSOpen(rs, sql) = False Then
      GoTo ExitHere
    End If

    If rs.EOF Then
      sShowMessage "Payment Details Data not saved Properly."
      GoTo ExitHere
    ElseIf rs("ChqAmt") <= 0 Then
      sShowMessage "Cheque Amount not saved Properly."
      GoTo ExitHere
    End If
  End If
  
  'Bill Details Validation
  sql = "Select * from"
  sql = sql & " (Select a.VTSNo, x.LedgerName, a.LrCode, (a.DrAmt-a.CrAmt) as VrAmt, isNull(Sum(b.Amount), 0) as BillAmt, x.isBillwise"
  sql = sql & " from FAS_VoucherTranSub a Left Outer Join FAS_BillDetails b"
  sql = sql & " On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.LrCode=b.LrCode and a.VTSNo=b.VTSNo and a.CompCode=b.CompCode"
  sql = sql & " Left Outer Join FAS_LedgerMaster x On a.LrCode=x.LedgerCode and a.CompCode=x.CompCode"
  sql = sql & " where a.docno=" & VoucherNo & " and a.DocDate='" & Format(fVrDate, "yyyy-MMM-dd") & "' and a.VrRoute='" & VoucherRoute & "' and x.isBillwise='1' and a.CompCode='" & gFasCompCode & "'"
  sql = sql & " Group by a.VTSNo, x.LedgerName, a.LrCode, (a.DrAmt-a.CrAmt), x.isBillwise) y"
  sql = sql & " Where y.VrAmt <> y.BillAmt"
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    sShowMessage "Voucher Amount not matched with Bill Amount for " & rs("LedgerName")
    GoTo ExitHere
  End If
  
  'Cost Centre Validation
  sql = "Select * from"
  sql = sql & " (Select x.LedgerName, a.LrCode, a.VTSNo, w.CategoryCode, Abs(a.DrAmt-a.CrAmt) as VrAmt,"
  sql = sql & " isNull(Sum(b.Amount), 0) as CostAmt, w.isMandatory"
  sql = sql & " from FAS_VoucherTranSub a Left Outer Join FAS_VoucherTranCostCentre b"
  sql = sql & " On a.DocNo=b.DocNo and a.DocDate=b.DocDate and a.VrRoute=b.VrRoute and a.LrCode=b.LrCode and a.VTSNo=b.VTSNo and a.CompCode=b.CompCode"
  sql = sql & " Left Outer Join FAS_CostCentre_Vw w"
  sql = sql & " On b.CostCentreCode=w.CostCentreCode and a.CompCode=w.CompCode"
  sql = sql & " Left Outer Join FAS_LedgerMaster x On a.LrCode=x.LedgerCode and a.CompCode=x.CompCode"
  sql = sql & " where a.DocNo=" & VoucherNo & " and a.DocDate='" & Format(fVrDate, "yyyy-MMM-dd") & "' and a.VrRoute='" & VoucherRoute & "' and x.isCostCentreApplicable='1' and a.CompCode='" & gFasCompCode & "'"
  sql = sql & " Group by x.LedgerName, a.LrCode, a.VTSNo, w.CategoryCode, Abs(a.DrAmt-a.CrAmt), w.isMandatory) y"
  sql = sql & " Where y.VrAmt <> y.CostAmt"
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    sShowMessage "Voucher Amount not matched with CostCentre Amount for " & rs("LedgerName")
    GoTo ExitHere
  End If
  
  FAS_Posting_Validate_After_Save = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function FetchValue_New(ByVal isMultiSelect As Boolean,ByVal  ConStr As String,ByVal  SqlStr As String,ByVal  ColVisStatus As String,ByVal  HelpCaption As String,ByVal  ParamArray FieldCaption()) As Boolean


  'isMultiSelect - True to allow multiselection
  'For SingleSelect mode getvalue will be a single dimension array
  'For multiselect mode getvalue will be a double dimension array, like array(rows, columns)
  'ConStr - db ConnectionString, SqlStr - Sql String to Select
  'ColVisStatus string like '10001110' 1-for visible, 0-for hidden
  'HelpCaption - Caption display on help form: FieldCaption is for only visible columns

  Dim lArray() As Variant

  Try
  FetchValue_New = False
  
  ReDim lArray(UBound(FieldCaption)) As Variant
  lArray = FieldCaption
  If (frmHelp.GetHelp(isMultiSelect, ConStr, SqlStr, ColVisStatus, CStr(HelpCaption), lArray)) Then

    If UBound(getvalue, 1) > 0 Then
      FetchValue_New = True
Exit Function
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function FixString(ByVal iVal As String,ByVal  ReqStrLen As Integer,ByVal  FillSingleChar As String) As String

  Dim str As String
  
  Try
  
  If ReqStrLen = 0 Then
    FixString = iVal
  ElseIf Len(iVal) >= ReqStrLen Then
    FixString = iVal
  Else
    str = String(ReqStrLen, Left$(FillSingleChar, 1))
    Mid(str, (ReqStrLen - Len(iVal) + 1)) = iVal
    FixString = str
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function
Public Function Floor(ByVal iNumber As String) As Long
    Floor = Int(Val(iNumber))
End Function

Public Function Form_Load_Routines() As Boolean
  Try
  Form_Load_Routines = False

  
  
  Form_Load_Routines = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function FormatDecimalString(ByVal NoOfDecimal As Integer) As String
' On Error Resume Next
  FormatDecimalString = "0"
  
  If NoOfDecimal > 0 Then
    FormatDecimalString = "0." & String(NoOfDecimal, "0")
  End If
End Function

Public Function FormExit(ByVal FormToExit As Form,ByVal  UnloadMode As Integer) As Boolean

  Try
  'Query UnloadMode Constants
  '0 vbFormControlMenu    User has chosen Close command from Control menu box on form.
  '1 vbFormCode           Unload method invoked from code.
  '2 vbAppWindows         Current Windows session ending.
  '3 vbAppTaskManager     Windows Task Manager is closing the application.
  '4 vbFormMDIForm        MDI child form is closing because the MDI form is closing
  '5 vbFormOwner          The owner of the form is closing
  '6                      User Forced to Close

  FormExit = False
  If UnloadMode = vbFormMDIForm Then
    Unload FormToExit
    Set FormToExit = Nothing
    FormExit = True
  Else
    If MsgBox("Do you want to Exit", vbApplicationModal + vbYesNo + vbDefaultButton1 + vbQuestion, App.Title) = vbYes Then


      Unload FormToExit
      Set FormToExit = Nothing
      FormExit = True
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function Get_ERPandTSNo(ByVal MenuFormName As String,ByVal  ModuleName As String,ByVal  sNo As eErpandTs) As String

  Dim rs As AceADODB.Recordset
  Dim sql As String
  
  Try

  sql = " Select * from dbo.fn_SAL_Menu_List('" & Trim$( MenuFormName) & "') where ModuleName='" & Trim$(ModuleName) & "'"

  Set rs = New Recordset
  If RSOpen(rs, sql) = False Then GoTo ExitHere
  
  If Not rs.EOF Then
    If sNo = mErp Then
      Get_ERPandTSNo = rs("ErpNo")
    Else
      Get_ERPandTSNo = rs("TsNo")
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function GetDocNo(ByVal aTableName As String,ByVal  aDocTypeColName As String,ByVal  aDocNoColName As String,ByVal  aDocType As String,Optional ByVal aPertainingTo As String = "",Optional ByVal aModuleName As String = "",Optional ByVal isDocTypeCommon As Boolean = False,Optional ByVal isBranchCommon As Boolean = False,Optional ByVal isCompanyCommon As Boolean = False,Optional ByVal isFinyearCommon As Boolean = True) As Variant






  Dim rs As AceADODB.Recordset
  Dim sql As String
  Dim Sql2 As String

  Try
  GetDocNo = 1
  
  sql = "Select Max(DocNo) as DocNo from ("
  sql = sql & " Select Max(StartingNo) - 1 as DocNo from COM_DocumentType_Master"

  sql = sql & " Where PertainingTo='" & Trim$(aPertainingTo) & "'"
  If Trim$(aModuleName) <> "" Then
    sql = sql & " and ModuleName='" & aModuleName & "'"
  End If
  sql = sql & "  and DocumentType='" & aDocType & "' and Not Exists (Select * from " & aTableName

  sql = sql & " where " & aDocTypeColName & "<>''"
  If isDocTypeCommon = False Then
    sql = sql & " and " & aDocTypeColName & "='" & aDocType & "'"
  End If

  Sql2 = ""
  If isBranchCommon = False Then
    Sql2 = Sql2 & " and branchID='" & gBranchID & "'"
  End If
  
  If isCompanyCommon = False Then
    Sql2 = Sql2 & " and compid = '" & gCompID & "'"
  End If
  
  If isFinyearCommon = True Then
    Sql2 = Sql2 & " and finyear='" & gFinyear & "'"
  End If
  
  sql = sql & Sql2 & ")"
  
  sql = sql & " Union All"
  sql = sql & " Select Max(convert(int," & aDocNoColName & ")) as DocNo from " & aTableName

  sql = sql & " where " & aDocTypeColName & "<>''"
  
  If isDocTypeCommon = False Then
    sql = sql & " and " & aDocTypeColName & "='" & aDocType & "'"
  End If
  
  If isFinyearCommon = True Then
    Sql2 = Sql2 & " and finyear='" & gFinyear & "'"
  End If
  
  sql = sql & Sql2
  
  sql = sql & " ) x Where DocNo is not null"
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    If Not IsNull(rs(0)) Then
      GetDocNo = rs(0) + 1
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

'Public Function SetEffStatus(sChar As String) As String
'  On Error Resume Next
'
'  Select Case sChar
'    Case "New"
'      SetEffStatus = "N"
'    Case "Modify"
'      SetEffStatus = "M"
'    Case "Authorised"
'      SetEffStatus = "A"
'    Case "Amendment"
'      SetEffStatus = "E"
'    Case "Hold"
'      SetEffStatus = "H"
'    Case "ShortClose"
'      SetEffStatus = "C"
'    Case "Active"
'      SetEffStatus = "Y"
'    Case "Inactive"
'      SetEffStatus = "I"
'    Case "Cancel"
'      SetEffStatus = "X"
'    Case "Delete"
'      SetEffStatus = "D"
'    Case Else
'      SetEffStatus = "Null"
'  End Select
'End Function

Public Function GetEffStatus(ByVal sChar As String) As String
' On Error Resume Next
  
  Select Case sChar
    Case "N"
      GetEffStatus = "New"
    Case "M"
      GetEffStatus = "Modify"
    Case "A"
      GetEffStatus = "Authorised"
    Case "E"
      GetEffStatus = "Amendment"
    Case "H"
      GetEffStatus = "Hold"
    Case "C"
      GetEffStatus = "ShortClose"
    Case "Y"
      GetEffStatus = "Active"
    Case "I"
      GetEffStatus = "Inactive"
    Case "X"
      GetEffStatus = "Cancel"
    Case "D"
      GetEffStatus = "Delete"
    Case Else
      sShowMessage "MisMatch Effective Status"
  End Select
  
End Function
Public Function GetItemNature(ByVal ItemTypeDesc As String) As String
  Dim tmp As String
  
  Select Case ItemTypeDesc
    Case "RAW MATERIAL"
      tmp = "R"
    Case "COMPONENT"
      tmp = "C"
    Case "FINISHED GOODS"
      tmp = "F"
    Case "SUB ASSEMBLY"
      tmp = "S"
    Case "CAPITAL"
      tmp = "P"
    Case "TOOL"
      tmp = "T"
    Case "SPARES"
      tmp = "A"
    Case "MISCELLANEOUS"
      tmp = "M"
    Case "SCRAP"
      tmp = "Z"
    Case "INSTRUMENT"
      tmp = "I"
    Case "CONSUMABLE"
      tmp = "L"
    Case "CUSTOMER SUPPLIED COMPONENT"
      tmp = "CC"
    Case "CUSTOMER SUPPLIED FINSHED GOODS"
      tmp = "CF"
    Case "CUSTOMER SUPPLIED RAW MATERIAL"
      tmp = "CR"
      
    Case Else
      tmp = ""
  End Select
  GetItemNature = tmp
End Function

Public Function GetItemNature1(ByVal ItemTypeDesc As String) As String
  Dim tmp As String
  
  Select Case ItemTypeDesc
    Case "RAW MATERIAL"
      tmp = "'R'"
    Case "COMPONENT"
      tmp = "'C'"
    Case "FINISHED GOODS"
      tmp = "'F'"
    Case "SUB ASSEMBLY"
      tmp = "'S'"
    Case "CAPITAL"
      tmp = "'P'"
    Case "TOOL"
      tmp = "'T'"
    Case "SPARES"
      tmp = "'A'"
    Case "MISCELLANEOUS"
      tmp = "'M'"
    Case "SCRAP"
      tmp = "'Z'"
    Case "INSTRUMENT"
      tmp = "'I'"
    Case "CONSUMABLE"
      tmp = "'L'"
    Case "CUSTOMER SUPPLIED"
      tmp = "'CC','CF','CR'"
    Case "CUSTOMER SUPPLIED FINSHED GOODS"
      tmp = "'CF'"
    Case "CUSTOMER SUPPLIED RAW MATERIAL"
      tmp = "'CR'"
    Case Else
      tmp = ""
  End Select
  GetItemNature1 = tmp
End Function

Public Function GetMax(ByVal TableName As String,ByVal  MaxFieldName As String,ByVal  ParamArray FldNameAndDataList()) As Long

  Dim rs As AceADODB.Recordset
  Dim sql As String
  Dim pCnt As Integer
  
  Try

  sql = "Select Max(Convert(int, " & Trim$(MaxFieldName) & ")) from " & Trim$( TableName) & " where "

  
  If Not IsMissing(FldNameAndDataList) Then
    For pCnt = 0 To UBound(FldNameAndDataList) Step 2
      sql = sql & FldNameAndDataList(pCnt) & "='" & FldNameAndDataList( pCnt + 1) & "' and "

    Next pCnt
'    Sql = Sql & " and"
  End If
  
  sql = sql & " BranchID='" & gBranchID & "' and CompID = '" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  If Not rs.EOF Then
    If Not IsNull(rs(0)) Then
      GetMax = Val(rs(0))
    Else
      GetMax = 0
    End If
  Else
    GetMax = 0
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function GetMaxRevNo(ByVal aTableName As String,ByVal  aDocTypeColName As String,ByVal  aDocType As String,Optional ByVal aRevColName As String = "RevisionNo",Optional ByVal isBranchCommon As Boolean = False,Optional ByVal isCompanyCommon As Boolean = False,Optional ByVal isFinyearCommon As Boolean = True) As Variant




  Dim rs As AceADODB.Recordset
  Dim sql As String

  Try
  
  sql = "select * from (Select Max(convert(int," & aRevColName & ")) as RevNo from " & aTableName

  sql = sql & " where " & aDocTypeColName & "='" & Trim$(aDocType) & "'"
  
  If isFinyearCommon = True Then
    sql = sql & " and finyear='" & gFinyear & "'"
  End If
  sql = sql & " ) x Where RevNo is not null"
  
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    If Not IsNull(rs(0)) Then
      GetMaxRevNo = rs(0) + 1
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function GetMaxVrNo(ByVal fDocDate As Date,ByVal  fVrRoute As String,Optional ByVal fVrNo As Integer = 0,Optional ByVal isMemorandum As Boolean = False) As Integer
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim MaxNo As String

  Try

  GetMaxVrNo = 0
  
'  If isMemorandum Then
'    Sql = "Select Max(DocNo) from FAS_MEM_VoucherTranMast"
'  Else
'    Sql = "Select Max(DocNo) from FAS_VoucherTranMast"
'  End If
'
'  Sql = Sql & " Where VrRoute='" & Trim$(fVrRoute) & "' and CompCode='" & gFasCompCode & "'"
'  Sql = Sql & " and DocDate <= '" & Format(fDocDate, "yyyy-MM-dd") & "' and DocNo<>" & fVrNo

  sql = "Select dbo.fn_FAS_GetMaxVoucherNo('" & Format(fDocDate, "yyyy-MMM-dd") & "', '" & Trim$(fVrRoute) & "', "
  sql = sql & fVrNo & ", '" & IIf(isMemorandum, 1, 0) & "', '" & gFasCompCode & "')"
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If

  If rs.EOF Or IsNull(rs(0)) Then
    MaxNo = "0"
  Else
    MaxNo = rs(0)
  End If

  GetMaxVrNo = Val(MaxNo)
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function
Public Function getMnuNo(ByVal Moduleinx As Long,ByVal  mnuInx As Long,ByVal  cnt As Long,ByVal  prvMnuCaption As String,ByVal  mnuCaption As String,ByVal  inx As Long,ByVal  ModuleName As String) As String


  Dim sql As String
  Dim snap As AceADODB.Recordset
  Dim sNo As Long
  Dim eNo As Long
  

  Try
  getMnuNo = ""

  Select Case mnuCaption
  
    Case "mnuMaster", "mnuOptions"
      getMnuNo = Moduleinx & ".1"
    Case "mnuTran", "mnuHelp"
      getMnuNo = Moduleinx & ".2"
    Case "mnuReports"
      getMnuNo = Moduleinx & ".3"
    Case "mnuNormCal"
      getMnuNo = Moduleinx & ".4"
    Case Else
      If inx = mnuInx Then
        sql = "Select SlNo,ERPNo,Menuindex from Sec_ModuleFormList where ModuleName='" & ModuleName & "'"

        sql = sql & " and SlNo=(Select Max(SlNo) from Sec_ModuleFormList where ModuleName='" & ModuleName & "')"

        Set snap = New AceADODB.Recordset
        If RSOpen(snap, sql) = False Then
          GoTo ExitHere
        End If
        If snap.EOF Then
          
        Else
          Select Case prvMnuCaption
            Case "mnuMaster", "mnuTran", "mnuReports", "mnuOptions"
              If inx = 0 Then
                getMnuNo = Trim$(snap("ERPNo")) & ".1"
              End If
            Case Else
              sNo = InStrRev(Trim$(snap("ERPNo")), ".", , vbTextCompare) + 1
              eNo = Mid$(Trim$(snap("ERPNo")), sNo, Val(Len(snap("ERPNo")) - (sNo - 1))) + 1

              getMnuNo = Left(Trim$(snap("ERPNo")), (sNo - 1)) & eNo
          End Select
        End If
      ElseIf inx < mnuInx Then
        sql = "Select SlNo,ERPNo,Menuindex from Sec_ModuleFormList where ModuleName='" & ModuleName & "'"

        sql = sql & " and SlNo=(Select Max(SlNo) from Sec_ModuleFormList where ModuleName='" & ModuleName & "')"

        Set snap = New AceADODB.Recordset
        If RSOpen(snap, sql) = False Then
          GoTo ExitHere
        End If
        If snap.EOF Then
          
        Else
          getMnuNo = Trim$(snap("ERPNo")) & ".1"
        End If
      Else
        sql = "Select SlNo,ERPNo,Menuindex from Sec_ModuleFormList where ModuleName='" & ModuleName & "'"

        sql = sql & " and Menuindex=" & mnuInx & " and SlNo=(Select Max(SlNo) from Sec_ModuleFormList where ModuleName='" & ModuleName & "' and Menuindex=" & mnuInx & " )"

        Set snap = New AceADODB.Recordset
        If RSOpen(snap, sql) = False Then
          GoTo ExitHere
        End If
          sNo = InStrRev(Trim$(snap("ERPNo")), ".", , vbTextCompare) + 1
          eNo = Mid$(Trim$(Trim$(snap("ERPNo"))), sNo, Val(Len(Trim$(snap("ERPNo"))) - (sNo - 1))) + 1

          getMnuNo = Left(Trim$(snap("ERPNo")), (sNo - 1)) & eNo
      End If
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function
Public Function GetOptionValue(ByVal OptionCode As eOptions,Optional ByVal MdlCode As Integer = 0) As String
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim CurDate As Date
  Dim rCnt As Integer
  
  Try
  GetOptionValue = ""
  
  CurDate = ServerDate
  
  sql = "Select OptionValue, OptionType"
  sql = sql & " from dbo.fn_COM_ERP_Options('" & Format(CurDate, "yyyy-MMM-dd HH:mm:ss") & "', '" & Format(CurDate, "yyyy-MMM-dd") & "', '" & gBranchID & "', '" & gCompID & "')"
  sql = sql & " Where OptionCode = " & Val(OptionCode)
  If MdlCode > 0 Then
    sql = sql & " And ModuleCode = " & Val(MdlCode) & ""
  End If
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If

  If Not rs.EOF Then
'    For rCnt = 1 To rs.RecordCount
      If rs("OptionType") = "DateTime" Then
        'To be Added
      Else
        GetOptionValue = rs("OptionValue")
      End If
'      rs.MoveNext
'    Next rCnt
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function GetOrderStatus(ByVal OrdStatus As String) As String
  Dim tmp As String
  
  Try

    Select Case OrdStatus
    Case "A"
      tmp = "Authorised"
    Case "H"
      tmp = "Hold"
    Case "N"
      tmp = "New"
    Case "M"
      tmp = "Modified"
    Case "S"
      tmp = "Short Close"
    Case "C"
      tmp = "Cancel"
  End Select

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

'Function to Get the Count of Search String in SearchIn String
'RecursiveSearch True return (Search 'aa' in 'aaa' results 2)
'RecursiveSearch False return (Search 'aa' in 'aaa' results 1)
Public Function GetStrCount(ByVal SearchString As String,ByVal  SearchIn As String,Optional ByVal RecursiveSearch As Boolean = False) As Integer

  Dim StrCnt As Integer
  Dim cnt As Integer
  Dim SrchLen As Integer
  Dim MaxPos As Integer
  
  Try

  cnt = 0

  If Len(SearchIn) = 0 Then
    GoTo ExitHere
  End If
  
  SrchLen = Len(SearchString)
  If Len(SrchLen) = 0 Then
    GoTo ExitHere
  End If
  
  MaxPos = Len(SearchIn) - SrchLen + 1
  StrCnt = 1
  Do While StrCnt <= MaxPos
    If SearchString = Mid$(SearchIn, StrCnt, SrchLen) Then
      cnt = cnt + 1
      If RecursiveSearch = True Then
        StrCnt = StrCnt + 1
      Else
        StrCnt = StrCnt + SrchLen
      End If
    Else
      StrCnt = StrCnt + 1
    End If
  Loop
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

'Function to Get the Position of Search String in SearchIn String
'It returns 1 index based array which contains the nth position in nth array
'RecursiveSearch True return (Search 'aa' in 'aaa' results 2)
'RecursiveSearch False return (Search 'aa' in 'aaa' results 1)
Public Function GetStrPositionArray(ByVal SearchString As String,ByVal  SearchIn As String,Optional ByVal RecursiveSearch As Boolean = False) As Integer()

  Dim StrCnt As Integer
  Dim cnt As Integer
  Dim SrchLen As Integer
  Dim MaxPos As Integer
  Dim StrPos() As Integer
  
  Try

  cnt = 0

  If Len(SearchIn) = 0 Then
    GoTo ExitHere
  End If
  
  SrchLen = Len(SearchString)
  If Len(SrchLen) = 0 Then
    GoTo ExitHere
  End If
  
  MaxPos = Len(SearchIn) - SrchLen + 1
  
  If MaxPos > 1 Then
    ReDim StrPos(1 To MaxPos) As Integer
    
    StrCnt = 1
    Do While StrCnt <= MaxPos
      If SearchString = Mid$(SearchIn, StrCnt, SrchLen) Then
        cnt = cnt + 1
        StrPos(cnt) = StrCnt
        If RecursiveSearch = True Then
          StrCnt = StrCnt + 1
        Else
          StrCnt = StrCnt + SrchLen
        End If
      Else
        StrCnt = StrCnt + 1
      End If
    Loop
    
    If cnt = 0 Then
      ReDim StrPos(1 To 1) As Integer
      StrPos(1) = 0
    Else
      ReDim Preserve StrPos(1 To cnt) As Integer
    End If
  Else
    ReDim StrPos(1 To 1) As Integer
    StrPos(1) = 0
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function
Public Function Load_Flow_Diagram(ByVal DiaGramCode As Long) As String
  Select Case DiaGramCode

    Case 1
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Operation Or Machining.JPG"

    Case 2
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "TransPortTation.JPG"
    Case 3
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Operation With Automatic Inspection.JPG"

    Case 4
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Operation With Multiple Product Streams.JPG"

    Case 5
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Inspection.JPG"
    Case 6
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Visual Inspection.JPG"
    Case 7
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Receipt of Material.JPG"

    Case 8
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Delay.JPG"
    Case 9
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Decision.JPG"
    Case 10
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Operator.JPG"
    Case 11
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Partial Operator.JPG"
    Case 12
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Primary Path.JPG"
    Case 13
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Secondary Path.JPG"
    Case 14
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Storage.JPG"
    Case 15
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Critical Characteristic.JPG"

    Case 16
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "Significant Characteristic.JPG"

    Case 17
      Load_Flow_Diagram = App.Path & "\Process Flow Report Icon\" & "PACKING.JPG"
  End Select

End Function

Public Function Load_Symbols(ByVal DiaGramCode As Long) As String
  Select Case DiaGramCode
    
    Case 0
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Empty.JPG"
    Case 1
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Operation Or Machining.JPG"
    Case 2
      Load_Symbols = App.Path & "\Process Flow Icon\" & "TransPortTation.JPG"
    Case 3
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Operation With Automatic Inspection.JPG"

    Case 4
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Operation With Multiple Product Streams.JPG"

    Case 5
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Inspection.JPG"
    Case 6
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Visual Inspection.JPG"
    Case 7
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Receipt of Material.JPG"
    Case 8
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Delay.JPG"
    Case 9
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Decision.JPG"
    Case 10
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Operator.JPG"
    Case 11
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Partial Operator.JPG"
    Case 12
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Primary Path.JPG"
    Case 13
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Secondary Path.JPG"
    Case 14
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Storage.JPG"
    Case 15
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Critical Characteristic.JPG"
    Case 16
      Load_Symbols = App.Path & "\Process Flow Icon\" & "Significant Characteristic.JPG"

    Case 17
      Load_Symbols = App.Path & "\Process Flow Icon\" & "PACKING.JPG"
  End Select

End Function

Public Sub LoginSecurity()
  Dim i As Integer
  Dim j As Integer
  Dim cmdLineArg As String
  
  cmdLineArg = Trim$(Command)
  
' On Error Resume Next
  If Len(cmdLineArg) = 0 Then
    MsgBox "You can run this only from ACE Profit", vbInformation + vbOKOnly, App.Title

    End
    Exit Sub
  End If

  i = InStr(1, cmdLineArg, "-")
  gBranchID = Mid$(cmdLineArg, 3, 2)
  gCompID = Left$(cmdLineArg, 2)
  
  j = InStr(i + 1, cmdLineArg, "-")
  gUserID = UCase(Mid$(cmdLineArg, i + 1, (Len(cmdLineArg) - 21)))
  gFinyear = Right$(cmdLineArg, 15)
  
End Sub

Public Function MaskSingleQuotes(ByVal InpStr As String) As String
  Dim tmp As String
  Dim Pos As Integer

  Try
  tmp = InpStr
  
  Pos = InStr(1, tmp, "'", vbTextCompare)
  Do While Pos > 0
    tmp = Left$(tmp, Pos) & "'" & Mid$(tmp, Pos + 1)
    
    Pos = InStr(Pos + 2, tmp, "'", vbTextCompare)
  Loop
  
  MaskSingleQuotes = tmp
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function MDIFormExit(ByVal FormToExit As MDIForm) As Boolean

  Try

  MDIFormExit = False

  If MsgBox("Do you want to Exit ACE Profit", vbApplicationModal + vbYesNo + vbDefaultButton1 + vbQuestion, App.Title) = vbYes Then


    Set db = Nothing
    Unload FormToExit
    Set FormToExit = Nothing
    MDIFormExit = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Sub ModuleFormListFillDB(ByVal MenuForm As MDIForm,ByVal  ModuleName As String)
  Dim MenuCtl As Control
  Dim cnt As Long
  Dim ERPNo As String
  Dim sql As String
  Dim CurTopMenu As String
  Dim snap As AceADODB.Recordset
  Dim Moduleinx As Long
  Dim previnx As Long
  Dim prvMnuCaption As String
  
  Try
  
  'This is to verify whether the Menu Controls are changed after last save
  'It only checks the No. of Menu Items.  So it is not reliable method
  cnt = 0
  For Each MenuCtl In MenuForm.Controls
    If TypeOf MenuCtl Is Menu Then
      Select Case MenuCtl.Name
        Case "mnuFile", "mnuModules", "mnuHelp"
          CurTopMenu = MenuCtl.Name     'No Need to Add
        Case Else
          Select Case MenuCtl.Name
            Case "mnuMaster", "mnuTran", "mnuReports", "mnuListing", "mnuOptions"

              CurTopMenu = MenuCtl.Name 'Must be taken for store
          End Select
          If CurTopMenu <> "mnuModules" Then    'Security fix by another procedure
            If CurTopMenu <> "mnuFile" Then     'Full rights given - so don't take
              If CurTopMenu <> "mnuHelp" Then   'Full rights given - so don't take
                If MenuCtl.Visible = True Then
                  cnt = cnt + 1       'Menu to be added. So SlNo incremented
                End If
              End If
            End If
          End If
      End Select
    End If
  Next
  
  sql = "Select Count(*) from Sec_ModuleFormList where ModuleName='" & ModuleName & "'"

  Set snap = New AceADODB.Recordset
  If RSOpen(snap, sql) = False Then
    GoTo ExitHere
  End If
  
  If cnt = snap(0) Then   'No Changes made in Menu Structure after last save
    GoTo ExitHere
  End If
  
  sql = "Select ERPNo from Sec_ModuleList where ModuleName='" & ModuleName & "'"
  If RSOpen(snap, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not snap.EOF Then
    Moduleinx = snap("ERPNo")
  End If
  
  'this is to insert the changed Menu Controls
  If DBStartTransaction = False Then
 GoTo ExitHere
End If

  'First delete the Old Menu Structure Data
  sql = "Delete from Sec_ModuleFormList where ModuleName='" & ModuleName & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
  'Add Current Menu Structure data
  cnt = 0
previnx = 0
prvMnuCaption = ""
  For Each MenuCtl In MenuForm.Controls
    
    If TypeOf MenuCtl Is Menu Then
      Select Case MenuCtl.Name
        'as per Guidelines these are the top level menu names
        Case "mnuFile", "mnuModules", "mnuHelp"
          CurTopMenu = MenuCtl.Name     'No Need to Add
        Case Else
          Select Case MenuCtl.Name
            Case "mnuMaster", "mnuTran", "mnuReports", "mnuListing", "mnuOptions"

              CurTopMenu = MenuCtl.Name 'Must be taken for store
          End Select
          If CurTopMenu <> "mnuModules" Then    'Security fix by another procedure
            If CurTopMenu <> "mnuFile" Then     'Full rights given - so don't take
              If CurTopMenu <> "mnuHelp" Then   'Full rights given - so don't take
                If MenuCtl.Visible = True Then
                  cnt = cnt + 1       'Menu to be added. So SlNo incremented
                  ERPNo = getMnuNo(Moduleinx, Val(MenuCtl.Index), cnt, prvMnuCaption, MenuCtl.Name, previnx, ModuleName)

                  sql = "Execute Sec_Ins_ModuleFormList "
                  sql = sql & " @ERPNo='" & Trim$(ERPNo) & "',"
                  sql = sql & " @TSNo='" & Trim$(cnt) & "',"
                  sql = sql & " @ModuleName='" & ModuleName & "',"
                  sql = sql & " @MenuCaption='" & MenuCtl.Caption & "',"
                  sql = sql & " @MenuName='" & MenuCtl.Name & "',"
                  sql = sql & " @MenuIndex=" & Val(MenuCtl.Index) & ","
                  sql = sql & " @SlNo=" & cnt
                  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
                  previnx = Val(MenuCtl.Index)
                  prvMnuCaption = MenuCtl.Name
                End If
              End If
            End If
          End If
      End Select
    End If
  Next
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
Catch ex As Exception
  If Err.Number = 343 Then    'No index provided for menu
    Call ErrMsgShow( "In MdiForm Menu Design all Menu should have the index atleast 0")

  Else
    Call ErrMsgShow(ex)
  End If

Finally

  Set MenuCtl = Nothing
End Try
End Sub

Public Sub ModuleMenuLoad(ByVal MenuForm As MDIForm)
  'Dependency: Public String Array Variable MoudleList()
  Dim rs As AceADODB.Recordset
  Dim sql As String
  Dim cnt As Integer
  
  Try

  sql = "Select ERPNo,TSNO,Caption, ModuleCode, ModuleName,"
  sql = sql & " (ExePath + (Case When Len(ExePath)> 0 then '\' Else '" & App.Path & "\' End) + ExeName + '.exe ') as Prog"

  sql = sql & " from Sec_ModuleList Order by convert(int,ERPNO)"
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  ReDim ModuleList(2, 100) As String
  cnt = 0
  If Not rs.EOF Then
    MenuForm.mnuModuleName(cnt).Caption = Format(rs("ERPNo"), "00") & " " & rs("Caption")

    MenuForm.mnuModuleName(cnt).Tag = rs("TSNO")
    ModuleList(1, cnt) = rs("ModuleName")
    ModuleList(2, cnt) = rs("Prog")
    rs.MoveNext
    Do While Not rs.EOF
      cnt = cnt + 1
      Load MenuForm.mnuModuleName(cnt)
      MenuForm.mnuModuleName(cnt).Caption = Format(rs("ERPNo"), "00") & " " & rs("Caption")

      MenuForm.mnuModuleName(cnt).Tag = rs("TSNO")
      MenuForm.mnuModuleName(cnt).Visible = True
      ModuleList(1, cnt) = rs("ModuleName")
      ModuleList(2, cnt) = rs("Prog")
      rs.MoveNext
    Loop
    ReDim Preserve ModuleList(2, cnt) As String
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub
Public Sub ModuleSecurity(ByVal MenuForm As MDIForm,ByVal  ThisModuleName As String)
  Dim rsUser As AceADODB.Recordset
  Dim sql As String
  Dim sUser As String
  
  Dim cnt As Integer
  Dim CurDate As Date
  Dim MenuCtl As Control

  Try
  
  CurDate = ServerDate
  
  sql = " Select UserID, UserGroupID "
  sql = sql & " From dbo.fn_SEC_Users ( '" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format( CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "')"


  sql = sql & " Where UserID = '" & gUserID & "'"
  Set rsUser = New AceADODB.Recordset
  If RSOpen(rsUser, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsUser.EOF Then
    sUser = Trim$(rsUser("UserGroupID"))
    If rsUser("UserGroupID") = "A" Or rsUser("UserGroupID") = "D" Then
      Call FormMenuSecurity(MenuForm, ThisModuleName, sUser)
      For Each MenuCtl In MenuForm.Controls
        If TypeOf MenuCtl Is Menu Then
          If MenuCtl.Name = "mnuModuleName" Then
              If ModuleList(1, MenuCtl.Index) = ThisModuleName Then
                MenuForm.mnuModuleName(MenuCtl.Index).Enabled = False
              Else
                MenuForm.mnuModuleName(MenuCtl.Index).Enabled = True
              End If
          Else
            MenuCtl.Enabled = True
          End If
        End If
      Next
      
    Else
      'Module Security done here
      sql = "select ModuleName, Count(*) as TotalCount From Sec_UserRights"
      sql = sql & " where UserID='" & gUserID & "' and BranchID='" & gBranchID & "' and CompId='" & gCompID & "' and finyear ='" & gFinyear & "'"

      sql = sql & " Group by ModuleName Having Count(*) > 0"
      If RSOpen(rsUser, sql) = False Then
        GoTo ExitHere
      End If
      
      If Not rsUser.EOF Then
        For cnt = 0 To MenuForm.mnuModuleName.UBound
          If ModuleList(1, cnt) <> ThisModuleName Then
            rsUser.Filter = adFilterNone
            rsUser.Filter = "ModuleName='" & ModuleList(1, cnt) & "'"
            If Not rsUser.EOF Then
              MenuForm.mnuModuleName(cnt).Enabled = True
            Else
              MenuForm.mnuModuleName(cnt).Enabled = False
            End If
          End If
        Next cnt
      End If
      'for other forms Verify the user rights & enable it
      Call FormMenuSecurity(MenuForm, ThisModuleName, sUser)
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsUser = Nothing
  Set MenuCtl = Nothing
End Try
End Sub

Public Sub QMS_ModuleSecurity(ByVal MenuForm As MDIForm,ByVal  ThisModuleName As String)
  Dim rsUser As AceADODB.Recordset
  Dim sql As String
  Dim sUser As String
  
  Dim cnt As Integer
  Dim CurDate As Date
  Dim MenuCtl As Control

  Try
  
  CurDate = ServerDate
  
  sql = " Select UserID, UserGroupID "
  sql = sql & " From dbo.fn_SEC_Users ( '" & Format(CurDate, "yyyy-MM-dd 23:59:59") & "', '" & Format( CurDate, "yyyy-MM-dd") & "', '" & gBranchID & "', '" & gCompID & "')"


  sql = sql & " Where UserID = '" & gUserID & "'"
  
  Set rsUser = New AceADODB.Recordset
  If RSOpen(rsUser, sql) = False Then GoTo ExitHere
  
  If Not rsUser.EOF Then
    sUser = Trim$(rsUser("UserGroupID"))
    
    If rsUser("UserGroupID") = "A" Or rsUser("UserGroupID") = "D" Then
      Call QMS_FormMenuSecurity(MenuForm, ThisModuleName, sUser)
      
      For Each MenuCtl In MenuForm.Controls
        If TypeOf MenuCtl Is Menu Then
          If MenuCtl.Name = "mnuModuleName" Then
              If ModuleList(1, MenuCtl.Index) = ThisModuleName Then
                MenuForm.mnuModuleName(MenuCtl.Index).Enabled = False
              Else
                MenuForm.mnuModuleName(MenuCtl.Index).Enabled = True
              End If
          Else
            MenuCtl.Enabled = True
          End If
        End If
      Next
      
    Else
      'Module Security done here
      sql = "select ModuleName, Count(*) as TotalCount From Sec_UserRights"
      sql = sql & " where UserID='" & gUserID & "' and BranchID='" & gBranchID & "' and CompId='" & gCompID & "' and finyear ='" & gFinyear & "'"

      sql = sql & " Group by ModuleName Having Count(*) > 0"
      
      If RSOpen(rsUser, sql) = False Then GoTo ExitHere
      
      If Not rsUser.EOF Then
        For cnt = 0 To MenuForm.mnuModuleName.UBound
        
          If ModuleList(1, cnt) <> ThisModuleName Then
            rsUser.Filter = adFilterNone
            rsUser.Filter = "ModuleName='" & ModuleList(1, cnt) & "'"
            
            If Not rsUser.EOF Then
              MenuForm.mnuModuleName(cnt).Enabled = True
            Else
              MenuForm.mnuModuleName(cnt).Enabled = False
            End If
          Else
            'Continue
          End If
        Next cnt
      End If
      'for other forms Verify the user rights & enable it
      Call QMS_FormMenuSecurity(MenuForm, ThisModuleName, sUser)
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsUser = Nothing
  Set MenuCtl = Nothing
End Try
End Sub

'Function to Get nth String like st for 1st, nd for 2nd, rd for 3rd and th for else
Public Function nthString(ByVal InputNo As Integer) As String
  Dim tmp As String
  
' On Error Resume Next
  nthString = ""
  
  Select Case InputNo
    Case 1
      tmp = "st"
    Case 2
      tmp = "nd"
    Case 3
      tmp = "rd"
    Case Else
      tmp = "th"
  End Select
  
  nthString = tmp
  
End Function
Public Function NumValFloat(ByVal TxtBox As TextBox,ByVal  KeyAscii As Integer,ByVal  Max As Double,ByVal  decimals As Integer,Optional ByVal Min As Integer) As Integer

Dim    fnd As Variant
Dim  Temp As Variant

  Temp = InStr(TxtBox, ".")
  If KeyAscii = 8 Then NumValFloat = KeyAscii
Exit Function

  If Not ( KeyAscii > 45 And KeyAscii < 58 Or KeyAscii = 8) Then NumValFloat = 0
Exit Function


  If KeyAscii = 46 Then fnd = InStr(TxtBox, ".")
If fnd <> 0 Then NumValFloat = 0
Exit Function

  If Temp <> 0 Then If Len(Mid(TxtBox.Text, Temp, Len(TxtBox.Text))) > decimals Then NumValFloat = 0
Exit Function

  If TxtBox.SelStart = 0 Then If Not (Val(TxtBox.Text & Chr(KeyAscii)) >= Val( Min)) Then NumValFloat = 0
Exit Function

  If TxtBox.SelStart = 1 Then If Mid(TxtBox.Text, 1, 1) = 0 Then NumValFloat = 0
Exit Function

  If TxtBox.SelStart >= 1 Then If Not (Val(TxtBox.Text & Chr(KeyAscii)) >= Val( Min) And Val(TxtBox.Text & Chr(KeyAscii)) <= Val( Max)) Then NumValFloat = 0
Exit Function


  NumValFloat = KeyAscii
End Function

Public Function NumValInt(ByVal TxtBox As TextBox,ByVal  KeyAscii As Integer,ByVal  MaxNo As Integer,Optional ByVal MinNo As Integer) As Integer

  
' On Error Resume Next
  NumValInt = 0
  
  Select Case KeyAscii
    Case vbKeyBack, vbKeyDelete ' vbKey0 To vbKey9
      'Allow for Further Process
      NumValInt = KeyAscii
      Exit Function
    Case Else
'      Exit Function
  End Select
  
'  Select Case TxtBox.SelStart
'    Case 0
'      If Val(TxtBox.Text & Chr(KeyAscii)) < MinNo Or Val(TxtBox.Text & Chr(KeyAscii)) > Val(MaxNo) Then
'        Exit Function
'      End If
'    Case 1
'      If Mid(TxtBox.Text, 1, 1) = 0 Then
'        Exit Function
'      End If
'    Case Else
      If Not (Val(TxtBox.Text & Chr(KeyAscii)) >= Val(MinNo) And Val( TxtBox.Text & Chr(KeyAscii)) <= Val(MaxNo)) Then

        Exit Function
      End If
'  End Select
  
  NumValInt = KeyAscii
  
'  If KeyAscii = 8 Then NumValInt = KeyAscii: Exit Function

'  If Not (KeyAscii > 46 And KeyAscii < 58 Or KeyAscii = 8) Then NumValInt = 0: Exit Function

'  If TxtBox.SelStart = 0 Then If Val(TxtBox.Text & Chr(KeyAscii)) < MinNo Or Val(TxtBox.Text & Chr(KeyAscii)) > Val(MaxNo) Then NumValInt = 0: Exit Function

'  If TxtBox.SelStart = 1 Then If Mid(TxtBox.Text, 1, 1) = 0 Then NumValInt = 0: Exit Function

'  If TxtBox.SelStart >= 1 Then If Not (Val(TxtBox.Text & Chr(KeyAscii)) >= Val(MinNo) And Val(TxtBox.Text & Chr(KeyAscii)) <= Val(MaxNo)) Then NumValInt = 0: Exit Function
'  NumValInt = KeyAscii
End Function

'Public Function NumValKeyDownPress(fTxtBox As TextBox, fKeyCode As Integer, Optional isKeyPress As Boolean = True, Optional isCtlKeyPressed As Integer = 0, Optional MaxVal As Variant, Optional MinVal As Variant, Optional DecPlaces As Variant) As Integer
'  Dim TextAfter As String
'  Dim DecPos As Integer
'  Dim isDecNoLimit As Boolean
'  Dim NoOfDec As Integer
'
'  On Error GoTo ErrHandler
'  NumValKeyDownPress = 0
'
'  With fTxtBox
'    Select Case fKeyCode
'      Case vbKey0 To vbKey9, vbKeyDecimal, vbKeySubtract, vbKeyDelete, vbKeyBack, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd
'        'Proceed
'      Case vbKeyNumpad0 To vbKeyNumpad9, 189, 190
'        If isKeyPress = True Then
'          GoTo ExitHere
'        Else
'          If isKeyPress = False Then
'            If isCtlKeyPressed > 0 Then
'              GoTo ExitHere
'            End If
'          End If
'        End If
'      Case Else
'        GoTo ExitHere
'    End Select
'
'    Select Case fKeyCode
'      Case vbKey0 To vbKey9
'        TextAfter = Left$(.Text, .SelStart) & Chr(fKeyCode) & Mid$(.Text, (.SelStart + .SelLength + 1))
'      Case vbKeyNumpad0 To vbKeyNumpad9
'        TextAfter = Left$(.Text, .SelStart) & Chr(fKeyCode - 48) & Mid$(.Text, (.SelStart + .SelLength + 1))
'      Case vbKeySubtract, 189   'Subtract sign non numpad key
'        TextAfter = Left$(.Text, .SelStart) & Chr(fKeyCode - 48) & Mid$(.Text, (.SelStart + .SelLength + 1))
'      Case vbKeyDecimal, 190    'Decimal point not in Numerical Pad
'        TextAfter = Left$(.Text, .SelStart) & "." & Mid$(.Text, (.SelStart + .SelLength + 1))
'      Case vbKeyDelete
'        If Len(.Text) > 0 Then
'          If .SelLength > 0 Then
'            TextAfter = Left$(.Text, .SelStart) & Mid$(.Text, (.SelStart + .SelLength + 1))
'          Else
'            TextAfter = Left$(.Text, .SelStart) & Mid$(.Text, (.SelStart + 2))
'          End If
'        End If
'      Case vbKeyBack
'        If Len(.Text) > 0 Then
'          If .SelLength > 0 Then
'            TextAfter = Left$(.Text, .SelStart) & Mid$(.Text, (.SelStart + .SelLength + 1))
'          Else
'            TextAfter = Left$(.Text, .SelStart - 1) & Mid$(.Text, (.SelStart + 1))
'          End If
'        End If
'      Case Else
'        TextAfter = .Text
'    End Select
'
'    If IsMissing(DecPlaces) Then
'      isDecNoLimit = True
'    Else
'      isDecNoLimit = False
'
'      If Val(DecPlaces) < 0 Then
'        NoOfDec = 0
'      Else
'        NoOfDec = Val(DecPlaces)
'      End If
'    End If
'
'    DecPos = InStr(1, Trim$(TextAfter), ".", vbTextCompare)
'
'    If isDecNoLimit = False Then
'      'For integer decimal point not allowed
'      If NoOfDec = 0 Then
'        If DecPos > 0 Then
'          GoTo ExitHere
'        End If
'      End If
'
'      'Length of decimal places should not exceed if it is a constraint
'      If DecPos > 0 Then
'        If Len(Mid$(TextAfter, DecPos + 1)) > DecPlaces Then
'          GoTo ExitHere
'        End If
'      End If
'    End If
'
'    'two decimal points not allowed
'    If InStr(DecPos + 1, Trim$(TextAfter), ".", vbTextCompare) > 0 And NoOfDec > 0 Then
'      GoTo ExitHere
'    End If
'
'    'Subtract sign should be the first character
'    If InStr(1, TextAfter, "-", vbTextCompare) > 1 Then
'      GoTo ExitHere
'    End If
'
'    'First digit Zero should be followed by decimal point
'    If Left$(TextAfter, 1) = "0" And Mid$(TextAfter, 2, 1) <> "." And Len(TextAfter) > 1 Then
'      GoTo ExitHere
'    End If
'
'    'Text after value should be less than the Maximum Value
''    If Not IsMissing(MaxVal) Then
''      If IsNumeric(MaxVal) Then
''        If CheckMaxMinLimits(TextAfter, , Val(MaxVal)) = False Then
''          GoTo ExitHere
''        End If
''      End If
''    End If
'  End With
'
'  NumValKeyDownPress = fKeyCode
'
'ExitHere:
'  On Error GoTo 0
'  Exit Function
'ErrHandler:
'  Resume ExitHere
'End Function

Public Function RoundingOff(ByVal amt As Double,ByVal  RoundingValue As Double,Optional ByVal NoOfDecimals As Integer = 2,Optional ByVal Direction As eRoundOffOptions = eRoundOffOptions.anNone) As Double


  'S.Anand on 04/01/2005
  'To round off the given amount to the slab of rounding value with direction options
  'Associated Procedure: 1) RoundingOffOptions - Enum
  
  Dim NoOfBlocks As Long
  Dim tmpAmt As Double
'  Dim NoOfDec As Integer
  Dim tmpStr As String

  Try
  
  If Direction = anNone Then
    RoundingOff = amt
  Else
    RoundingOff = 0
    NoOfBlocks = Int(amt / RoundingValue)
    
    tmpAmt = NoOfBlocks * RoundingValue
    
    If tmpAmt <> amt Then
      tmpStr = RoundingValue / 2#
  '    NoOfDec = InStr(1, tmpStr, ".", vbTextCompare)
      
  '    tmpStr = Mid$(tmpStr, NoOfDec + 1, Len(tmpStr) - NoOfDec)
      
      Select Case Direction
        Case anNearest
  '        If Round((amt - tmpAmt), Len(tmpStr)) >= (RoundingValue / 2#) Then
          If Round((amt - tmpAmt), NoOfDecimals) >= (RoundingValue / 2#) Then
            tmpAmt = (NoOfBlocks + 1) * RoundingValue
          Else
            'tmpAmt is already in downward direction
          End If
        Case anUpward
          tmpAmt = (NoOfBlocks + 1) * RoundingValue
        Case anDownward
          'tmpAmt is already in downward direction
        Case Else
          tmpAmt = amt
      End Select
    End If
    
    RoundingOff = tmpAmt
  End If
Catch ex As Exception
  Call ErrMsgShow("Error Occured while rounding the value.")

Finally

End Try
End Function

Public Function RptStandard_Header(ByVal sMaxCol As Long,ByVal  StartingRow As Long,ByVal  ERPNo As String,ByVal  TSNo As String,Optional ByVal PrintDate As Boolean = True,Optional ByVal sERPno As Boolean = True)


  Dim sColLeft As Long
  Dim sColRight As Long
  
    frmPrint.sprPrint.BlockMode = False
    'Set Page Number
    frmPrint.sprPrint.MaxCols = sMaxCol
    
    sColLeft = sMaxCol / 2
    sColRight = sMaxCol - sColLeft
    
    frmPrint.sprPrint.AddCellSpan 1, StartingRow, sColLeft, 1
      
    frmPrint.sprPrint.AddCellSpan sColLeft + 1, StartingRow, sColRight, 1

    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = StartingRow

    frmPrint.sprPrint.SetText 1, StartingRow, "Page Number :" + str( iPageNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter


    
    'Set Print Date
    If PrintDate = True Then
      frmPrint.sprPrint.Col = sColLeft + 1
frmPrint.sprPrint.Row = StartingRow

      frmPrint.sprPrint.SetText sColLeft + 1, StartingRow, "Printed Date : " & Format(ServerDate, "dd-MM-yyyy HH:MM")

      frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter

    End If
    StartingRow = StartingRow + 1
    
    'Set Erp & Ts No
    If sERPno = True Then
     ' If Trim$(ERPNo) = "" Then
     '   sShowMessage "ErpNo is Not Found"
     ' End If
      
      frmPrint.sprPrint.AddCellSpan 1, StartingRow, sColLeft, 1
      frmPrint.sprPrint.AddCellSpan sColLeft + 1, StartingRow, sColRight, 1
      
      frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = StartingRow
frmPrint.sprPrint.FontUnderline = True

      frmPrint.sprPrint.SetText 1, StartingRow, "ERP Ref. No   :" & Space(1) & Trim$( ERPNo)
frmPrint.sprPrint.TypeHAlign = TypeHAlignLeft
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter


      frmPrint.sprPrint.Col = sColLeft + 1
frmPrint.sprPrint.Row = StartingRow
frmPrint.sprPrint.FontUnderline = True

      frmPrint.sprPrint.SetText sColLeft + 1, StartingRow, "TS Ref. No" & Chr(32) & Chr(32) & " : " & TSNo & String(Val(14 - Len( TSNo)), "_")
frmPrint.sprPrint.TypeHAlign = TypeHAlignRight
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter



    End If
    StartingRow = StartingRow + 1
    'Set Company Name
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = StartingRow
    frmPrint.sprPrint.AddCellSpan 1, StartingRow, sMaxCol, 1
    frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 12

    frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
    frmPrint.sprPrint.SetText 1, StartingRow, gCompName
    StartingRow = StartingRow + 2
    
    'Set ReportHeading
    frmPrint.sprPrint.Col = 1
frmPrint.sprPrint.Row = StartingRow
    frmPrint.sprPrint.AddCellSpan 1, StartingRow, sMaxCol, 1
    frmPrint.sprPrint.FontBold = True
frmPrint.sprPrint.FontSize = 10

    frmPrint.sprPrint.TypeHAlign = TypeHAlignCenter
frmPrint.sprPrint.TypeVAlign = TypeVAlignCenter
    frmPrint.sprPrint.SetText 1, StartingRow, gStrReportHeading
    StartingRow = StartingRow + 2
  
End Function

'Public Function RupeesInWords(Number As Double, Optional rs As String, '  Optional ps As String) As String

'  Dim IntPart As String
'  Dim tmpStr As String
'  Dim tmpStr2 As String
'  Dim OpStr As String
'  Dim Cnt As Integer
'
'  On Error GoTo ErrHandler
'
'  IntPart = Int(Number / 1#)
'
'  OpStr = ""
'
'  If Val(IntPart) > 0 Then
'    tmpStr = IntPart
'    Do
'      If Trim$(OpStr) <> "" Then
'        OpStr = "Crores " & OpStr
'      End If
'
'      'Last 7 digits of remaining string
'      tmpStr2 = Right$(tmpStr, 7)
'
'      'To insert a zero before hundreth place digit ex: 4562 to 40562
'      If Len(tmpStr2) >= 3 Then
'        tmpStr2 = Left$(tmpStr2, Len(tmpStr2) - 3) & "0" & Right$(tmpStr2, 3)
'      End If
'
'      'To make it even no. digits
'      If Len(tmpStr2) Mod 2 = 1 Then
'        tmpStr2 = "0" & tmpStr2
'      End If
'
'      For Cnt = 1 To Round((Len(tmpStr2) / 2) + 0.01)
'        Select Case Cnt
'          Case 1    'Units
'            OpStr = TwoDigitNumberToWord(Val(Right$(tmpStr2, 2))) & " " & OpStr
'          Case 2    'Hundreds
'            OpStr = TwoDigitNumberToWord(Val(Mid$(tmpStr2, Len(tmpStr2) - 3, '              2))) & " Hundred " & IIf(Len(Trim$(OpStr)) > 0, "and ", " ") & OpStr

'          Case 3    'Thousands
'            OpStr = TwoDigitNumberToWord(Val(Mid$(tmpStr2, Len(tmpStr2) - 5, '              2))) & " Thousand " & OpStr

'          Case 4    'Lakhs
'            OpStr = TwoDigitNumberToWord(Val(Mid$(tmpStr2, Len(tmpStr2) - 7, '              2))) & " Lakhs " & OpStr

'        End Select
'      Next Cnt
'
'      If Len(tmpStr) > 7 Then
'        tmpStr = Left$(tmpStr, Len(tmpStr) - 7)
'      Else
'        tmpStr = ""
'      End If
'    Loop While Len(Trim$(tmpStr)) > 0
'
'    'Add Rupees String
'    OpStr = rs & " " & OpStr
'  End If
'
'  'For Paise Conversion and Paise String Add
'  If Val(IntPart) <> Number Then
'    OpStr = OpStr & IIf(Len(Trim$(OpStr)) > 0, " and ", '      "") & ps & " " & TwoDigitNumberToWord(Val(Format((Number - Val(IntPart)), '      "0.00")) * 100)


'  End If
'
'  OpStr = OpStr & " Only"
'
'  RupeesInWords = OpStr
'
'ExitHere:
'  On Error GoTo 0
'  Exit Function
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Function

Public Function RupeesInWords(ByVal Number As Double,Optional ByVal rs As String,Optional ByVal ps As String) As String

  Dim IntPart As String
  Dim tmpStr As String
  Dim tmpStr2 As String
  Dim OpStr As String
  Dim cnt As Integer
  Dim TwoDigitNo As Integer

  Try

  IntPart = Int(Number / 1#)

  OpStr = ""
  
  If Val(IntPart) > 0 Then
    tmpStr = IntPart
    Do
'      If Trim$(OpStr) <> "" And Len(tmpStr) > 7 Then
'        OpStr = "Crores " & OpStr
'      End If
      
      'Last 7 digits of remaining string
      tmpStr2 = Right$(tmpStr, 7)
  
      'To insert a zero before hundreth place digit ex: 4562 to 40562
      If Len(tmpStr2) >= 3 Then
        tmpStr2 = Left$(tmpStr2, Len(tmpStr2) - 3) & "0" & Right$(tmpStr2, 3)
      End If
      
      'To make it even no. digits
      If Len(tmpStr2) Mod 2 = 1 Then
        tmpStr2 = "0" & tmpStr2
      End If
      
      For cnt = 1 To Round((Len(tmpStr2) / 2) + 0.01)
        Select Case cnt
          Case 1    'Units
            TwoDigitNo = Val(Right$(tmpStr2, 2))
            OpStr = TwoDigitNumberToWord(TwoDigitNo) & " " & OpStr
          Case 2    'Hundreds
            TwoDigitNo = Val(Mid$(tmpStr2, Len(tmpStr2) - 3, 2))
            If TwoDigitNo > 0 Then
              OpStr = TwoDigitNumberToWord(TwoDigitNo) & " Hundred " & IIf(Len(Trim$(OpStr)) > 0, "and ", " ") & OpStr
            End If
          Case 3    'Thousands
            TwoDigitNo = Val(Mid$(tmpStr2, Len(tmpStr2) - 5, 2))
            If TwoDigitNo > 0 Then
              OpStr = TwoDigitNumberToWord(TwoDigitNo) & " Thousand " & OpStr
            End If
          Case 4    'Lakhs
            TwoDigitNo = Val(Mid$(tmpStr2, Len(tmpStr2) - 7, 2))
            If TwoDigitNo > 0 Then
              OpStr = TwoDigitNumberToWord(TwoDigitNo) & IIf(TwoDigitNo = 1, " Lakh ", " Lakhs ") & OpStr
            End If
        End Select
      Next cnt
  
      If Len(tmpStr) > 7 Then
        tmpStr = Left$(tmpStr, Len(tmpStr) - 7)
        OpStr = "Crores " & OpStr
      Else
        tmpStr = ""
      End If
    Loop While Len(Trim$(tmpStr)) > 0
    
    'Add Rupees String
    OpStr = rs & " " & OpStr
  End If
  
  'For Paise Conversion and Paise String Add
  If Val(IntPart) <> Number Then
    OpStr = OpStr & IIf(Len(Trim$(OpStr)) > 0, " and ", "") & ps & " " & TwoDigitNumberToWord(Val(Format((Number - Val(IntPart)), "0.00")) * 100)


  End If
  
  OpStr = Trim$(OpStr) & " Only"
  
  RupeesInWords = OpStr
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function ServerDate() As Date
  Dim sdate As AceADODB.Recordset
  Dim sql As String

  Try

  sql = "select getdate()"
  Set sdate = New AceADODB.Recordset
  If RSOpen(sdate, sql) = False Then
    GoTo ExitHere
  End If
  ServerDate = Format(sdate(0), "dd/MM/yyyy HH:mm:ss")
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set sdate = Nothing
End Try
End Function

Public Sub SetApplicationCaption(ByVal frm As MDIForm)
  Dim sql As String
  
' On Error Resume Next
  
  sql = App.Path & "\" & App.EXEName & ".exe"
  If Dir$(sql) <> "" Then
    frm.Caption = App.Title & " - " & gCompName & " [" & Format(FinStart, "MM-yyyy") & " " & Format(FinEnd, "MM-yyyy") & "]" & Space(5) & " Build: " & FileDateTime(sql)


  Else
    frm.Caption = App.Title & " - " & gCompName & " [" & Format(FinStart, "MM-yyyy") & " " & Format(FinEnd, "MM-yyyy") & "]"

  End If

End Sub

Public Function SetEffStatus(ByVal sChar As String) As String
' On Error Resume Next
  
  Select Case sChar
    Case "New"
      SetEffStatus = "N"
    Case "Modify"
      SetEffStatus = "M"
    Case "Authorised"
      SetEffStatus = "A"
    Case "Amendment"
      SetEffStatus = "E"
    Case "Hold"
      SetEffStatus = "H"
    Case "ShortClose"
      SetEffStatus = "C"
    Case "Active"
      SetEffStatus = "Y"
    Case "Inactive"
      SetEffStatus = "I"
    Case "Cancel"
      SetEffStatus = "X"
    Case "Delete"
      SetEffStatus = "D"
    Case Else
      SetEffStatus = "Null"
  End Select
  
End Function

Public Function SetItemNature(ByVal ItemType As String) As String
  Dim tmp As String
  
' On Error Resume Next
  Select Case ItemType
    Case "R"
      tmp = "RAW MATERIAL"
    Case "C"
      tmp = "COMPONENT"
    Case "F"
      tmp = "FINISHED GOODS"
    Case "S"
      tmp = "SUB ASSEMBLY"
    Case "P"
      tmp = "CAPITAL"
    Case "T"
      tmp = "TOOL"
    Case "A"
      tmp = "SPARES"
    Case "M"
      tmp = "MISCELLANEOUS"
    Case "Z"
      tmp = "SCRAP"
    Case "I"
      tmp = "INSTRUMENT"
    Case "L"
      tmp = "CONSUMABLE"
    Case "CC"
      tmp = "CUSTOMER SUPPLIED COMPONENT"
    Case "CF"
      tmp = "CUSTOMER SUPPLIED FINSHED GOODS"
    Case "CR"
      tmp = "CUSTOMER SUPPLIED RAW MATERIAL"
      
    Case Else
      tmp = ""
  End Select
  SetItemNature = tmp
End Function
Public Function spCAlign(ByRef sSpread As AceSpread,ByVal  sCol As Long,ByVal  sRow As Long,ByVal  sColWith As Integer,ByVal  strText As String,Optional ByVal sFontBold As Boolean = False)

' On Error Resume Next
  
    sSpread.Row = sRow
sSpread.Col = sCol
    sSpread.SetText sCol, sRow, strText
    sSpread.ColWidth(sCol) = sColWith
    sSpread.TypeHAlign = TypeHAlignCenter
    If sFontBold = True Then
      sSpread.FontBold = True
    Else
      sSpread.FontBold = False
    End If
  
End Function

Public Function spLAlign(ByRef sSpread As AceSpread,ByVal  sCol As Long,ByVal  sRow As Long,ByVal  sColWith As Integer,ByVal  strText As String,Optional ByVal sFontBold As Boolean = False)

' On Error Resume Next
  
    sSpread.Row = sRow
sSpread.Col = sCol
    sSpread.SetText sCol, sRow, strText
    sSpread.ColWidth(sCol) = sColWith
    sSpread.TypeHAlign = TypeHAlignLeft
    If sFontBold = True Then
      sSpread.FontBold = True
    Else
      sSpread.FontBold = False
    End If
  
End Function

Public Function spRAlign(ByRef sSpread As AceSpread,ByVal  sCol As Long,ByVal  sRow As Long,ByVal  sColWith As Integer,ByVal  strText As String,Optional ByVal sFontBold As Boolean = False)

' On Error Resume Next
  
    sSpread.Row = sRow
sSpread.Col = sCol
    sSpread.SetText sCol, sRow, strText
    sSpread.ColWidth(sCol) = sColWith
    sSpread.TypeHAlign = TypeHAlignRight
    If sFontBold = True Then
      sSpread.FontBold = True
    Else
      sSpread.FontBold = False
    End If
  
End Function

Public Function SqlQueryResultDebugPrint(ByVal sql As String) As Boolean
  Dim rs As AceADODB.Recordset
  Dim cCnt As Long

  Try
  SqlQueryResultDebugPrint = False

  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    Do While Not rs.EOF
      For cCnt = 0 To rs.Fields.Count - 1
        Debug.Print "'" & rs(cCnt) & "', ";
      Next cCnt
      Debug.Print
      rs.MoveNext
    Loop
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Sub sShowMessage(ByVal MsgStr As String,Optional ByVal IconToShow As eMsgBoxIcon = eMsgBoxIcon.mInformation,Optional ByVal MsgBoxTitle As String = "")
' On Error Resume Next
  MsgBox MsgStr, IconToShow + vbOKOnly, IIf(Trim$(MsgBoxTitle) = "", App.Title, MsgBoxTitle)
End Sub

Public Function StepClear(ByVal frm As Form) As Boolean
  Dim cnt As Integer
  
  Try
  StepClear = False
  
    For cnt = 0 To frm.frastep.UBound
      frm.frastep(cnt).Enabled = False
      frm.CmdStep(cnt).Enabled = False
      frm.CmdView.Enabled = False
    Next cnt
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Sub StepEnable(ByVal sForm As Form)
  Dim i As Long
  
  Try

  For i = 0 To sForm.frastep.UBound
    If sForm.frastep(i).Visible = True Then
      sForm.frastep(i).Enabled = True
      sForm.CmdStep(i).Enabled = True
      Exit For
    End If
  Next i
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Public Function StepProceedNext(ByVal frm As Form,ByVal  Index As Integer) As Boolean

  Try
  StepProceedNext = False
  
    If frm.frastep.UBound > Index Then        'Modify Rajapriyan on 08/05/2008
      Do While frm.CmdStep(Index + 1).Visible = False
        frm.frastep(Index).Enabled = False
        frm.CmdStep(Index).Enabled = False
        Index = Index + 1
        If frm.CmdStep.UBound = Index Then
          frm.CmdView.Enabled = True
          If frm.frastep(Index).Visible = True Then
            frm.frastep(Index).Enabled = False
            frm.CmdStep(Index).Enabled = False
          End If
          Exit Do
        End If
      Loop
      
    End If                                  'end
    
    If frm.frastep.UBound > Index Then
      frm.frastep(Index + 1).Enabled = True
      frm.CmdStep(Index + 1).Enabled = True
      frm.frastep(Index).Enabled = False
      frm.CmdStep(Index).Enabled = False
    Else
      frm.CmdView.Enabled = True
      frm.frastep(Index).Enabled = False
      frm.CmdStep(Index).Enabled = False
    End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function ToUpper(ByVal Ascii As Integer) As Integer
  'To make Character in Upper case while typing
' On Error Resume Next
  Select Case Ascii
    Case 97 To 122
      ToUpper = Ascii - 32
    Case Else
      ToUpper = Ascii
  End Select
End Function

Public Sub WordSplit(ByVal InpStr As String,ByVal  WidthPerLine As Integer)
Dim    mSplit() As String
Dim  mstr As String
Dim  Pos As Integer
Dim  Pos1 As Integer

  mstr = Trim$(InpStr)
  If mstr = "" Then
    ReDim mSplit(1)
ReDim WSplit(1)

    mSplit(0) = 1
mSplit(1) = mstr
    WSplit = mSplit
    Exit Sub
  End If
  If WidthPerLine < 5 And Len(mstr) < 5 Then
    MsgBox "The Length of Line is Very Minimum. So Exiting...", vbOKOnly, App.Title

    Exit Sub
  End If
  ReDim mSplit(0)
  Do While Len(mstr) > WidthPerLine
    Pos = InStrRev(mstr, " ", WidthPerLine + 1, vbTextCompare)
    'Pos1 = InStrRev(mstr, Chr(13), WidthPerLine + 1, vbTextCompare)
    Pos1 = InStr(1, mstr, Chr(13), vbTextCompare)
    
    'If Pos1 > Pos Then Pos = Pos1
    If Pos1 > 0 Then
      If Pos1 < Pos Then
        Pos = Pos1
      Else
        If Pos1 < WidthPerLine Then
          Pos = Pos1
        End If
      End If
    End If
    ReDim Preserve mSplit(UBound(mSplit) + 1)
    If Pos <> 0 Then
      mSplit(UBound(mSplit)) = Left$(mstr, Pos - 1)
      mstr = Mid$(mstr, Pos + 1)
    Else
      mSplit(UBound(mSplit)) = Left$(mstr, WidthPerLine - 1) & "-"
      mstr = Mid$(mstr, WidthPerLine - 1)
    End If
  Loop
  If Trim$(mstr) <> "" Then
    ReDim Preserve mSplit(UBound(mSplit) + 1)
    mSplit(UBound(mSplit)) = mstr
  End If
  mSplit(0) = UBound(mSplit)
  ReDim WSplit(mSplit(0))
  WSplit = mSplit
End Sub

Private Function DBConnectThruIni() As Boolean

  Dim dname As String
  Dim server As String
  Dim uid As String
  Dim uid1 As String
  Dim pwd1 As String
  Dim db1 As String
  Dim server1 As String
  Dim port1 As String
  Dim rettype As String
  Dim ConStr As String
  
  Dim ret As Boolean
  
  Try
  DBConnectThruIni = False
  
  dname = App.Path & "\ACE.ini"
  server = "Server"
  
  uid = "uid"
  rettype = String(100, 32)
  ret = GetPrivateProfileString(server, uid, " ", rettype, Len(rettype), dname)
  rettype = Mid(rettype, 1, Len(RTrim(rettype)) - 1)
  uid1 = rettype
  
  If Trim$(uid1) = "" Then
    sShowMessage "Please provide UserID."
    GoTo ExitHere
  End If

  uid = "pwd"
  rettype = String(100, 32)
  ret = GetPrivateProfileString(server, uid, " ", rettype, Len(rettype), dname)
  rettype = Mid(rettype, 1, Len(RTrim(rettype)) - 1)
  pwd1 = rettype

  uid = "database"
  rettype = String(100, 32)
  ret = GetPrivateProfileString(server, uid, " ", rettype, Len(rettype), dname)
  rettype = Mid(rettype, 1, Len(RTrim(rettype)) - 1)
  db1 = rettype
  gDatabase = Trim$(db1)
  
  If Trim$(db1) = "" Then
    sShowMessage "Database Name Missing."
    GoTo ExitHere
  End If
  
  uid = "server"
  rettype = String(100, 32)
  ret = GetPrivateProfileString(server, uid, " ", rettype, Len(rettype), dname)
  rettype = Mid(rettype, 1, Len(RTrim(rettype)) - 1)
  server1 = rettype
  gServer = server1
  
  If Trim$(server1) = "" Then
    sShowMessage "Database Name Missing."
    GoTo ExitHere
  End If
  
  uid = "port"
  rettype = String(100, 32)
  ret = GetPrivateProfileString(server, uid, " ", rettype, Len(rettype), dname)
  rettype = Mid(rettype, 1, Len(RTrim(rettype)) - 1)
  port1 = rettype
  
  If Trim$(port1) = "" Then
    port1 = "1433"
  End If
  
'  #If isWeb = False Then
'    ConStr = "Provider=SQLOLEDB.1;uid=" & uid1 & ";pwd=" & pwd1 & ";database=" '      & db1 & ";server=" & server1 & ""

'  #Else
'     ConStr = "Driver={SQL Server};" & "Server=" & server1 & ";" & "Address=" & '       server1 & "," & port1 & ";" & "Network=DBMSSOCN;" & "Database=" & db1 & '       ";" & "Pwd=" & pwd1 & ";" & "Uid=" & uid1 & ""


       
       ConStr = "Provider=SQLOLEDB.1;uid=" & uid1 & ";pwd=" & pwd1 & ";database=" & db1 & ";server=" & server1 & ""
'  #End If
  
  Set db = New AceADODB.Connection
  db.CursorLocation = adUseClient
  db.CommandTimeout = 100         'db.CommandTimeout = 600 - changed on 25/11/2025
  db.Open ConStr
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function DBConnectThruPrompt() As Boolean

  Try
  DBConnectThruPrompt = False
  
  If frmDBConnect.ConnectToDatabase = False Then
    sShowMessage "Error - while trying to set up Database Connection."
    GoTo ExitHere
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub FormMenuSecurity(ByVal MenuForm As MDIForm,ByVal  ThisModuleName As String,ByVal  sUser As String)

  Dim rsUser As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  
  Dim sql As String
  Dim CurTopMenu As String
  
  Dim MenuCtl As Control
   
  Try
  
  Set rsUser = New AceADODB.Recordset
  sql = "Select MenuName, MenuIndex from Sec_UserRights"
  sql = sql & " where ModuleName='" & ThisModuleName & "' and UserID='" & gUserID & "'"

  sql = sql & " and BranchID='" & gBranchID & "' and CompId='" & gCompID & "' and finyear ='" & gFinyear & "'"

  
  Set rsUser = New AceADODB.Recordset
  If RSOpen(rsUser, sql) = False Then
    GoTo ExitHere
  End If
  
  sql = "Select Distinct ERPNo,MenuName,MenuIndex from Sec_ModuleFormList where ModuleName='" & ThisModuleName & "'"

  Set snap = New AceADODB.Recordset
  If RSOpen(snap, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsUser.EOF Then
    For Each MenuCtl In MenuForm.Controls
      If TypeOf MenuCtl Is Menu Then
        Select Case MenuCtl.Name
          Case "mnuFile", "mnuModules", "mnuHelp", "mnuModuleName"
            CurTopMenu = MenuCtl.Name     'No Need to Verify
          Case Else
            Select Case MenuCtl.Name
              Case "mnuMaster", "mnuTran", "mnuReports", "mnuListing", "mnuOptions"

                CurTopMenu = MenuCtl.Name 'Must be taken to verify
            End Select
            If CurTopMenu <> "mnuModules" Then    'Security fix by another procedure
              If CurTopMenu <> "mnuFile" Then     'Full rights given - so don't take
                If CurTopMenu <> "mnuHelp" Then   'Full rights given - so don't take
                  rsUser.Filter = adFilterNone
                  rsUser.Filter = "MenuName='" & MenuCtl.Name & "' and MenuIndex='" & MenuCtl.Index & "'"

                  snap.Filter = adFilterNone
                  snap.Filter = "MenuName='" & MenuCtl.Name & "' and MenuIndex='" & MenuCtl.Index & "'"

                  If Not snap.EOF Then
                    MenuCtl.Caption = snap("ERPNo") & " " & MenuCtl.Caption
                  End If
                  If sUser = "A" Or sUser = "D" Then
                    'no tran
                  Else
                    If rsUser.EOF Then
                      MenuCtl.Enabled = False
                    Else
                      MenuCtl.Enabled = True
                    End If
                  End If
                End If
              End If
            End If
        End Select
      End If
    Next
  Else
    For Each MenuCtl In MenuForm.Controls
      If TypeOf MenuCtl Is Menu Then
        Select Case MenuCtl.Name
          Case "mnuFile", "mnuModules", "mnuHelp", "mnuModuleName"
            CurTopMenu = MenuCtl.Name     'No Need to Verify
          Case Else
            Select Case MenuCtl.Name
              Case "mnuMaster", "mnuTran", "mnuReports", "mnuListing", "mnuOptions"

                CurTopMenu = MenuCtl.Name 'Must be taken to verify
            End Select
            If CurTopMenu <> "mnuModules" Then    'Security fix by another procedure
              If CurTopMenu <> "mnuFile" Then     'Full rights given - so don't take
                If CurTopMenu <> "mnuHelp" Then   'Full rights given - so don't take
                  snap.Filter = adFilterNone
                  snap.Filter = "MenuName='" & MenuCtl.Name & "' and MenuIndex='" & MenuCtl.Index & "'"

                  If Not snap.EOF Then
                    MenuCtl.Caption = snap("ERPNo") & " " & MenuCtl.Caption
                  End If
                  If sUser = "A" Or sUser = "D" Then
                    'no tran
                  Else
                    If rsUser.EOF Then
                      MenuCtl.Enabled = False
                    Else
                      MenuCtl.Enabled = True
                    End If
                  End If
                End If
              End If
            End If
        End Select
      End If
    Next
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsUser = Nothing
  Set snap = Nothing
  Set MenuCtl = Nothing
End Try
End Sub

Private Sub QMS_FormMenuSecurity(ByVal MenuForm As MDIForm,ByVal  ThisModuleName As String,ByVal  sUser As String)

  Dim rsUser As AceADODB.Recordset
  Dim snap As AceADODB.Recordset
  
  Dim sql As String
  Dim CurTopMenu As String
  
  Dim MenuCtl As Control
   
  Try
  
  Set rsUser = New AceADODB.Recordset
  sql = "Select MenuName, MenuIndex from Sec_UserRights"
  sql = sql & " where ModuleName='" & ThisModuleName & "' and UserID='" & gUserID & "'"

  sql = sql & " and BranchID='" & gBranchID & "' and CompId='" & gCompID & "' and finyear ='" & gFinyear & "'"

  
  Set rsUser = New AceADODB.Recordset
  If RSOpen(rsUser, sql) = False Then GoTo ExitHere
  
  sql = "Select Distinct ERPNo,MenuName,MenuIndex from Sec_ModuleFormList where ModuleName='" & ThisModuleName & "'"

  Set snap = New AceADODB.Recordset
  If RSOpen(snap, sql) = False Then GoTo ExitHere
  
  If Not rsUser.EOF Then
    For Each MenuCtl In MenuForm.Controls
      If TypeOf MenuCtl Is Menu Then
        Select Case MenuCtl.Name
          Case "mnuFile", "mnuModules", "mnuHelp", "mnuModuleName"
            CurTopMenu = MenuCtl.Name     'No Need to Verify
          Case Else
            Select Case MenuCtl.Name
              Case "mnuMaster", "mnuTran", "mnuReports", "mnuListing", "mnuOptions"

                CurTopMenu = MenuCtl.Name 'Must be taken to verify
            End Select
            If CurTopMenu <> "mnuModules" Then    'Security fix by another procedure
              If CurTopMenu <> "mnuFile" Then     'Full rights given - so don't take
                If CurTopMenu <> "mnuHelp" Then   'Full rights given - so don't take
                  rsUser.Filter = adFilterNone
                  rsUser.Filter = "MenuName='" & MenuCtl.Name & "' and MenuIndex='" & MenuCtl.Index & "'"

                  snap.Filter = adFilterNone
                  snap.Filter = "MenuName='" & MenuCtl.Name & "' and MenuIndex='" & MenuCtl.Index & "'"

                  If Not snap.EOF Then
                    MenuCtl.Caption = IIf((MenuCtl.Index >= 100 And MenuCtl.Index <= 250), "", snap("ERPNo") & " ") & MenuCtl.Caption
                  End If
                  If sUser = "A" Or sUser = "D" Then
                    'no tran
                  Else
                    If rsUser.EOF Then
                      MenuCtl.Enabled = False
                    Else
                      MenuCtl.Enabled = True
                    End If
                  End If
                End If
              End If
            End If
        End Select
      End If
    Next
  Else
    For Each MenuCtl In MenuForm.Controls
      If TypeOf MenuCtl Is Menu Then
        Select Case MenuCtl.Name
          Case "mnuFile", "mnuModules", "mnuHelp", "mnuModuleName"
            CurTopMenu = MenuCtl.Name     'No Need to Verify
          Case Else
            Select Case MenuCtl.Name
              Case "mnuMaster", "mnuTran", "mnuReports", "mnuListing", "mnuOptions"

                CurTopMenu = MenuCtl.Name 'Must be taken to verify
            End Select
            If CurTopMenu <> "mnuModules" Then    'Security fix by another procedure
              If CurTopMenu <> "mnuFile" Then     'Full rights given - so don't take
                If CurTopMenu <> "mnuHelp" Then   'Full rights given - so don't take
                  snap.Filter = adFilterNone
                  snap.Filter = "MenuName='" & MenuCtl.Name & "' and MenuIndex='" & MenuCtl.Index & "'"

                  If Not snap.EOF Then
                  
                    MenuCtl.Caption = IIf((MenuCtl.Index >= 100 And MenuCtl.Index <= 250), "", snap("ERPNo") & " ") & MenuCtl.Caption
                  End If
                  If sUser = "A" Or sUser = "D" Then
                    'no tran
                  Else
                    If rsUser.EOF Then
                      MenuCtl.Enabled = False
                    Else
                      MenuCtl.Enabled = True
                    End If
                  End If
                End If
              End If
            End If
        End Select
      End If
    Next
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsUser = Nothing
  Set snap = Nothing
  Set MenuCtl = Nothing
End Try
End Sub

Private Function TwoDigitNumberToWord(ByVal Number As Integer) As String
  Dim i As Integer
  Dim j As Integer
  Dim A(20) As String
  Dim B(10) As String

  Try

  If Number >= 10 ^ 2 Then
    MsgBox " INVALID DATA ", 32, "STOP"
    Exit Function
  End If

  A(0) = ""
A(1) = "One"
A(2) = "Two"
A(3) = "Three"
A(4) = "Four"
A(5) = "Five"
A(6) = "Six"
A(7) = "Seven"

  A(8) = "Eight"
A(9) = "Nine"
A(10) = "Ten"
A(11) = "Eleven"
A(12) = "Twelve"
A(13) = "Thirteen"
A(14) = "Fourteen"

  A(15) = "Fifteen"
A(16) = "Sixteen"
A(17) = "Seventeen"
A(18) = "Eighteen"
A(19) = "Nineteen"

  B(0) = ""
B(1) = ""
B(2) = "Twenty"
B(3) = "Thirty"
B(4) = "Forty"

  B(5) = "Fifty"
B(6) = "Sixty"
B(7) = "Seventy"
B(8) = "Eighty"
B(9) = "Ninety"

  If Number > 19 Then
    i = Int(Number / 10)
    j = Int(Number - (i * 10))
    TwoDigitNumberToWord = B(i) & " " & A(j)
    Exit Function
  Else
    TwoDigitNumberToWord = A(Number)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function LoadComboItemNature(ByVal cmbctrl As ComboBox,Optional ByVal isReqItemNatureAll As Boolean = True) As String
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  
  Try
  
  sql = " Execute COM_ItemNature_Fil"
  sql = sql & " @isReqItemNatureAll='" & IIf(isReqItemNatureAll = True, 1, 0) & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, sql) = False Then
    GoTo ExitHere
  End If
  
  cmbctrl.Clear
  If Not rs.EOF Then
    For rCnt = 1 To rs.RecordCount
      cmbctrl.AddItem (Trim$(rs("ItemNature")))
      rs.MoveNext
    Next rCnt
    cmbctrl.ListIndex = 0
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function AvoidToEnterSpecialCharacters(ByVal KeyAscii As Integer) As Integer
  Dim CharCnt As Integer
  
' On Error Resume Next
  AvoidToEnterSpecialCharacters = 0
  
  ' ! "  # $ % & ^ * ( ) + , . ~ ' : ; ? \ > < Space ~
  Select Case KeyAscii
    Case 13, 32 To 44, 46, 58 To 64, 91 To 96, 123 To 127
      Exit Function
  End Select
  
  AvoidToEnterSpecialCharacters = KeyAscii
End Function

Public Function StockValidation(ByVal fItemCode As String,ByVal  fWHSCode As String,ByVal  fReqdQty As Double,Optional ByVal RefType As String = "",Optional ByVal RefNo As Long = 0,Optional ByVal RefDate As Date = "1900-01-01") As Boolean

  Dim sql As String
  Dim rs As AceADODB.Recordset
  
  Dim CurDate As Date
  
  Try
  
  StockValidation = False
  
  'Get ServerDate
  CurDate = ServerDate
  
  sql = " Execute INV_StockConversion_Val_Stock"
  
  sql = sql & " @ItemCode='" & Trim$(fItemCode) & "',"
  sql = sql & " @WHSCode='" & Trim$(fWHSCode) & "',"
  
  If Trim$(RefType) <> "" Then
    sql = sql & " @RefType='" & Trim$(RefType) & "',"
    sql = sql & " @RefNo=" & Val(RefNo) & ","
    sql = sql & " @RefDate='" & Format(RefDate, "yyyy-MM-dd") & "',"
  Else
    sql = sql & " @RefType='',"
    sql = sql & " @RefNo=0,"
    sql = sql & " @RefDate='1900-01-01',"
  End If
  
  sql = sql & " @AsOnDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
  sql = sql & " @RptDate='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "',"
  sql = sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "', @Finyear='" & gFinyear & "'"

  If RSOpen(rs, sql) = False Then GoTo ExitHere
  
  If Not rs.EOF Then
    If rs("StockQty") < Val(fReqdQty) Then
      sShowMessage "Required Qty for the Item " & Trim$(fItemCode) & " is less than the available Stock Qty " & Val(rs("StockQty")) & "."
'      Call DBRollBackWithMessage(mCustom, "Required Qty for the Item " & Trim$(fItemCode) & " is less than the available Stock Qty " & Val(rs("StockQty")) & ".")
      GoTo ExitHere
    Else
      'Continue
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function GetFileName(ByVal fRptHeadName As String) As String
  Dim rCnt As Long
  Dim retVal As Variant
  
  Try
  
  For rCnt = 1 To Len(fRptHeadName)
    Select Case Mid$(fRptHeadName, rCnt, 1)
      Case "a" To "z", "A" To "Z", "0" To "9"
        retVal = retVal & Mid$(fRptHeadName, rCnt, 1)
    End Select
  Next rCnt
  
  GetFileName = retVal
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function
