Attribute VB_Name = "mdlFormula"
Option Explicit
Option Compare Text

Private Const PercBlockOpen As String = "�"
Private Const PercBlockClose As String = "�"

Private Const SpreadTwoRows As Integer = 485

Private Enum eParamFrom
  mBill = 1
  mBook = 2
End Enum

'For Calc Spread Column Numbers
Private Enum eCalc
  mDesc = 1
  mValue = 2
End Enum

Public Enum eVrType
  mDebit = 0
  mCredit = 1
  mSales = 2
  mPurchase = 3
End Enum

Public Enum eDrCrCalcFor
  mNone = 0
  mCalculationMistake = 1
  mSequenceChange = 2
  mTermsDifference = 3
  mRateDifference = 4
  mQuantityDifference = 5
End Enum

Public Function CalculateExprValues(ByVal FormulaID As Integer,ByVal  fUserKey As String) As Boolean
  Dim Sql As String
  Dim rsExpr As AceADODB.Recordset
  Dim rs As AceADODB.Recordset
  Dim rsCnt As AceADODB.Recordset
  Dim rsROff As AceADODB.Recordset
  Dim ROff As eRoundOffOptions
  Dim ExprStr As String
  Dim tmpStr As String
  Dim ExprResult() As Double
  
  Dim GTFlag As Integer
  Dim CurExprSeq As Integer
  Dim RecExprSeq As Integer
  Dim VBScr As ScriptControl
  
  Dim MultiCnt As Integer                       'Max Seq for current expression
  Dim MultiExpr() As String                     '(MultiCnt) Single Dimensional Array for MultiSeq Values
  Dim MultiSeqFound() As Boolean                '(MultiCnt) Single Dimensional Array to store whether Seq no found in the database
  Dim sCnt As Integer
  
  Dim pLevel As Integer                         'Percentage Block Level
  Dim MultiSeqLockStatus() As Boolean           '(MultiCnt) Whether the Sequence is locked for appending formula
  Dim MultiSeqLockedLevel() As Integer          '(MultiCnt) When it is locked
  Dim CallParamIsValue() As Boolean             '(MultiCnt)
  
  Dim BlockContinue As Boolean                  'To Skip the Percentage Block Expression
  Dim pCnt As Integer
  Dim varBookMark As Variant
  
  Try
  CalculateExprValues = False

  Sql = "Select a.ExprSeq, b.ExprGroupName as ExprName, a.RoundOff, a.RoundOffValue, a.GrandTotalFlag, a.PostingAc"
  Sql = Sql & " from Com_FormulaMast_Dtl a Left Outer Join COM_FormulaExpr_Group b "
  Sql = Sql & " On a.ExprRoute=b.Route+b.Width and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
  Sql = Sql & " where a.FrmID=" & FormulaID & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Sql = Sql & " Order by ExprSeq"
  Set rsExpr = New AceADODB.Recordset
  If RSOpen(rsExpr, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsExpr.EOF Then
    rsExpr.MoveLast
rsExpr.MoveFirst
    ReDim ExprResult(rsExpr.RecordCount) As Double
    
    Sql = "Select a.ExprSeq, isNull(Max(b.SeqCnt), 0) as MultiCnt"
    Sql = Sql & " from COM_FormulaMast_Block_Dtl a Left outer join"
    Sql = Sql & " (Select BlockType, Block, Count(KeyCol) as SeqCnt from COM_tmpFormulaCalc_Dtl"
    Sql = Sql & " Where DTKey='" & fUserKey & "' and KeyCol<>'' and FrmID=" & FormulaID & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    Sql = Sql & " Group by BlockType, Block) b"
    Sql = Sql & " on a.Block=b.Block and a.BlockType=b.BlockType"
    Sql = Sql & " Where a.FrmID=" & FormulaID & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    Sql = Sql & " Group by a.ExprSeq" ') x"
    Set rsCnt = New AceADODB.Recordset
    If RSOpen(rsCnt, Sql) = False Then
      GoTo ExitHere
    End If
  
    Sql = "Select * from dbo.fn_COM_CalculateExpression ('" & fUserKey & "', " & gModuleCode & ", " & FormulaID & ", '" & gUserID & "', '" & gBranchID & "', '" & gCompID & "')"
    Sql = Sql & " Order by ExprSeq, BlockSeq, MultiSeq"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    
    If rs.EOF Then
      GoTo ExitHere
    End If
    
    If DBStartTransaction = False Then
 GoTo ExitHere
End If

    'To delete the Previous Data in Temp table for the Current User
    Sql = "Delete From Com_tmpFormulaCalc_Hdr"
    Sql = Sql & " where DTKey='" & fUserKey & "' and UserID='" & gUserID & "' and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    If DBExecuteSql(db, sql) = False Then GoTo ExitHere

    RecExprSeq = rs("ExprSeq")    'Initialize with First Record ExprSeq
    CurExprSeq = RecExprSeq
    
    Set VBScr = New ScriptControl
    VBScr.Language = "VBScript"
    'To Add IIF function facility to Windows Script Control. This codes add to Global module of script control
    VBScr.AddCode ("Function IIf(expr, truepart, falsepart) " & Chr(13) & " If expr Then IIf = truepart Else IIf = falsepart" & Chr(13) & "End function")
    
    
    MultiCnt = 0
    If Not rsCnt.EOF Then
      rsCnt.Filter = "ExprSeq=" & rs("ExprSeq")
      If Not rsCnt.EOF Then
        MultiCnt = rsCnt("MultiCnt")
      End If
    End If
    
    ReDim MultiExpr(MultiCnt)
    ReDim MultiSeqFound(MultiCnt)
    ReDim MultiSeqLockStatus(MultiCnt)
    ReDim MultiSeqLockedLevel(MultiCnt)
    ReDim CallParamIsValue(MultiCnt)
    
    pLevel = 0
    
    For sCnt = 0 To MultiCnt
MultiSeqLockedLevel(sCnt) = -1
Next sCnt
    
    If MultiCnt > 0 Then
      varBookMark = rs.Bookmark
      rs.Filter = "ExprSeq = 1 and MultiSeq > 0"
      If rs.EOF Then
        MultiSeqFound(0) = True
      Else
        Do While Not rs.EOF
          MultiSeqFound(rs("MultiSeq")) = True
          rs.MoveNext
        Loop
      End If
      
      rs.Filter = adFilterNone
      rs.Bookmark = varBookMark
    Else
      MultiSeqFound(0) = True
    End If
    
    Do While Not rs.EOF
      Select Case rs("BlockType")
        Case "ExprValue"
          MultiExpr(0) = CheckNull(rs("ValueUsed"))
          MultiCnt = 0
        Case "Param"
          tmpStr = "(" & CheckNull(rs("ValueUsed"), 0)
          If rs("IsTypePercentage") = True Then
            tmpStr = tmpStr & " / 100)"
            CallParamIsValue(rs("MultiSeq")) = False
          Else
            tmpStr = tmpStr & ")"
            CallParamIsValue(rs("MultiSeq")) = True
          End If
          
          If MultiCnt > 0 And rs("MultiSeq") = 0 Then
            For sCnt = 1 To MultiCnt
              If MultiSeqLockStatus(sCnt) = False Then
                If MultiSeqFound(sCnt) = True Then
                  MultiExpr(sCnt) = MultiExpr(sCnt) & Space(1) & tmpStr
                End If
              End If
            Next sCnt
          Else
            If MultiSeqLockStatus(rs("MultiSeq")) = False Then
              MultiExpr(rs("MultiSeq")) = MultiExpr(rs("MultiSeq")) & Space(1) & tmpStr
            End If
          End If
          
          If MultiCnt = 0 And CallParamIsValue(rs("MultiSeq")) = True Then
            'Param is Value & No Multi Seq then skip the percentage block
            tmpStr = tmpStr & ")"
            pCnt = 0
            Do While BlockContinue
              rs.MoveNext
              If rs.EOF Then
                BlockContinue = False
              Else
                If rs("BlockType") = "PercOpen" Then
                  pCnt = pCnt + 1
                ElseIf rs("BlockType") = "PercClose" Then
                  pCnt = pCnt - 1
                End If
                
                If pCnt = 0 Then
                  BlockContinue = False
                End If
              End If
            Loop
          End If

        Case Else
          'It all require same methodology i.e. all multi seq contain same value
          Select Case rs("BlockType")
            Case "ExprSeq"
              tmpStr = Space(1) & ExprResult(Val(rs("Block")))
            Case "LogicIF"
              tmpStr = Space(1) & "IIF(("
            Case "LogicTHEN", "LogicELSE"
              tmpStr = "), ("
            Case "LogicENDIF"
              tmpStr = "))"
            Case "LogicCompare", "LogicANDOR", "LogicNOT"
              tmpStr = Space(1) & rs("Block") & Space(1)
            Case "PercOpen"
              tmpStr = Space(1) & " * ("
              For sCnt = 0 To MultiCnt
                If CallParamIsValue(sCnt) = True Then
                  MultiSeqLockStatus(sCnt) = True
                  MultiSeqLockedLevel(sCnt) = pLevel
                End If
              Next sCnt
              pLevel = pLevel + 1
            Case "PercClose"
              tmpStr = ")"
            Case "OprOther", "OprOpenBrace"
              tmpStr = Space(1) & rs("Block")
            Case "OprCloseBrace"
              tmpStr = rs("Block")
            Case "NoBlock"
              tmpStr = Space(1) & rs("Block")
          End Select
          
          For sCnt = 0 To MultiCnt
            If MultiSeqLockStatus(sCnt) = False Then
              If MultiSeqFound(sCnt) = True Then
                If Trim$(MultiExpr(0)) <> "" And Trim$(MultiExpr(sCnt)) = "" Then
                  MultiExpr(sCnt) = MultiExpr(0)
                End If
                  
                MultiExpr(sCnt) = MultiExpr(sCnt) & tmpStr
              End If
            End If
          Next sCnt
          
          If rs("BlockType") = "PercClose" Then
            pLevel = pLevel - 1
            
            For sCnt = 0 To MultiCnt
              If MultiSeqLockedLevel(sCnt) = pLevel Then
                MultiSeqLockStatus(sCnt) = False
              End If
            Next sCnt
          End If
      End Select
      
      rs.MoveNext
        
      If Not rs.EOF Then
        RecExprSeq = rs("ExprSeq")    'Store Next ExprSeq
      Else
        RecExprSeq = 0
      End If
      
      If CurExprSeq <> RecExprSeq Then  'If running ExprSeq <> Recordset ExprSeq then Evaluate the Result
        If MultiCnt > 0 Then
          For sCnt = 1 To MultiCnt
            If MultiSeqFound(sCnt) = True Then
'            If Trim$(MultiExpr(sCnt)) <> "" Then
              'Stored in Array for future ExprSeq reference
              ExprResult(CurExprSeq) = ExprResult(CurExprSeq) + VBScr.Eval(MultiExpr(sCnt))
            End If
          Next sCnt
        Else
          'Stored in Array for future ExprSeq reference
          ExprResult(CurExprSeq) = ExprResult(CurExprSeq) + VBScr.Eval(MultiExpr(0))
        End If
        
        'Reinitialize the variable for multi seq expr
        If Not rs.EOF Then
          MultiCnt = 0
          If Not rsCnt.EOF Then
            rsCnt.Filter = "ExprSeq=" & rs("ExprSeq")
            If Not rsCnt.EOF Then
              MultiCnt = rsCnt("MultiCnt")
            End If
          End If
          
          Erase MultiExpr
            ReDim MultiExpr(MultiCnt)
          
          Erase MultiSeqFound
            ReDim MultiSeqFound(MultiCnt)
          Erase MultiSeqLockStatus
            ReDim MultiSeqLockStatus(MultiCnt)
          Erase MultiSeqLockedLevel
            ReDim MultiSeqLockedLevel(MultiCnt)
          Erase CallParamIsValue
            ReDim CallParamIsValue(MultiCnt)
          
          pLevel = 0
          
          For sCnt = 0 To MultiCnt
MultiSeqLockedLevel(sCnt) = -1
Next sCnt
          
          If MultiCnt > 0 Then
            varBookMark = rs.Bookmark
            rs.Filter = "MultiSeq > 0"
            If rs.EOF Then
              MultiSeqFound(0) = True
            Else
              rs.MoveFirst
              Do While Not rs.EOF
                MultiSeqFound(rs("MultiSeq")) = True
                rs.MoveNext
              Loop
            End If
            
            rs.Filter = adFilterNone
            rs.Bookmark = varBookMark
          Else
            MultiSeqFound(0) = True
          End If
        End If
        
        'Rounding Off Process
        rsExpr.Filter = adFilterNone
        rsExpr.Filter = "ExprSeq=" & CurExprSeq & ""
        If Not rsExpr.EOF Then
          If rsExpr("RoundOff") <> "N" Then
            Select Case rsExpr("RoundOff")
              Case "B"
                ROff = anNearest
              Case "U"
                ROff = anUpward
              Case "D"
                ROff = anDownward
            End Select
            'Update the Array with new Rounded off Value
            ExprResult(CurExprSeq) = RoundingOff(ExprResult(CurExprSeq), rsExpr("RoundOffValue"), , ROff)
          Else
            'Update the Array with new Rounded off Value for nearest paise
            ExprResult(CurExprSeq) = RoundingOff(ExprResult(CurExprSeq), 0.01, , anNearest)
          End If
          'For Grand Total Calculation
          GTFlag = rsExpr("GrandTotalFlag")
        Else
          GTFlag = 1
        End If
        
        'Insert the Result in Table for further process
        If SaveExpressionValuesToTemp(fUserKey, FormulaID, CurExprSeq, GTFlag, Format(ExprResult(CurExprSeq), "0.00")) = False Then
          GoTo ExitHere
        End If
              
        ExprStr = ""
        CurExprSeq = RecExprSeq
      End If
    Loop
    
    If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsExpr = Nothing
  Set rs = Nothing
  Set rsROff = Nothing
  Set VBScr = Nothing
End Try
End Function

Public Function CalculateFormula(ByVal fPartyCode As Integer,ByVal  fFormulaID As Integer,ByVal  UserKeyTxtBox As TextBox,ByVal  SupplyThruCodeValue() As String,ByVal  MultiItemList() As String,ByVal  GrandTotalTxtBox As TextBox,ByRef  ExprResultSpread As AceSpread,Optional ByVal ExprDescCol As Long = 1,Optional ByVal ExprValueCol As Long = 2,Optional ByVal ExprStartRow As Long = 1,Optional ByVal ScreenFrom As String,Optional ByVal IncludeSupplyThruCode As Boolean = True,Optional ByVal ShowPromptScreen As Boolean = True,Optional ByVal LockRefType As String,Optional ByVal LockRefNo As Integer,Optional ByVal LockRefDate As Date,Optional ByVal ValRefType As String,Optional ByVal ValRefNo As Integer,Optional ByVal ValRefDate As Date,Optional ByVal RefUserKey As String) As Boolean





  
  'SupplyThruCodeValue Should be in the format ParamID-MultiSeq-Value, ...
  'MultiItemList Should be in the format MultiSeq-Value, ...
  
  Dim fUserKey As String
  Dim fPrevUserKey As String
  
  Try
  CalculateFormula = False
  
  'To Clear the Result Spread
    If ScreenFrom <> "Formula Tester" And Left(ScreenFrom, 7) <> "Invoice" Then
      ExprResultSpread.ClearRange ExprDescCol, ExprStartRow, ExprDescCol, ExprResultSpread.MaxRows, True
    End If
    ExprResultSpread.ClearRange ExprValueCol, ExprStartRow, ExprValueCol, ExprResultSpread.MaxRows, True
  
  'Testing the SupplyThruCodeValue Format
  If IncludeSupplyThruCode = False Then
    If ValidateSupplyThruCodeValue(SupplyThruCodeValue()) = False Then
      GoTo ExitHere
    End If
  End If
  'Values in Given Format. So Proceed Further
  
  fPrevUserKey = Trim$(UserKeyTxtBox.Text)
  UserKeyTxtBox.Tag = fPrevUserKey
  
  fUserKey = CreateUserKey
  UserKeyTxtBox.Text = fUserKey
  
  If IncludeSupplyThruCode = False Then     'Supply thru code not included in prompt screen then it should insert to DB
    If fPrevUserKey <> "" Then              'Prev SC Param Values to be deleted in tmpTable
      If DeletePromptValue(fFormulaID, fPrevUserKey, True) = False Then
        GoTo ExitHere
      End If
    End If
    
    If InsertCodePromptValue(fFormulaID, fUserKey, SupplyThruCodeValue) = False Then
      GoTo ExitHere
    End If
  End If
  
  If fPrevUserKey = "" Then
    If InsertAlreadyLockedValues(fFormulaID, fUserKey, LockRefType, LockRefNo, LockRefDate, RefUserKey) = False Then
      GoTo ExitHere
    End If
    
    If ShowPromptScreen = True Then
      If frmPromptValueEntry.GetPromptValuesHelp(fPartyCode, fFormulaID, fUserKey, fPrevUserKey, ScreenFrom, MultiItemList, IncludeSupplyThruCode, LockRefType, LockRefNo, LockRefDate, ValRefType, ValRefNo, ValRefDate) = vbFalse Then

        GoTo ExitHere
      End If
    End If
  Else
    If ShowPromptScreen = True Then
      If InsertAlreadyLockedValues(fFormulaID, fUserKey, LockRefType, LockRefNo, LockRefDate, RefUserKey) = False Then
        GoTo ExitHere
      End If
    
      If frmPromptValueEntry.GetPromptValuesHelp(fPartyCode, fFormulaID, fUserKey, fPrevUserKey, ScreenFrom, MultiItemList, IncludeSupplyThruCode, LockRefType, LockRefNo, LockRefDate, ValRefType, ValRefNo, ValRefDate) = vbFalse Then

        GoTo ExitHere
      End If
    Else
      If UpdateUserKey(fFormulaID, fUserKey, fPrevUserKey, IncludeSupplyThruCode) = False Then
        GoTo ExitHere
      End If
    End If
  End If
  
  If CalculateExprValues(fFormulaID, fUserKey) = False Then
    GoTo ExitHere
  End If
  
  If Trim$(ScreenFrom) = "Formula Tester" Then
    If GetExprResults(fFormulaID, fUserKey, ExprResultSpread, GrandTotalTxtBox, ExprDescCol, ExprValueCol, ExprStartRow, False) = False Then
      GoTo ExitHere
    End If
  Else
    If GetExprResults(fFormulaID, fUserKey, ExprResultSpread, GrandTotalTxtBox, ExprDescCol, ExprValueCol, ExprStartRow) = False Then
      GoTo ExitHere
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function CalculateFormulaDrCr(ByVal fFormulaID As Integer,ByVal  FromUserKey As String,ByVal  ToUserKeyTxtBox As TextBox,ByVal  SupplyThruCodeValue() As String,ByVal  GrandTotalTxtBox As TextBox,ByRef  ExprResultSpread As AceSpread,Optional ByVal ExprDescCol As Long = 1,Optional ByVal ExprValueCol As Long = 2,Optional ByVal ExprStartRow As Long = 1,Optional ByVal DrCrCalcType As eDrCrCalcFor = eDrCrCalcFor.mNone) As Boolean



  
  Dim fUserKey As String
  Dim fPrevUserKey As String
  Dim SkipFixedValues As Boolean
  
  'SupplyThruCodeValue Should be in the format ParamID-MultiSeq-Value, ...
  'MultiItemList Should be in the format MultiSeq-Value, ...
  
  Try
  CalculateFormulaDrCr = False
  
  'Testing the SupplyThruCodeValue Format
  If ValidateSupplyThruCodeValue(SupplyThruCodeValue) = False Then
    GoTo ExitHere
  End If
  'Values in Given Format So Proceed Further
  
  fPrevUserKey = Trim$(ToUserKeyTxtBox.Text)
  ToUserKeyTxtBox.Tag = fPrevUserKey
  
  fUserKey = CreateUserKey
  ToUserKeyTxtBox.Text = fUserKey
  
  If fPrevUserKey <> "" Then              'Prev SC Param Values to be deleted in tmpTable
    If DeletePromptValue(fFormulaID, fPrevUserKey) = False Then
      GoTo ExitHere
    End If
  End If
  
  If InsertCodePromptValue(fFormulaID, fUserKey, SupplyThruCodeValue) = False Then
    GoTo ExitHere
  End If
  
  SkipFixedValues = IIf(DrCrCalcType = eDrCrCalcFor.mRateDifference, True, False)
  
  If CopyPromptValue(fFormulaID, FromUserKey, fUserKey, False, SkipFixedValues) = False Then
    GoTo ExitHere
  End If
      
  If CalculateExprValues(fFormulaID, fUserKey) = False Then
    GoTo ExitHere
  End If
  
  If GetExprResults(fFormulaID, fUserKey, ExprResultSpread, GrandTotalTxtBox, ExprDescCol, ExprValueCol, ExprStartRow) = False Then
    GoTo ExitHere
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function CheckBlockIntegrity(ByVal fExpr As String) As Boolean
  Dim Expr As String
  Dim iEnd As String
  Dim iStart As String
  
  Dim IfCnt As Integer
  Dim ThenCnt As Integer
  Dim ElseCnt As Integer
  Dim EndCnt As Integer
  
  Try
  CheckBlockIntegrity = False
  
  Expr = fExpr
  
  'Check If-Then-Else-Endif Count
  IfCnt = GetStrCount("IF {", fExpr)
  ThenCnt = GetStrCount("} THEN {", fExpr)
  ElseCnt = GetStrCount("} ELSE {", fExpr)
  EndCnt = GetStrCount("} ENDIF", fExpr)
  
  If IfCnt <> ThenCnt Then
    sShowMessage "IF-THEN-ELSE-ENDIF logic not constructed properly."
  Else
    If ThenCnt <> ElseCnt Then
      sShowMessage "IF-THEN-ELSE-ENDIF logic not constructed properly."
    Else
      If ElseCnt <> EndCnt Then
        sShowMessage "IF-THEN-ELSE-ENDIF logic not constructed properly."
      End If
    End If
  End If
  
  If IfCnt > 0 Then
    Do While Len(Expr) > 0
      'Find If Then Else Endif Block One by One
      iEnd = InStr(1, Expr, "} ENDIF", vbTextCompare)
      If iEnd >= 1 Then
        iStart = InStrRev(Expr, "} ELSE {", iEnd, vbTextCompare)
          If iStart >= 1 Then
            If CheckBracesIntegrity(Mid$(Expr, iStart, iEnd - iStart + Len("} ENDIF"))) = True Then
              Expr = Left$(Expr, iStart - 1) & Mid$(Expr, iEnd + Len("} ENDIF"))
            Else
              GoTo ExitHere
            End If
          End If
      Else
        If CheckBracesIntegrity(Expr) = True Then
          Expr = ""
        Else
          GoTo ExitHere
        End If
      End If
      
      
      iEnd = iStart - 1
      iStart = InStrRev(Expr, "} THEN {", iEnd, vbTextCompare)
      If iStart >= 1 Then
        If CheckBracesIntegrity(Mid$(Expr, iStart, iEnd - iStart + 1)) = True Then
          Expr = Left$(Expr, iStart - 1) & Mid$(Expr, iEnd + 1)
        Else
          GoTo ExitHere
        End If
      End If
  
      iEnd = iStart - 1
      iStart = InStrRev(Expr, "IF {", iEnd, vbTextCompare)
      If iStart >= 1 Then
        If CheckBracesIntegrity(Mid$(Expr, iStart, iEnd - iStart + 1)) = True Then
          Expr = Left$(Expr, iStart - 1) & Mid$(Expr, iEnd + 1)
        Else
          GoTo ExitHere
        End If
      End If
    Loop
  Else
    If CheckBracesIntegrity(Expr) = False Then
      GoTo ExitHere
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function CheckBracesIntegrity(ByVal cExpr As String) As Boolean
  Dim Expr As String
  Dim cBrace As Integer
  Dim cStart As Integer
  Dim cEnd As Integer
  Dim Cnt As Integer
  
  Dim OpCnt As Integer
  Dim ClCnt As Integer
  
  Try
  CheckBracesIntegrity = False
  
  Expr = cExpr
  
  OpCnt = GetStrCount("(", Expr)
  ClCnt = GetStrCount(")", Expr)
  
  If OpCnt <> ClCnt Then
    sShowMessage "Opening & Closing Paranthesis not equal."
    GoTo ExitHere
  Else
  
    If OpCnt = 0 Then
      If CheckPercBlockIntegrity(cExpr) = False Then
        GoTo ExitHere
      Else
        CheckBracesIntegrity = True
        GoTo ExitHere
      End If
    Else
      For Cnt = 1 To Len(Expr)
        Select Case Mid$(Expr, Cnt, 1)
          Case "("
            cBrace = cBrace + 1
          Case ")"
            cBrace = cBrace - 1
        End Select
      Next Cnt
      If cBrace < 0 Then
        sShowMessage "Opening & Closing Paranthesis not constructed properly in the Expression" & Chr(13) & cExpr
        GoTo ExitHere
      End If
    End If
  End If
  
  If cBrace <> 0 Then
    sShowMessage "Opening & Closing Paranthesis not equal in the Expression" & Chr(13) & cExpr
    GoTo ExitHere
  End If

  Do While Len(Expr) > 0
    cEnd = InStr(1, Expr, ")", vbTextCompare)
    If cEnd >= 1 Then
      cStart = InStrRev(Expr, "(", cEnd, vbTextCompare)
      If cStart >= 1 Then
        If CheckPercBlockIntegrity(Mid$(Expr, cStart, cEnd - cStart + 1)) = True Then
          Expr = Left$(Expr, cStart - 1) & Mid$(Expr, cEnd + 1)
        Else
          GoTo ExitHere
        End If
      End If
    Else
      If CheckPercBlockIntegrity(Expr) = True Then
        Expr = ""
      Else
        GoTo ExitHere
      End If
    End If
  Loop

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function CheckPercBlockIntegrity(ByVal pExpr As String) As Boolean
  Dim Expr As String
  Dim pBrace As Integer
  Dim Cnt As Integer

  Try
  CheckPercBlockIntegrity = False
  
  Expr = pExpr
  
  For Cnt = 1 To Len(Expr)
    Select Case Mid$(Expr, Cnt, 1)
      Case PercBlockOpen                'PercBlockOpen
        pBrace = pBrace + 1
      Case PercBlockClose                'PercBlockClose
        pBrace = pBrace - 1
    End Select
    
    If pBrace < 0 Then
      sShowMessage "Percentage Block " & PercBlockOpen & " & " & PercBlockClose & " not constructed properly in the Expression" & Chr(13) & pExpr
      GoTo ExitHere
    End If
  Next Cnt
  
  If pBrace <> 0 Then
    sShowMessage "Percentage Block " & PercBlockOpen & " & " & PercBlockClose & " not equal in the Expression" & Chr(13) & pExpr
    GoTo ExitHere
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function ClearTmpPromptTable(Optional ByVal fUserKey As String = "-1") As Boolean
  Dim Sql As String
  
  Try
  ClearTmpPromptTable = False
  
  Sql = "Delete from COM_tmpFormulaCalc_Dtl"
  Sql = Sql & " where UserID='" & gUserID & "' "
  Sql = Sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  If fUserKey <> "-1" Then
    Sql = Sql & " and DTKey='" & fUserKey & "'"
  End If
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
  Sql = "Delete from COM_tmpFormulaCalc_Hdr"
  Sql = Sql & " where UserID='" & gUserID & "' "
  Sql = Sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  If fUserKey <> "-1" Then
    Sql = Sql & " and DTKey='" & fUserKey & "'"
  End If
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function DeletePromptValue(ByVal fFormulaID As Integer,ByVal  fPrevUserKey As String,Optional ByVal SupplyThruCodeOnly As Boolean = True) As Boolean
  Dim Sql As String
  
  Try
  DeletePromptValue = False
  
  Sql = "Delete from COM_tmpFormulaCalc_Dtl"
  Sql = Sql & " Where DTKey='" & fPrevUserKey & "' and FrmID=" & fFormulaID & " and BlockType='Param'"
  Sql = Sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  If SupplyThruCodeOnly = True Then
    Sql = Sql & " and Convert(int, Block) in (Select b.ParamID from Com_FrmParamMast_Hdr b"
    Sql = Sql & " Where b.ValInputMethod='SC' and b.ModuleCode=" & gModuleCode & " and b.BranchId='" & gBranchID & "' and b.CompID='" & gCompID & "')"
  End If

  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function FASPosting(ByVal CurDate As Date,ByVal  VrType As eVrType,ByVal  VrRoute As String,ByVal  VrDate As Date,ByVal  PartyCode As Integer,ByVal  OrderPostingAc As Integer,ByVal  FrmID As Integer,ByRef  CalcSpread As AceSpread,ByVal  GrandTotalBox As TextBox,Optional ByVal isCalcDiff As Boolean = False,Optional ByVal TotalMistakeBox As TextBox,Optional ByVal RefType As String = "",Optional ByVal RefNo As Integer = 0,Optional ByVal RefDate As Date,Optional ByVal PartyBillRef As String = "",Optional ByVal PartyBillDate As Date,Optional ByVal BatchType As String = "",Optional ByVal BatchNo As Integer = 0,Optional ByVal BatchDate As Date,Optional ByVal StartRow As Long = 1,Optional ByVal DescCol As Long = 1,Optional ByVal ValCol As Long = 2,Optional ByVal fNarration As String = "") As Boolean






  Dim Sql As String
  Dim rsExpr As AceADODB.Recordset
  Dim retVal As Variant
  Dim rCnt As Long

  Dim VrNo As Integer
  Dim VTSNo As Integer
'  Dim VrRoute As String  'Comment on 14-03-2008
  Dim FileRef As Long

  Dim amt As Double
  Dim MultiplyFactor As Integer

  Try
  FASPosting = False

'  VrRoute = GetVrRoute(VrType) ' Comment on 14/03/2008
  VrNo = GetMaxVrNo(VrDate, Trim$(VrRoute)) + 1
  
  FileRef = GetMaxNo("FAS_VoucherTranMast", "FileRef", gFasCompCode, "VrRoute='" & Trim$(VrRoute) & "'") + 1

  Sql = "Execute FAS_Mod_VrTranReNum_From @StNo=" & VrNo & ", @VrDate = '" & Format(VrDate, "yyyy-MM-dd") & "', @VrRoute='" & Trim$(VrRoute) & "', @NumToAdd=1, @CmCode='" & gFasCompCode & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere

  'Voucher Tran Master
  Sql = "Execute FAS_Ins_VoucherTranMast"
  Sql = Sql & " @DocType='VT',"
  Sql = Sql & " @DocNo=" & VrNo & ","
  Sql = Sql & " @DocDate='" & Format(VrDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @VrRoute='" & Trim$(VrRoute) & "',"
'  Sql = Sql & " @Ref='" & Trim$(RefType) & ":" & Val(RefNo) & "',"
  Sql = Sql & " @Ref = '" & Trim$(PartyBillRef) & "/" & Format(PartyBillDate, "dd.MM.yyyy") & "',"
  Sql = Sql & " @FileRef=" & FileRef & ","
  Sql = Sql & " @Amount=" & Val(GrandTotalBox.Text) & ","
  Sql = Sql & " @PayMode='',"
  Sql = Sql & " @Narration='" & fNarration & "',"
  Sql = Sql & " @CompCode='" & gFasCompCode & "',"
  Sql = Sql & " @UserID='" & gUserID & "',"
  Sql = Sql & " @TranStamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere

  'Voucher Ref Table
  If RefNo <> 0 Then
    Sql = "Execute FAS_Ins_VoucherTranRef"
    Sql = Sql & " @DocType='VT',"
    Sql = Sql & " @DocNo=" & VrNo & ","
    Sql = Sql & " @DocDate='" & Format(VrDate, "yyyy-MM-dd") & "',"
    Sql = Sql & " @VrRoute='" & Trim$(VrRoute) & "',"
    Sql = Sql & " @RefType='" & Trim$(RefType) & "',"
    Sql = Sql & " @RefNo=" & RefNo & ","
    If RefNo > 0 Then
      Sql = Sql & " @RefDate='" & Format(RefDate, "yyyy-MM-dd") & "',"
    Else
      Sql = Sql & " @RefDate=Null,"
    End If
    Sql = Sql & " @BatchType='" & Trim$(BatchType) & "',"
    Sql = Sql & " @BatchNo=" & BatchNo & ","
    If BatchNo > 0 Then
      Sql = Sql & " @BatchDate='" & Format(BatchDate, "yyyy-MM-dd") & "',"
    Else
      Sql = Sql & " @BatchDate=Null,"
    End If
    Sql = Sql & " @CompCode='" & gFasCompCode & "',"
    Sql = Sql & " @UserID='" & gUserID & "',"
    Sql = Sql & " @TranStamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
    If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  End If

  'Voucher Bill Details
  Sql = "Execute FAS_Ins_BillDetails"
  Sql = Sql & " @DocType='VT',"
  Sql = Sql & " @DocNo=" & VrNo & ","
  Sql = Sql & " @DocDate='" & Format(VrDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @VTSno=1,"
  Sql = Sql & " @VrRoute='" & Trim$(VrRoute) & "',"
  Sql = Sql & " @LrCode=" & PartyCode & ","
  Sql = Sql & " @BillIndex=1,"
  Sql = Sql & " @RefNo='" & PartyBillRef & "/" & Format(PartyBillDate, "dd.MM.yyyy") & "',"
  Sql = Sql & " @RefDate='" & Format(PartyBillDate, "yyyy-MM-dd") & "',"

  Select Case VrType
    Case mSales
      Sql = Sql & " @Type='New Ref',"
      Sql = Sql & " @Amount=" & Val(GrandTotalBox.Text) & ","
    Case mPurchase
      Sql = Sql & " @Type='New Ref',"
      Sql = Sql & " @Amount=" & -Val(GrandTotalBox.Text) & ","
    Case mCredit
      Sql = Sql & " @Type='Agst Ref',"
      Sql = Sql & " @Amount=" & -Val(GrandTotalBox.Text) & ","
    Case mDebit
      Sql = Sql & " @Type='Agst Ref',"
      Sql = Sql & " @Amount=" & Val(GrandTotalBox.Text) & ","
  End Select

  Sql = Sql & " @CreditDays=0,"
  Sql = Sql & " @CompCode='" & Left$(gFasCompCode, 4) & "',"
  Sql = Sql & " @UserID='" & gUserID & "',"
  Sql = Sql & " @TranStamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere

  'Voucher Tran Sub
  'For Party Account
  Select Case VrType
    Case mSales, mDebit
      MultiplyFactor = -1
    Case mPurchase, mCredit
      MultiplyFactor = 1
  End Select

  amt = MultiplyFactor * Val(GrandTotalBox.Text)
  VTSNo = 1
  If SaveVoucherTranSub(CurDate, "VT", VrNo, VrDate, VrRoute, 1, PartyCode, Format(-amt, "0.00")) = False Then
    GoTo ExitHere
  End If

  Sql = "Select Case When Convert(int, a.ExprRoute) < 0 then 'SUBTOTAL' else b.ExprGroupName end as ExprGroupName, "
  Sql = Sql & " b.Route+b.Width as ExprRoute, a.GrandTotalFlag, a.PostingAc"
  Sql = Sql & " from COM_FormulaMast_Dtl a Left Outer Join COM_FormulaExpr_Group b"
  Sql = Sql & " On a.ExprRoute=b.Route+b.Width and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
  Sql = Sql & " Where a.FrmID=" & FrmID & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Sql = Sql & " Union All"
  Sql = Sql & " Select 'ROUNDOFF', '', 1, c.PostingAc from COM_FormulaMast_Hdr c"
  Sql = Sql & " Where c.FrmID=" & FrmID & " and c.ModuleCode=" & gModuleCode & " and c.BranchID='" & gBranchID & "' and c.CompID='" & gCompID & "'"
  Set rsExpr = New AceADODB.Recordset
  If RSOpen(rsExpr, Sql) = False Then
    GoTo ExitHere
  End If

  If Not rsExpr.EOF Then
      For rCnt = StartRow To CalcSpread.MaxRows
        CalcSpread.GetText eCalc.mDesc, rCnt, retVal
        rsExpr.Filter = adFilterNone
        rsExpr.Filter = "ExprGroupName='" & Trim$(retVal) & "'"
        If Not rsExpr.EOF Then
          If rsExpr("GrandTotalFlag") <> 0 Then
            CalcSpread.GetText ValCol, rCnt, retVal
            amt = Val(retVal) * MultiplyFactor * Abs(rsExpr("GrandTotalFlag"))

            VTSNo = VTSNo + 1
            If SaveVoucherTranSub(CurDate, "VT", VrNo, VrDate, VrRoute, VTSNo, IIf(rsExpr("PostingAc") <= 0, OrderPostingAc, rsExpr("PostingAc")), Format(amt, "0.00")) = False Then
              GoTo ExitHere
            End If
          End If
        End If
      Next rCnt
    
    If isCalcDiff = True Then
      If Val(TotalMistakeBox.Text) <> 0 Then
        VTSNo = VTSNo + 1
        If SaveVoucherTranSub(CurDate, "VT", VrNo, VrDate, VrRoute, VTSNo, OrderPostingAc, Format(amt, "0.00")) = False Then
          GoTo ExitHere
        End If
      End If
    End If
  End If

  If FAS_Posting_Validate_After_Save(VrNo, VrDate, VrRoute) = False Then
    GoTo ExitHere
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsExpr = Nothing
End Try
End Function

Public Function GetExprResults(ByVal fFormulaID As Integer,ByVal  fUserKey As String,ByRef  fResultSpread As AceSpread,ByVal  fGrandTotalTextBox As TextBox,ByVal  fDescCol As Long,Optional ByVal fValueCol As Long = -2,Optional ByVal fStartRow As Long = 1,Optional ByVal RestrictZeroValue As Boolean = True) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim CurRow As Long
  Dim mEndRow As Long
  Dim ValCol As Long
  Dim tmpValue As Double
  
  Try
  GetExprResults = False
  
'  Sql = "Select a.ExprSeq, isNull(b.ExprName, 'ROUNDOFF') as ExprName, a.Result, a.GrandTotalFlag from COM_tmpFormulaCalc_Hdr a "
'  Sql = Sql & " Left outer join (Select c.FrmID, c.ExprSeq, d.ExprGroupName as ExprName, c.BranchID, c.CompID, c.ModuleCode"
'  Sql = Sql & " from COM_FORMULAMAST_DTL c Left Outer Join COM_FormulaExpr_Group d"
'  Sql = Sql & " On c.ExprRoute=d.Route+d.Width and c.ModuleCode=d.ModuleCode and c.BranchID=d.BranchID and c.CompID=d.CompID) b"
'  Sql = Sql & " On a.FrmId=b.FrmId and a.ExprSeq=b.ExprSeq and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
'  Sql = Sql & " Where a.DTKey='" & fUserKey & "' and a.FrmID=" & fFormulaID & " and a.UserID='" & gUserID & "'"
'  Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and a.finyear='" & gFinyear & "'"
'  Sql = Sql & " Order by a.ExprSeq"
  
  Sql = "Select * from dbo.fn_COM_GetExprResults_Temp('" & fUserKey & "', " & gModuleCode & ", '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')"
  Sql = Sql & " Order by ExprSeq"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    If fValueCol = -2 Then
      ValCol = fDescCol + 1
    Else
      ValCol = fValueCol
    End If
  
    rs.MoveLast
rs.MoveFirst
      mEndRow = fStartRow + rs.RecordCount - 1
      If fResultSpread.MaxRows < mEndRow Then
        fResultSpread.MaxRows = mEndRow
      End If
      
      Call Spread_CellHorAlign_Set(fResultSpread, TypeHAlignRight, fDescCol + 1, fStartRow, , mEndRow)
      
      CurRow = fStartRow - 1
      For rCnt = fStartRow To mEndRow
        If RestrictZeroValue = True Then
          tmpValue = rs("ResultValue") * rs("GrandTotalFlag")
          If tmpValue <> 0 Then
            CurRow = CurRow + 1
            fResultSpread.SetText fDescCol, CurRow, rs("ExprName")
            fResultSpread.SetText ValCol, CurRow, Format(tmpValue, "0.00")
'          Else
'            .ClearRange fDescCol, CurRow, ValCol, CurRow, True
          End If
        Else
          CurRow = CurRow + 1
          fResultSpread.SetText fDescCol, CurRow, rs("ExprName")
          fResultSpread.SetText ValCol, CurRow, Format(rs("ResultValue"), "0.00")
        End If
        rs.MoveNext
      Next rCnt
      
      fResultSpread.MaxRows = fResultSpread.DataRowCnt
      
      'To Calculate & Display the Grand Total
      Sql = "Select Sum(a.ResultValue*a.GrandTotalFlag) as GrandTotal from dbo.fn_COM_GetExprResults_Temp('" & fUserKey & "', " & gModuleCode & ", '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "') a"
      Sql = Sql & " Where a.FrmID=" & fFormulaID & ""
      
'      Sql = "Select Sum(a.Result*a.GrandTotalFlag) as GrandTotal from COM_tmpFormulaCalc_Hdr a "
'      Sql = Sql & " Where a.DTKey='" & fUserKey & "' and a.FrmID=" & fFormulaID & " and a.UserID='" & gUserID & "'"
'      Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and a.finyear='" & gFinyear & "'"
      If RSOpen(rs, Sql) = False Then
        GoTo ExitHere
      End If
      
      If Not rs.EOF Then
        fGrandTotalTextBox.Text = Format(rs("GrandTotal"), "0.00")
      End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function GetFrontEndExpr(ByVal FormulaID As Integer,ByVal  ExprSeq As Integer) As String
  Dim rs As AceADODB.Recordset
  Dim Sql As String

  Try
  
  GetFrontEndExpr = ""
  
  Sql = "Select dbo.fn_FrontEndExpression(" & gModuleCode & ", " & FormulaID & ", " & ExprSeq & ", '" & gBranchID & "', '" & gCompID & "') as Expr"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    GetFrontEndExpr = rs("Expr")
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function GetRoundingOffCode(ByVal RoundingOffText As String) As String
' On Error Resume Next
  
  GetRoundingOffCode = "N"
  
  Select Case RoundingOffText
    Case "To Up"
      GetRoundingOffCode = "U"
    Case "To Down"
      GetRoundingOffCode = "D"
    Case "Nearest"
      GetRoundingOffCode = "B"
    Case Else
      GetRoundingOffCode = "N"
  End Select
  
End Function

Public Function GetRoundingOffText(ByVal RoundingOffCode As String) As String
' On Error Resume Next
  
  GetRoundingOffText = "None"
  
  Select Case GetRoundingOffText
    Case "U"
      GetRoundingOffText = "To Up"
    Case "D"
      GetRoundingOffText = "To Down"
    Case "B"
      GetRoundingOffText = "Nearest"
    Case Else
      GetRoundingOffText = "None"
  End Select
  
End Function

Public Function InsertCodePromptValue(ByVal fFormulaID As Integer,ByVal  fUserKey As String,ByVal  ParamSName_Value() As String) As Boolean
  'ParamSName_Value should be supplied in the format <ParamShortName, ValueToBeUsed, ...>
  Dim ArrCnt As Integer
  Dim Res() As String
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim CurMaxSeq As Integer
  Dim CurDate As Date
  
  Try
  InsertCodePromptValue = False
  
  CurDate = ServerDate
  
  If UBound(ParamSName_Value) < 0 Then
    sShowMessage "No Parameter & Values supplied to process"
    GoTo ExitHere
  Else
    For ArrCnt = 0 To UBound(ParamSName_Value)
      Res = Split(ParamSName_Value(ArrCnt), Chr(145), 3, vbTextCompare)
      If InsertPromptValues(fUserKey, fFormulaID, "Param", Val(Res(0)), Trim$(Res(1)), False, False, Val(Res(2)), False) = False Then
        GoTo ExitHere
      End If
    Next ArrCnt
    
    'Validate HSN No
    If ValidateHSNNoAvailability(CurDate, fUserKey, fFormulaID, Val(Res(0))) = False Then
      GoTo ExitHere
    End If
    
    'Inserting HSN Based Rate Details in Tmp Table
    If InsertHSNBasedTaxRate(CurDate, fUserKey, fFormulaID, Val(Res(0))) = False Then
      sShowMessage "HSN Details not inserted properly."
      GoTo ExitHere
    End If
    
    Sql = "Select a.Block, Max(MultiSeq) SeqMax"
    Sql = Sql & " from COM_tmpFormulaCalc_Dtl a Left Outer Join COM_FrmParamMast_Hdr b"
    Sql = Sql & " On a.ModuleCode=b.ModuleCode and a.Block=b.ParamID and a.BranchID=b.BranchID and a.CompID=b.CompID"
    Sql = Sql & " Where b.isMultiple='1' and a.DTKey='" & fUserKey & "' and a.FrmID=" & fFormulaID
    Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    Sql = Sql & " Group by a.Block"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, Sql) = False Then
      GoTo ExitHere
    End If
    If Not rs.EOF Then
      CurMaxSeq = rs("SeqMax")
      rs.MoveNext
      Do While Not rs.EOF
        If CurMaxSeq <> rs("SeqMax") Then
          sShowMessage "Supply thru code value not supplied for all the items."
          Call DeletePromptValue(fFormulaID, fUserKey)
        End If
        rs.MoveNext
      Loop
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function InsertHSNBasedTaxRate(ByVal fCurDate As Date,ByVal  fUserKey As String,ByVal  fFormulaID As Integer,ByVal  fParamID As Integer)
  Dim Sql As String
  Dim CurDate As Date
  
  Try
  
  InsertHSNBasedTaxRate = False
  
  Sql = " Execute COM_tmpFormulaCalcDtl_Sel_Ins_HSNBasedTaxRate"
  Sql = Sql & " @UserKey = '" & fUserKey & "',"
  Sql = Sql & " @ModuleCode = " & IIf(gModuleCode = eMdlCode.mPPC, eMdlCode.mPUR, gModuleCode) & ","
  Sql = Sql & " @FrmID = " & fFormulaID & ","
  Sql = Sql & " @ParamID = " & fParamID & ","
  Sql = Sql & " @IGSTParamID = " & eParamID.mIGST & ","
  Sql = Sql & " @CGSTParamID = " & eParamID.mCGST & ","
  Sql = Sql & " @SGSTParamID = " & eParamID.mSGST & ","
  Sql = Sql & " @AsonDate = '" & Format(fCurDate, "yyyy-MM-dd HH:mm:ss") & "',"
  Sql = Sql & " @RptDate = '" & Format(fCurDate, "yyyy-MM-dd HH:mm:ss") & "',"
  Sql = Sql & " @BranchID = '" & gBranchID & "',"
  Sql = Sql & " @CompID = '" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then GoTo ExitHere

  InsertHSNBasedTaxRate = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateHSNNoAvailability(ByVal fCurDate As Date,ByVal  fUserKey As String,ByVal  fFormulaID As Integer,ByVal  fParamID As Integer) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset

  Try

  ValidateHSNNoAvailability = False
  
  Sql = " Execute COM_Val_HSNNoBasedOnGSTParamFormula"
  Sql = Sql & " @IsGSTRequired = '" & IIf(IsGSTRequired, "1", "0") & "',"
  Sql = Sql & " @UserKey = '" & fUserKey & "',"
  Sql = Sql & " @ModuleCode = " & IIf(gModuleCode = eMdlCode.mPPC, eMdlCode.mPUR, gModuleCode) & ","
  Sql = Sql & " @FrmID = " & fFormulaID & ","
  Sql = Sql & " @ParamID = " & fParamID & ","
  Sql = Sql & " @IGSTParamID = " & eParamID.mIGST & ","
  Sql = Sql & " @CGSTParamID = " & eParamID.mCGST & ","
  Sql = Sql & " @SGSTParamID = " & eParamID.mSGST & ","
  Sql = Sql & " @BranchID = '" & gBranchID & "',"
  Sql = Sql & " @CompID = '" & gCompID & "'"
  If RSOpen(rs, Sql) = False Then
 GoTo ExitHere
End If
  
  If Not rs.EOF Then
    sShowMessage "HSN No is not available for the Items '" & Trim$(rs("Item")) & "' in Item Master. So, update the HSN No in Item Master."
    GoTo ExitHere
  End If
  
  ValidateHSNNoAvailability = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function InsertPromptValues(ByVal fUserKey As String,ByVal  fFormulaID As Integer,ByVal  fBlockType As String,ByVal  fBlock As String,ByVal  fKeyCol As String,ByVal  fIsPerc As Boolean,ByVal  fIsTypeEdit As Boolean,ByVal  FValue As Double,ByVal  fIsValueEdit As Boolean) As Boolean
  Dim Sql As String
  
  Try
  InsertPromptValues = False
  
  Sql = "Execute Com_Ins_tmpFormulaCalc_Dtl"
  Sql = Sql & " @DTKey='" & fUserKey & "',"
  Sql = Sql & " @ModuleCode=" & gModuleCode & ","
  Sql = Sql & " @FrmID=" & fFormulaID & ","
  Sql = Sql & " @BlockType='" & fBlockType & "',"
  Sql = Sql & " @Block='" & fBlock & "',"
  Sql = Sql & " @KeyCol='" & fKeyCol & "',"
  Sql = Sql & " @isTypePercentage='" & IIf(fIsPerc, 1, 0) & "',"
  Sql = Sql & " @isTypeEditable='" & IIf(fIsTypeEdit, 1, 0) & "',"
  Sql = Sql & " @ValueUsed=" & FValue & ","
  Sql = Sql & " @isValueEditable='" & IIf(fIsValueEdit, 1, 0) & "',"
  Sql = Sql & " @UserID='" & gUserID & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "',"
  Sql = Sql & " @CompID='" & gCompID & "',"
  Sql = Sql & " @Finyear='" & gFinyear & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

'Whether the particular formula requires the itemwise parameter values
Public Function MultiSeqParamFound(ByVal fFormulaID As Integer) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  MultiSeqParamFound = False
  
  Sql = "Select Count(distinct Block) from COM_FormulaMast_Block_Dtl a"
  Sql = Sql & " where a.BlockType='Param' and a.frmID=" & fFormulaID
  Sql = Sql & " and Convert(int, a.Block) in (select ParamID from COM_FrmParamMast_Hdr b "
  Sql = Sql & " where b.ModuleCode=" & gModuleCode & " and b.isMultiple='1' and b.BranchID='" & gBranchID & "' and b.CompID='" & gCompID & "')"
  Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rs.EOF Then
    If rs(0) > 0 Then
      MultiSeqParamFound = True
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Public Function SaveCalcDetails(ByVal fUserKey As String,ByVal  fModuleCode As Integer,ByVal  fFrmID As Integer,ByVal  fDocType As String,ByVal  fDocNo As Integer,ByVal  fDocDate As Date) As Boolean
  Dim Sql As String

  Try
  SaveCalcDetails = False

'  Sql = "Execute COM_SaveCalcDetails '" & fUserKey & "', " & fModuleCode &", "
'  Sql = Sql & " " & fFrmID & ", '" & fDocType & "', " & fDocNo & ", '" & Format(fDocDate, "yyyy-MM-dd") & "',"
'  Sql = Sql & " '" & gUserID & "', '" & gBranchID & "', '" & gCompID & "','" & gFinyear & "'"
'  db.Execute Sql, rCnt
'
'  If rCnt = 0 Then
'    GoTo ExitHere
'  End If

  Sql = " Delete from COM_FormulaCalc_Hdr"
  Sql = Sql & " Where DocType='" & fDocType & "' and DocNo=" & fDocNo & " and DocDate='" & Format(fDocDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and ModuleCode=" & fModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "' and Finyear='" & gFinyear & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere

  Sql = " Delete from COM_FormulaCalc_Dtl"
  Sql = Sql & " Where DocType='" & fDocType & "' and DocNo=" & fDocNo & " and DocDate='" & Format(fDocDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and ModuleCode=" & fModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "' and Finyear='" & gFinyear & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere

  Sql = " Insert into COM_FormulaCalc_Hdr"
  Sql = Sql & " Select '" & fDocType & "', " & fDocNo & ", '" & Format(fDocDate, "yyyy-MM-dd") & "', " & fModuleCode & ", " & fFrmID & ", ExprSeq, GrandTotalFlag, ResultValue, PostingAc, '','" & gUserID & "','" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "','" & Format(ServerDate, "yyyy-MM-dd HH:mm:dd") & "'"
  Sql = Sql & " from dbo.fn_COM_GetExprResults_Temp('" & fUserKey & "', " & fModuleCode & ", '" & gBranchID & "', '" & gCompID & "', '" & gFinyear & "')"
  Sql = Sql & " Order by ExprSeq"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere

  Sql = " Insert into COM_FORMULACALC_DTL"
  Sql = Sql & " Select '" & fDocType & "', " & fDocNo & ", '" & Format(fDocDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " a.ModuleCode , a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, a.IsTypePercentage, a.IsValueEditable, a.ValueUsed, a.IsValueEditable,'', a.UserID, a.BranchID, a.CompID, a.Finyear,'" & Format(ServerDate, "yyyy-MM-dd HH:mm:dd") & "'"
  Sql = Sql & " from COM_tmpFormulaCalc_Dtl a "
  Sql = Sql & " Where a.DTkey='" & fUserKey & "' and a.UserID='" & gUserID & "'"
  Sql = Sql & " and a.ModuleCode=" & fModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and a.Finyear='" & gFinyear & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere

  SaveCalcDetails = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function SaveExpressionValuesToTemp(ByVal fUserKey As String,ByVal  fFormulaID As Integer,ByVal  fExprSeq As Integer,ByVal  fGrandTotalFlag As Integer,ByVal  fResult As Currency) As Boolean
  Dim Sql As String

  Try
  SaveExpressionValuesToTemp = False
  
  Sql = "Execute COM_Ins_tmpFormulaCalc_Hdr"
  Sql = Sql & " @DTKey='" & fUserKey & "',"
  Sql = Sql & " @ModuleCode=" & gModuleCode & ","
  Sql = Sql & " @FrmId=" & fFormulaID & ","
  Sql = Sql & " @ExprSeq=" & fExprSeq & ","
  Sql = Sql & " @GrandTotalFlag=" & fGrandTotalFlag & ","
  Sql = Sql & " @Result=" & fResult & ","
  Sql = Sql & " @UserID='" & gUserID & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "',"
  Sql = Sql & " @CompID='" & gCompID & "',"
  Sql = Sql & " @Finyear='" & gFinyear & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function UpdateFinyearToCalcDetails(ByVal fUserKey As String,ByVal  fModuleCode As Integer,ByVal  fFrmID As Integer,ByVal  fDocType As String,ByVal  fDocNo As Integer,ByVal  fDocDate As Date,ByVal  OrdFinyear As String) As Boolean
  Dim Sql As String
  
  Try
  
  UpdateFinyearToCalcDetails = False
  
  Sql = " Update COM_FormulaCalc_Hdr Set Finyear = '" & OrdFinyear & "'"
  Sql = Sql & " Where DocType = '" & fDocType & "' and DocNo = " & fDocNo & " and DocDate = '" & Format(fDocDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and ModuleCode=" & fModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Sql = Sql & " and Finyear = '" & gFinyear & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
  Sql = " Update COM_FORMULACALC_DTL Set Finyear = '" & OrdFinyear & "'"
  Sql = Sql & " Where DocType = '" & fDocType & "' and DocNo = " & fDocNo & " and DocDate = '" & Format(fDocDate, "yyyy-MM-dd") & "'"
  Sql = Sql & " and ModuleCode=" & fModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Sql = Sql & " and Finyear = '" & gFinyear & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
    
  UpdateFinyearToCalcDetails = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function ValidateExpression(ByVal InExpr As String) As Boolean
  Dim Expr As String

  Try
  ValidateExpression = False
  
  Expr = InExpr

  If CheckBlockIntegrity(InExpr) = False Then
    GoTo ExitHere
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Public Function ValidateSupplyThruCodeValue(ByVal SCValue() As String) As Boolean
  Dim Cnt As Integer
  Dim pCnt As Integer
  Dim ScVal As String
  Dim ChkParam() As String

  Try
  ValidateSupplyThruCodeValue = False
  
  If UBound(SCValue) = -1 Then
    sShowMessage "Values not passed for the parameter type Supplied Thru Code"
    GoTo ExitHere
  End If
  
  For Cnt = 0 To UBound(SCValue)
    Erase ChkParam
    ScVal = SCValue(Cnt)
    ChkParam = Split(ScVal, Chr(145), 3, vbTextCompare)
    If UBound(ChkParam) < 2 Then
      sShowMessage "Error in " & Cnt & nthString(Cnt) & " value in the list. Supply thru Code Values not passed properly."
      GoTo ExitHere
    Else
      For pCnt = 0 To 2
        If pCnt <> 1 Then
          If Not IsNumeric(ChkParam(pCnt)) Then
            sShowMessage "Error in " & Cnt & nthString(Cnt) & " value in the list. Numeric values only allowed for the SupplyThruCode Value Argument."
            GoTo ExitHere
          End If
        End If
      Next pCnt
    End If
  Next Cnt

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function CopyPromptValue(ByVal fFormulaID As Integer,ByVal  FromUserKey As String,ByVal  ToUserKey As String,Optional ByVal SupplyThruCodeInclude As Boolean = False,Optional ByVal isSkipFixedValue As Boolean = False) As Boolean

  Dim Sql As String
  
  Try
  CopyPromptValue = False
  
  If SupplyThruCodeInclude = True Then
    Sql = "Delete from COM_tmpFormulaCalc_Dtl"
    Sql = Sql & " Where Where DTKey='" & ToUserKey & "' and UserID='" & gUserID & "' and BranchID='" & gBranchID & "' and CompID='" & gCompID & "' and Finyear='" & gFinyear & "'"
    If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  End If
  
  Sql = "Insert into COM_tmpFormulaCalc_Dtl"
  Sql = Sql & " Select '" & ToUserKey & "', a.ModuleCode, a.FrmID, a.BlockType, a.Block, a.MultiSeq, a.KeyCol, a.isTypePercentage, a.isTypeEditable,"
  Sql = Sql & " Case When " & IIf(isSkipFixedValue = True, 1, 0) & "=1 and IsTypePercentage='0' Then 0 Else a.ValueUsed End as ValueUsed,"
  Sql = Sql & " a.isValueEditable, a.UserID, a.BranchID, a.CompID, a.Finyear"
  Sql = Sql & " from COM_tmpFormulaCalc_Dtl a"
  Sql = Sql & " Where a.DTKey='" & FromUserKey & "' and a.UserID='" & gUserID & "' and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and a.Finyear='" & gFinyear & "'"
  If SupplyThruCodeInclude = False Then
    Sql = Sql & " and Convert(int, a.Block) not in (Select ParamID from COM_FrmParamMast_Hdr"
    Sql = Sql & " Where ValInputMethod='SC' and BranchID='" & gBranchID & "' and CompID='" & gCompID & "')"
  End If
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function GetVrRoute(ByVal VrType As eVrType) As String

  Try

  Select Case VrType
    Case mSales
      GetVrRoute = "0K"
    Case mPurchase
      GetVrRoute = "0B"
    Case mCredit
      GetVrRoute = "02"
    Case mDebit
      GetVrRoute = "03"
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function InsertAlreadyLockedValues(ByVal fFormulaID As Integer,ByVal  fUserKey As String,ByVal  LockRefType As String,ByVal  LockRefNo As Integer,ByVal  LockRefDate As Date,Optional ByVal RefUserKey As String) As Boolean
  Dim Sql As String

  Try
  
  InsertAlreadyLockedValues = False
  
  Sql = " Insert into COM_tmpFormulaCalc_Dtl"
  Sql = Sql & " Select '" & fUserKey & "', " & gModuleCode & ", " & fFormulaID & ", 'Param', f1.Block,0 as Multiseq, isNull(f2.KeyCol, f1.KeyCol) as KeyCol,"
  Sql = Sql & " isNull(f2.isTypePercentage, f1.isTypePercentage) as isTypePercentage, isNull(f1.isTypeEditable, f1.isTypeEditable) as isTypeEditable,"
  Sql = Sql & " isNull(f2.ValueUsed, f1.ValueUsed) as ValueUsed, isNull(f2.isValueEditable, f1.isValueEditable) as isValueEditable,"
  Sql = Sql & " '" & gUserID & "', f1.BranchID, f1.CompID, '" & gFinyear & "'"
  Sql = Sql & " From (Select ModuleCode, FrmID, Block, '' as KeyCol, isTypePercentage, isTypeEditable, ValueUsed, isValueEditable, BranchID, CompID"
  Sql = Sql & " From COM_FormulaMast_Block_Dtl"
  Sql = Sql & " Where FrmID=" & fFormulaID & " and BlockType='Param' and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "' ) f1"
  Sql = Sql & " Left Outer Join"
  Sql = Sql & " (Select ModuleCode, FrmID, Block, KeyCol, isTypePercentage, isTypeEditable, ValueUsed, isValueEditable"
  Sql = Sql & " From COM_FormulaCalc_Dtl a"
  
'  Sql = Sql & " Where DocNo=" & LockRefNo & " and DocType='" & LockRefType & "' and DocDate='" & Format(LockRefDate, "yyyy-MM-dd") & "'"
  
  Sql = Sql & " Where Exists (Select * From Com_TmpFormulaCalc_Ref Where UserKey = '" & Trim$(RefUserKey) & "' "
  Sql = Sql & " and a.DocDate = DocDate and a.DocNo = DocNo and a.DocType = DocType"
  Sql = Sql & " and ModuleCode = " & gModuleCode & " and isLockType = '1' and BranchID = '" & gBranchID & "' and CompID = '" & gCompID & "' and Finyear = '" & gFinyear & "' )"
  Sql = Sql & " and BlockType='Param' and isValueEditable='0' and FrmID=" & fFormulaID & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "' "
  Sql = Sql & " and Exists (Select * from COM_tmpFormulaCalc_Dtl c"
  Sql = Sql & " Where c.DTKey = '" & fUserKey & "'  and c.KeyCol = a.KeyCol and c.ModuleCode = a.ModuleCode "
  Sql = Sql & " and c.FrmID=a.FrmID and c.BlockType=a.BlockType and c.Block=a.Block"
  Sql = Sql & " and c.BranchID = '" & gBranchID & "' and c.CompID = '" & gCompID & "' and c.Finyear = '" & gFinyear & "'"
  Sql = Sql & " )"
  Sql = Sql & " ) f2"
  Sql = Sql & " On f1.Block=f2.Block"
  Sql = Sql & " Left Outer Join COM_FrmParamMast_Hdr f3"
  Sql = Sql & " On f1.Block=f3.ParamID and f1.ModuleCode=f3.ModuleCode and f1.BranchID=f3.BranchID and f1.CompID=f3.CompID"
  Sql = Sql & " Where f3.ValInputMethod<>'SC' and isNull(f2.isValueEditable, f1.isValueEditable)='0'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
  InsertAlreadyLockedValues = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function SaveVoucherTranSub(ByVal CurDate As Date,ByVal  DocType As String,ByVal  DocNo As Integer,ByVal  DocDate As Date,ByVal  VrRoute As String,ByVal  VTSNo As Integer,ByVal  LrCode As Integer,ByVal  Amount As Double) As Boolean

  Dim Sql As String

  Try
  SaveVoucherTranSub = False
  
  If LrCode = 0 Or DocNo = 0 Then
    sShowMessage "Error While Posting Voucher"
    GoTo ExitHere
  End If

  Sql = "Execute FAS_InsUp_VoucherTranSub"
  Sql = Sql & " @DocType='" & Trim$(DocType) & "',"
  Sql = Sql & " @DocNo=" & DocNo & ","
  Sql = Sql & " @DocDate='" & Format(DocDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @VTSNo=" & VTSNo & ","
  Sql = Sql & " @VrRoute='" & Trim$(VrRoute) & "',"
  Sql = Sql & " @LrCode=" & LrCode & ","
  If Amount > 0 Then
    Sql = Sql & " @DrAmt=" & Format(Amount, "0.00") & ","
    Sql = Sql & " @CrAmt=0.00,"
  Else
    Sql = Sql & " @DrAmt=0.00,"
    Sql = Sql & " @CrAmt=" & Format(-Amount, "0.00") & ","
  End If
  Sql = Sql & " @CompCode='" & gFasCompCode & "',"
  Sql = Sql & " @UserID='" & gUserID & "',"
  Sql = Sql & " @TranStamp='" & Format(CurDate, "yyyy-MM-dd HH:mm:ss") & "'"
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function UpdateUserKey(ByVal fFormulaID As Integer,ByVal  fUserKey As String,ByVal  fPrevUserKey As String,Optional ByVal IncludeSupplyThruCode As Boolean = False) As Boolean
  Dim Sql As String
  
  Try
  UpdateUserKey = False
  
  Sql = "Update COM_tmpFormulaCalc_Dtl"
  Sql = Sql & " Set DTKey='" & fUserKey & "'"
  Sql = Sql & " from COM_tmpFormulaCalc_Dtl a Left Outer Join COM_FrmParamMast_Hdr b"
  Sql = Sql & " On Convert(int, a.Block) = b.ParamID and a.ModuleCode=b.ModuleCode and a.BranchID=b.BranchID and a.CompID=b.CompID"
  Sql = Sql & " Where a.DTKey='" & fPrevUserKey & "' and a.FrmID=" & fFormulaID
  Sql = Sql & " and a.UserID='" & gUserID & "' and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "' and a.Finyear='" & gFinyear & "'"
  If IncludeSupplyThruCode = False Then
    Sql = Sql & " and b.ValInputMethod <> 'SC'"
  End If
  If DBExecuteSql(db, sql) = False Then GoTo ExitHere
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

