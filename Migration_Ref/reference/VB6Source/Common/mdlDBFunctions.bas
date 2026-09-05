Attribute VB_Name = "mdlDBFunctions"
Option Explicit

Public Function DBCmdExecute(ByRef dbCon As AceADODB.Connection, ByRef fCmd As AceADODB.Command,Optional ByVal ErrorObj As ClsLastTranErrorDtl = Null) As Boolean
  'Execute the command object
  'For efficient connection pooling open the connection object when needed and close after it's use ceased.
  'Written by V.Sridevi on 09/09/2011 Modified by Sujatha for ErrorObj on 2019/06/18
    
  Try
  DBCmdExecute = False
  
  If dbCon Is Nothing Then Set dbCon = db
  If dbCon.State = adStateClosed Then
 dbCon.Open
End If
  
  fCmd.ActiveConnection = dbCon
  fCmd.Execute

  If gTransactionCount <= 0 Then dbCon.Close

  DBCmdExecute = True
Catch ex As Exception
  Select Case Err.Number
    Case 3246       'db Connection is in Middle of Transaction.
         'So no need to close
    Case Else
    
      If Not ErrorObj Is Nothing Then
        ErrorObj.ErrorNo = Err.Number
        ErrorObj.ErrorDesc = Err.Description
        ErrorObj.ErrorTime = Now()
      End If
      
      Call ErrMsgShow(ex)

  End Select
Finally

End Try
End Function

Public Function DBCommitTransaction(Optional ByVal dbCon As AceADODB.Connection) As Boolean

  Try
  DBCommitTransaction = False
    
  If dbCon Is Nothing Then Set dbCon = db
  
  If gTransactionCount > 0 Then
    dbCon.CommitTrans
    DBCommitTransaction = True
    dbCon.Close

' On Error Resume Next
    gTransactionCount = gTransactionCount - 1
  Else
' On Error Resume Next
    sShowMessage "No Active Transaction found to commit."
  End If
  
Catch ex As Exception
  sShowMessage "Error while committing the transaction." & Chr(13) & Chr(10) & "[" & Err.Number & "] " & Err.Description, mCritical

Finally

End Try
End Function

Public Function DBExecuteSql(ByRef dbCon As AceADODB.Connection, ByVal fSql As String,Optional ByVal ByRef OutParam1 As Variant,Optional ByVal ErrorObj As ClsLastTranErrorDtl = Null) As Boolean
  'Execute the Sql by using connection object
  'For efficient connection pooling open the connection object when needed and close after it's use ceased.
  'Written by V.Sridevi on 09/09/2011 Modified by Sujatha for ErrorObj on 2019/06/18
  
  Dim UbndOut As Integer
  
  Try
  DBExecuteSql = False
  
  If IsMissing(OutParam1) Then
    UbndOut = 0
  Else
    UbndOut = 1
  End If

  If dbCon Is Nothing Then Set dbCon = db
  If dbCon.State = adStateClosed Then
 dbCon.Open
End If
  
  If UbndOut = 0 Then
    dbCon.Execute fSql
  Else
    dbCon.Execute fSql, OutParam1
  End If
  
  If gTransactionCount <= 0 Then dbCon.Close
  
  DBExecuteSql = True
Catch ex As Exception
  Select Case Err.Number
    Case 3246   'db Connection is in Middle of Transaction. So no need to close
      
    Case Else
      If Not ErrorObj Is Nothing Then
        ErrorObj.ErrorNo = Err.Number
        ErrorObj.ErrorDesc = Err.Description
        ErrorObj.ErrorTime = Now()
      End If
      Call ErrMsgShow(ex)

  End Select
Finally

End Try
End Function

Public Function DBRollBackTransaction(Optional ByVal ByRef dbCon As AceADODB.Connection) As Boolean
  Dim ErrNo As Long
  Dim ErrDesc As String
  
  DBRollBackTransaction = False
  
  If Err.Number <> 0 Then
    ErrNo = Err.Number
    ErrDesc = Err.Description
    Err.Clear
  End If
  
  Try
  
  If dbCon Is Nothing Then Set dbCon = db
  
  If gTransactionCount > 0 Then
    dbCon.RollbackTrans
    
' On Error Resume Next
    gTransactionCount = gTransactionCount - 1
  Else
' On Error Resume Next
    sShowMessage "No Active Transaction found to rollback."
  End If
  
Catch ex As Exception
  If Err.Number <> 0 Then
    ErrNo = Err.Number
    ErrDesc = Err.Description
  End If
  sShowMessage "Error while rollback the transaction." & Chr(13) & Chr(10) & "[" & ErrNo & "] " & ErrDesc, mCritical

Finally

End Try
End Function

Public Function DBRollBackWithMessage(Optional ByVal MsgType As eSaveSkipMsgType = eSaveSkipMsgType.mNone,Optional ByVal MsgStr As String = "",Optional ByVal MsgBoxIcon As eMsgBoxIcon = eMsgBoxIcon.mInformation,Optional ByVal ByRef dbCon As AceADODB.Connection) As Boolean

  Dim tmpMsg As String
  Try
  DBRollBackWithMessage = False
  
  
  If dbCon Is Nothing Then Set dbCon = db
  
  
  tmpMsg = IIf(Trim$(MsgStr) = "", "Error Occurred.", MsgStr)
  
  If gTransactionCount > 0 Then
    dbCon.RollbackTrans
    
' On Error Resume Next
    gTransactionCount = gTransactionCount - 1
  End If
  
' On Error Resume Next
  
  Select Case MsgType
    Case mCustom
      sShowMessage tmpMsg, MsgBoxIcon
    Case mNoChanges
      sShowMessage "No Data Modified to Save.", MsgBoxIcon
    Case mHoldDelDoc
      sShowMessage "This document could not be Modified.", MsgBoxIcon
    Case Else
      If Trim$(MsgStr) <> "" Then
        sShowMessage tmpMsg, MsgBoxIcon
      Else
        'Continue
      End If
  End Select

Catch ex As Exception
  sShowMessage "Error while rollback the transaction." & Chr(13) & Chr(10) & "[" & Err.Number & "] " & Err.Description

Finally

End Try
End Function

Public Function DBStartTransaction(Optional ByVal ByRef dbCon As AceADODB.Connection) As Boolean

  Try
  DBStartTransaction = False
  
  If dbCon Is Nothing Then Set dbCon = db
    
  If gTransactionCount = 0 Then
    If dbCon.State = adStateClosed Then dbCon.Open
    
    dbCon.BeginTrans
    DBStartTransaction = True
    
' On Error Resume Next
    gTransactionCount = gTransactionCount + 1
  Else
' On Error Resume Next
    sShowMessage "No more Transaction allowed to begin while a Transaction is Active."
  End If
Catch ex As Exception
  sShowMessage "Error while beginning the transaction." & Chr(13) & Chr(10) & "[" & Err.Number & "] " & Err.Description, mCritical

Finally

End Try
End Function

Public Function RSOpen(ByRef RecSet As AceADODB.Recordset, ByVal SqlStr As String,Optional ByVal RSLocation As CursorLocationEnum = adUseClient,Optional ByVal RecOpenType As AceADODB.CursorTypeEnum = adOpenStatic,Optional ByVal RecLockType As AceADODB.LockTypeEnum = adLockReadOnly,Optional ByVal dbCon As AceADODB.Connection,Optional ByVal ErrorObj As ClsLastTranErrorDtl = Null) As Boolean




  
  Try
  RSOpen = False
  
  If RecSet Is Nothing Then
    'RecSet not instantiated
    Set RecSet = New AceADODB.Recordset
  Else
    'Already RecSet Instantiated
  End If

  If RecSet.State = adStateOpen Then
    RecSet.Filter = adFilterNone
    RecSet.Close
  End If
  
  RecSet.CursorLocation = RSLocation
  
  If dbCon Is Nothing Then Set dbCon = db
  
  If dbCon.State = adStateClosed Then
 dbCon.Open
End If
    
  RecSet.Open SqlStr, dbCon, RecOpenType, RecLockType
  
  If RecSet.State = adStateOpen Then
  If RSLocation = adUseClient Then RecSet.ActiveConnection = Nothing
    
    RSOpen = True
  Else
    Call DBRollBackWithMessage(mCustom, "Error: Unable to Open Recordset.", eMsgBoxIcon.mCritical, dbCon)
  End If
  
  If gTransactionCount <= 0 Then dbCon.Close
Catch ex As Exception
  Select Case Err.Number
    Case -2147467259    'Communication link failure
        Try
        dbCon.Open  'retry to open
        
    Case 3246       'db Connection is in Middle of Transaction.
         'So no need to close
    Case Else
       If Not ErrorObj Is Nothing Then
        ErrorObj.ErrorNo = Err.Number
        ErrorObj.ErrorDesc = Err.Description
        ErrorObj.ErrorTime = Now()
      End If
      Call ErrMsgShow("Error Occured while Opening RecordSet.")

  End Select
Finally

End Try
End Function

Public Function RSOpenDB(ByRef dbCon As AceADODB.Connection, ByRef RecSet As AceADODB.Recordset, ByVal SqlStr As String,Optional ByVal ByVal RSLocation As CursorLocationEnum = adUseClient,Optional ByVal ByVal RecOpenType As AceADODB.CursorTypeEnum = adOpenStatic,Optional ByVal ByVal RecLockType As AceADODB.LockTypeEnum = adLockReadOnly) As Boolean





  Try
  RSOpenDB = False
  
  If RecSet Is Nothing Then
    'RecSet not instantiated
    Set RecSet = New AceADODB.Recordset
  Else
    'Already RecSet Instantiated
  End If

  If RecSet.State = adStateOpen Then
    RecSet.Filter = adFilterNone
    RecSet.Close
  End If
  
  RecSet.CursorLocation = RSLocation
  
  If dbCon.State = adStateClosed Then dbCon.Open
  
  RecSet.Open SqlStr, dbCon, RecOpenType, RecLockType
  
  If RecSet.State = adStateOpen Then
    If RSLocation = adUseClient Then RecSet.ActiveConnection = Nothing
    
    RSOpenDB = True
  Else
    Call DBRollBackWithMessage(mCustom, "Error: Unable to Open Recordset.", eMsgBoxIcon.mCritical)
  End If
  
  If dbCon.State = adStateOpen Then dbCon.Close
Catch ex As Exception
  Call ErrMsgShow("Error Occured while Opening RecordSet.")

Finally

End Try
End Function

