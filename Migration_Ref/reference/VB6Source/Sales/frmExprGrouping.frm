VERSION 5.00
Begin VB.Form frmExprGrouping 
   Caption         =   "Group Classification"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Group"
      Height          =   400
      Left            =   5700
      TabIndex        =   4
      Top             =   5220
      Width           =   1300
   End
Begin AceSpread sprExpr
      Height          =   4620
      Left            =   495
      TabIndex        =   3
      Top             =   285
      Width           =   10470
      _Version        =   458752
      _ExtentX        =   18468
      _ExtentY        =   8149
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      UserResize      =   0
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6375
      TabIndex        =   0
      Top             =   5805
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5085
      TabIndex        =   1
      Top             =   5805
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   400
      Left            =   3795
      TabIndex        =   2
      Top             =   5805
      Width           =   1300
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Child"
      Height          =   400
      Left            =   4410
      TabIndex        =   5
      Top             =   5220
      Width           =   1300
   End
   Begin VB.Menu mnuPopup 
      Caption         =   " "
      Visible         =   0   'False
      Begin VB.Menu mnuNewChildElem 
         Caption         =   "New Child"
      End
      Begin VB.Menu mnuDeleteElem 
         Caption         =   "Delete"
      End
   End
End
Attribute VB_Name = "frmExprGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private LoadProcess As Boolean

Private Enum eExpr
  mExprID = 1
  mExprShort = 2
  mExprName = 3
  mRoute = 4
  mWidth = 5
  mIndStatus = 6
  mFinalStatus = 7
  mNew = 8
End Enum

Private Sub CmdClear_Click()
' On Error Resume Next
  Call LoadExprGroupInSpread
End Sub

Private Sub cmdexit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Function ValidateData() As Boolean
  Dim retVal As Variant
  Dim rCnt As Long
  
  Try
  ValidateData = False
  
    For rCnt = 1 To sprExpr.MaxRows
'      .GetText eExpr.mExprShort, rCnt, retVal
'      If Trim$(retVal) = "" Then
'        sShowMessage "Please Type the Expression Group Short Name in the Row " & rCnt
'        .SetFocus
'        .SetActiveCell eExpr.mExprShort, rCnt
'        GoTo ExitHere
'      End If
      
      sprExpr.GetText eExpr.mExprName, rCnt, retVal
      If Trim$(retVal) = "" Then
        sShowMessage "Please Type the Expression Group Long Name in the Row " & rCnt
        sprExpr.SetFocus
        sprExpr.SetActiveCell eExpr.mExprName, rCnt
        GoTo ExitHere
      End If
      
    Next rCnt
    
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdSave_Click()
  Dim Sql As String
  Dim mMaxGroupID As Integer
  Dim rCnt As Long
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  mMaxGroupID = GetMax("COM_FormulaExpr_Group", "ExprGroupID", "ModuleCode", gModuleCode)
  
    For rCnt = 1 To sprExpr.MaxRows
      sprExpr.GetText eExpr.mNew, rCnt, retVal
      If UCase(Trim$(retVal)) = "NEW" Then
        mMaxGroupID = mMaxGroupID + 1
        Sql = "Execute COM_Ins_FormulaExpr_Group"
        Sql = Sql & " @ModuleCode=" & gModuleCode & ","
        Sql = Sql & " @ExprGroupID=" & mMaxGroupID & ","
        sprExpr.GetText eExpr.mExprShort, rCnt, retVal
          Sql = Sql & " @ExprShortName='" & Trim$(retVal) & "',"
        sprExpr.GetText eExpr.mExprName, rCnt, retVal
          Sql = Sql & " @ExprGroupName='" & Trim$(retVal) & "',"
        sprExpr.GetText eExpr.mRoute, rCnt, retVal
          Sql = Sql & " @Route='" & Trim$(retVal) & "',"
        sprExpr.GetText eExpr.mWidth, rCnt, retVal
          Sql = Sql & " @Width='" & Trim$(retVal) & "',"
        sprExpr.GetText eExpr.mIndStatus, rCnt, retVal
          Sql = Sql & " @isActive='" & Val(retVal) & "',"
        Sql = Sql & " @BranchID='" & gBranchID & "',"
        Sql = Sql & " @CompID='" & gCompID & "'"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      Else
        Sql = "Update COM_FormulaExpr_Group"
        sprExpr.GetText eExpr.mExprShort, rCnt, retVal
          Sql = Sql & " Set ExprShortName='" & Trim$(retVal) & "',"
        sprExpr.GetText eExpr.mExprName, rCnt, retVal
          Sql = Sql & " ExprGroupName='" & Trim$(retVal) & "',"
        sprExpr.GetText eExpr.mIndStatus, rCnt, retVal
          Sql = Sql & " isActive='" & Val(retVal) & "'"
        sprExpr.GetText eExpr.mExprID, rCnt, retVal
          Sql = Sql & " Where ExprGroupID=" & Val(retVal) & ""
        sprExpr.GetText eExpr.mRoute, rCnt, retVal
          Sql = Sql & " and Route='" & Trim$(retVal) & "'"
        sprExpr.GetText eExpr.mWidth, rCnt, retVal
          Sql = Sql & " and Width='" & Trim$(retVal) & "'"
        Sql = Sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
        If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
      End If
    Next rCnt
  
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  sShowMessage "Record Saved Successfully"
  
  Call CmdClear_Click
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

  Call LoadExprGroupInSpread
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

Private Sub sprExpr_ButtonClicked(sender as object, e as ButtonClickedEventArgs)
  Dim retVal As Variant
  Dim pStatus As Integer
  Dim pRoute As String
  Dim rCnt As Long
  
  Try
  Me.MousePointer = vbHourglass

  If LoadProcess = False Then
      If Col = eExpr.mIndStatus Then
        sprExpr.GetText eExpr.mIndStatus, Row, retVal
        pStatus = Val(retVal)
        
        sprExpr.GetText eExpr.mRoute, Row, retVal
        pRoute = Trim$(retVal)
        
        pStatus = pStatus * IIf(GetParentStatus(pRoute), 1, 0)
        sprExpr.SetText eExpr.mFinalStatus, Row, pStatus
      Else
        sprExpr.GetText eExpr.mFinalStatus, Row, retVal
        pStatus = Val(retVal)
      
        sprExpr.GetText eExpr.mRoute, Row, retVal
        pRoute = Trim$(retVal)
      
        sprExpr.GetText eExpr.mWidth, Row, retVal
        pRoute = pRoute & Trim$(retVal)
        
        For rCnt = Row + 1 To sprExpr.MaxRows
          sprExpr.GetText eExpr.mRoute, rCnt, retVal
          If Trim$(retVal) <> pRoute Then
            Exit For
          Else
            sprExpr.GetText eExpr.mIndStatus, rCnt, retVal
            
            'To get Final Status, Multiply individual status with parent status i.e. only both are true it will true
            sprExpr.SetText eExpr.mFinalStatus, rCnt, Val(retVal) * pStatus
          End If
        Next rCnt
      End If
  Else
    LoadProcess = False
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprExpr_EditMode(sender as object, e as EditModeEventArgs)
  Dim retVal As Variant
  Dim retVal2 As Variant
  
  Try
  Me.MousePointer = vbHourglass

  'ExprName Converted to Upper Case
  If Mode = 0 And ChangeMade = True Then
      If Col = eExpr.mExprName Then
        sprExpr.GetText eExpr.mRoute, Row, retVal2
        sprExpr.GetText Col, Row, retVal
        sprExpr.SetText Col, Row, Space(Len(Trim$(retVal2)) * 4) & UCase(Trim$(retVal))
      ElseIf Col = eExpr.mExprShort Then
        sprExpr.GetText Col, Row, retVal
        sprExpr.SetText Col, Row, UCase(Trim$(retVal))
      End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdDelete_Click()
  Dim retVal As Variant
  
  Try
  Me.MousePointer = vbHourglass

    sprExpr.GetText eExpr.mNew, sprExpr.ActiveRow, retVal
    If Trim$(retVal) = "Old" Then
      sShowMessage "Saved Expression Cannot be deleted."
    Else
      sprExpr.GetText eExpr.mExprName, sprExpr.ActiveRow, retVal
      If MsgBox("Do you really want to delete '" & Trim$(retVal) & "'?", vbYesNo, "Delete?") = vbYes Then
        sprExpr.DeleteRows sprExpr.ActiveRow, 1
        sprExpr.MaxRows = sprExpr.DataRowCnt
      End If
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdAdd_Click()
  Dim retVal As Variant
  Dim Route As String
  Dim Width As String
  Dim Path As String
  Dim rPos As Long
  Dim rCnt As Long
  Dim cRow As Long
  Dim RouteAvail As Integer

  Try
  Me.MousePointer = vbHourglass

    sprExpr.GetText eExpr.mWidth, sprExpr.ActiveRow, retVal
    Width = Trim$(retVal)
    
    sprExpr.GetText eExpr.mRoute, sprExpr.ActiveRow, retVal
    Route = Trim$(retVal)
    
    Path = Route & Width
    
    If NoOfOccurence(sprExpr, eExpr.mRoute, Path) < 99 Then
      rPos = sprExpr.SearchCol(eExpr.mRoute, 1, sprExpr.MaxRows, Path, SearchFlagsNone)
      If rPos <= 0 Then
        rPos = sprExpr.ActiveRow
        cRow = rPos
      Else
        RouteAvail = 1
        
        If rPos = sprExpr.MaxRows Then
          cRow = sprExpr.MaxRows
        Else
          For rCnt = rPos + 1 To sprExpr.MaxRows
            sprExpr.GetText eExpr.mRoute, rCnt, retVal
            If Trim$(retVal) = Path Then
              RouteAvail = RouteAvail + 1
              
              If RouteAvail = 99 Then
                sprExpr.GetText eExpr.mExprName, sprExpr.ActiveRow, retVal
                MsgBox "Max No. of Child Element reached for '" & Trim$(retVal) & "'"
                GoTo ExitHere
              End If
            Else
              Exit For
            End If
          Next rCnt
          
          cRow = RouteAvail + rPos - 1
        End If
      End If
      
      If cRow = sprExpr.MaxRows Then
        sprExpr.MaxRows = sprExpr.MaxRows + 1
        cRow = sprExpr.MaxRows
      Else
        sprExpr.InsertRows cRow + 1, 1
        cRow = cRow + 1
      End If
      
      sprExpr.SetText eExpr.mRoute, cRow, Path
      
      sprExpr.GetText eExpr.mWidth, cRow - 1, retVal
      sprExpr.SetText eExpr.mWidth, cRow, Format(Val(retVal) + 1, "00")
    
      LoadProcess = True
      sprExpr.SetText eExpr.mIndStatus, cRow, "1"
      
      sprExpr.GetText eExpr.mFinalStatus, rPos, retVal
      If retVal = False Then
        sprExpr.SetText eExpr.mFinalStatus, cRow, "0"
      Else
        sprExpr.GetText eExpr.mIndStatus, rPos, retVal
        If retVal = False Then
          sprExpr.SetText eExpr.mFinalStatus, cRow, "0"
        Else
          sprExpr.SetText eExpr.mFinalStatus, cRow, "1"
        End If
      End If
      
      sprExpr.SetText eExpr.mNew, cRow, "New"
      
      sprExpr.SetFocus
      sprExpr.SetActiveCell eExpr.mExprName, cRow
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub LoadExprGroupInSpread()
  Dim Sql As String
  Dim rsNode As AceADODB.Recordset
  Dim Cnt As Integer
  
  Try
  
  Sql = "Select ExprGroupID, ExprShortName, ExprGroupName, Route, Width, isActive, ByParentStatus"
  Sql = Sql & " from ExpressionGrouping"
  Sql = Sql & " Where ModuleCode=" & gModuleCode ' & "  and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Sql = Sql & " Order by Route+Width"
  Set rsNode = New AceADODB.Recordset
  If RSOpen(rsNode, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsNode.EOF Then
    rsNode.MoveLast
rsNode.MoveFirst
'    mSavedNodeCount = rsNode.RecordCount + 1      'Add 1 for the Root Node
    
      sprExpr.MaxRows = rsNode.RecordCount
      For Cnt = 1 To rsNode.RecordCount
        sprExpr.SetText eExpr.mExprID, Cnt, rsNode("ExprGroupID")
        sprExpr.SetText eExpr.mExprShort, Cnt, rsNode("ExprShortName")
        sprExpr.SetText eExpr.mExprName, Cnt, Space(Len(rsNode("Route")) * 4) & Trim$(rsNode("ExprGroupName"))
        sprExpr.SetText eExpr.mRoute, Cnt, rsNode("Route")
        sprExpr.SetText eExpr.mWidth, Cnt, rsNode("Width")
        sprExpr.SetText eExpr.mIndStatus, Cnt, IIf(rsNode("isActive"), "1", "0")
        sprExpr.SetText eExpr.mFinalStatus, Cnt, IIf(rsNode("ByParentStatus"), 1, 0) * IIf(rsNode("isActive"), 1, 0)
        sprExpr.SetText eExpr.mNew, Cnt, "Old"
        
        rsNode.MoveNext
      Next Cnt
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsNode = Nothing
End Try
End Sub

Private Function GetParentStatus(ByVal ChildRoute As String) As Boolean
  Dim retVal As Variant
  Dim rPos As Integer
  Dim pStatus As Boolean

  Try
  GetParentStatus = False
  
  If Len(Trim$(ChildRoute)) = 0 Then
    pStatus = True
  Else
      rPos = sprExpr.SearchCol(eExpr.mRoute, 0, sprExpr.MaxRows, Left$(Trim$(ChildRoute), Len(Trim$(ChildRoute)) - 2), SearchFlagsNone)
      If rPos <= 0 Then
        GoTo ExitHere
      Else
        sprExpr.GetText eExpr.mFinalStatus, rPos, retVal
        pStatus = IIf(Val(retVal) = 1, True, False)
      End If
  End If
  
  GetParentStatus = pStatus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_sprExpr(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprExpr


sprExpr.ButtonClicked +=  new EventHandler(sprExpr_ButtonClicked)
sprExpr.EditMode +=  new EventHandler(sprExpr_EditMode)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprExpr("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmExprGrouping.frm", sprExpr, "")

End Sub
