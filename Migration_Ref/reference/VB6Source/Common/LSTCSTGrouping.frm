VERSION 5.00
Begin VB.Form frmLSTCSTGrouping 
   Caption         =   "LST, CST & ED Grouping"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame fraMain 
      Height          =   615
      Left            =   2820
      TabIndex        =   5
      Top             =   990
      Width           =   5880
      Begin VB.TextBox txttype 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1095
         TabIndex        =   7
         Top             =   180
         Width           =   3675
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   330
         Left            =   5160
         TabIndex        =   0
         Top             =   195
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   210
         Left            =   165
         TabIndex        =   6
         Top             =   240
         Width           =   810
      End
   End
Begin AceSpread spGroup
      Height          =   1500
      Left            =   4005
      TabIndex        =   1
      Top             =   1665
      Width           =   3525
      _Version        =   458752
      _ExtentX        =   6218
      _ExtentY        =   2646
      _StockProps     =   64
      Enabled         =   0   'False
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
      MaxCols         =   3
      ScrollBars      =   2
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6405
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   3240
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5115
      TabIndex        =   3
      ToolTipText     =   "Click to Clear the Data"
      Top             =   3240
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3825
      TabIndex        =   2
      ToolTipText     =   "Click to Save the Data"
      Top             =   3240
      Width           =   1300
   End
End
Attribute VB_Name = "frmLSTCSTGrouping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub cmbDesc_Change()

End Sub

Private Sub cmdClear_Click()
  
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  fraMain.Enabled = True
  spGroup.Enabled = False
  cmdSave.Enabled = False
 ' cmbDesc.SetFocus

 ' cmbDesc.ListIndex = 0
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
  Dim Sql As String
  Dim snap As AceADODB.Recordset
  Dim i As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  Sql = "select distinct a.ExprGroupID, (select b.ExprGroupName  from "
  Sql = Sql + " COM_FormulaExpr_Group b where b.ModuleCode = " & gModuleCode & " "
  Sql = Sql + " and b.ExprGroupID = a.ExprGroupID and b.Branchid = a.Branchid "
  Sql = Sql + " and b.Compid = a.Compid) ExprGroupName from COM_LSTCSTGrouping a "
  Sql = Sql + " where a.TaxDesc = '" + Trim$(txttype.Text) + "'"
  Sql = Sql + " and a.Branchid = '" & gBranchID & "' and a.Compid = '" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  If RSOpen(snap, Sql) = False Then
 GoTo ExitHere
End If
  If Not snap.EOF Then
    snap.MoveLast
snap.MoveFirst
    spGroup.MaxRows = snap.RecordCount + 1
    For i = 1 To snap.RecordCount
      spGroup.SetText 1, i, snap("ExprGroupID")
      spGroup.SetText 2, i, snap("ExprGroupName")
      spGroup.SetText 3, i, 1
      snap.MoveNext
    Next i
  End If
  
  fraMain.Enabled = False
  spGroup.Enabled = True
  cmdSave.Enabled = True
  spGroup.SetFocus
  Spread_Set_Focus spGroup, 1, 2
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim i As Long
  Dim tmp As Variant
  Dim CurrDate As Date
  
  Try
  Me.MousePointer = vbHourglass

  If Validate = False Then
 GoTo ExitHere
End If
  
  Set cmd = New AceADODB.Command
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "COM_Ins_LSTCSTGrouping"

  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  For i = 1 To spGroup.DataRowCnt
    commClear cmd
    cmd.Parameters.Append cmd.CreateParameter("ModuleCode", adInteger, adParamInput, 30, Trim$(gModuleCode))
    cmd.Parameters.Append cmd.CreateParameter("TaxCode", adInteger, adParamInput, 30, Trim$(txttype.Tag))
    cmd.Parameters.Append cmd.CreateParameter("TaxDesc", adVarChar, adParamInput, 30, Trim$(txttype.Text))
    spGroup.GetText 1, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("ExprGroupID", adInteger, adParamInput, 30, Val(tmp))
    spGroup.GetText 2, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("ExprGroup", adVarChar, adParamInput, 30, tmp)
    cmd.Parameters.Append cmd.CreateParameter("IncludeSubGroup", adInteger, adParamInput, , "0")
    cmd.Parameters.Append cmd.CreateParameter("RevNo", adInteger, adParamInput, 30, "0")
    cmd.Parameters.Append cmd.CreateParameter("RevDate", adDBTimeStamp, adParamInput, 30, Format(ServerDate, "dd-mm-yyyy"))
    cmd.Parameters.Append cmd.CreateParameter("EffStatus", adVarChar, adParamInput, 30, "AN")
    cmd.Parameters.Append cmd.CreateParameter("EffDate", adDBTimeStamp, adParamInput, 30, Format(ServerDate, "dd-mm-yyyy"))
    cmd.Parameters.Append cmd.CreateParameter("isTranLock", adVarChar, adParamInput, 30, "0")
    Call AddCommonCmdParameters(cmd, , False)
    cmd.Parameters.Append cmd.CreateParameter("TranStamp", adDBTimeStamp, adParamInput, 30, Format(ServerDate, "dd-mm-yyyy"))
    cmd.Execute
  Next i
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  sShowMessage "Record Saved"
  Call cmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set cmd = Nothing
  Me.MousePointer = Default
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
    
 ' cmbDesc.ListIndex = 0
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

Private Sub spGroup_DblClick(sender as object, e as DblClickEventArgs)
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass
  
  If Col = 2 Then
    Sql = "Select ExprGroupName,ExprGroupID From COM_FormulaExpr_Group where Route <> ''"
    Sql = Sql + " and isActive = 1 and ModuleCode = " & gModuleCode & " and Branchid ='" & gBranchID & "'"
    Sql = Sql + " and Compid = '" & gCompID & "'"
    If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Expression Group", "Expression Group Name") = True Then
      spGroup.SetText 2, Row, getvalue(1)
      spGroup.SetText 1, Row, getvalue(2)
      Spread_Clear_Duplicate spGroup, Row, 2
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub spGroup_KeyDown(sender as object, e as KeyDownEventArgs)
' On Error Resume Next
  If spGroup.ActiveCol = 2 And KeyCode = 13 Then
    spGroup.MaxRows = spGroup.DataRowCnt + 1
  ElseIf KeyCode = vbKeyF2 Then
    spGroup_DblClick 2, spGroup.ActiveRow
  ElseIf KeyCode = 46 And spGroup.ActiveCol = 2 And Val(SpVal(spGroup, spGroup.ActiveRow, 3)) = 0 Then
    Spread_Clear_Row spGroup, 1, spGroup.MaxCols, spGroup.ActiveRow, spGroup.ActiveRow
    spGroup.DeleteRows spGroup.ActiveRow, 1
    spGroup.MaxRows = spGroup.DataRowCnt + 1
  End If
End Sub

Private Sub txttype_DblClick()
Dim Sql As String
  Dim rs As AceADODB.Recordset

  Try
  Me.MousePointer = vbHourglass

  Sql = "select  TaxCode,TaxDesc from COM_LSTCSTGrouping_Master a "
  Sql = Sql & " where Branchid = '" & gBranchID & "' and CompId = '" & gCompID & "'"
 ' sql = sql & " order by DocNo"
  If FetchValue_New(False, db.ConnectionString, Sql, "01", "List of Tax", "TaxType", "TaxType", "PlanDate", "Production Date", "Shift") = True Then
    txttype.Tag = getvalue(1)
    txttype.Text = getvalue(2)
    
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Function Validate() As Boolean
  Dim i As Long
  Dim tmp As Variant
  
  Try
  
  Validate = False
  
  If spGroup.DataRowCnt = 0 Then
    sShowMessage "Please select the Expression"
    spGroup.SetFocus
    Spread_Set_Focus spGroup, 1, 2
    Validate = False
    GoTo ExitHere
  End If
  
  For i = 1 To spGroup.DataRowCnt
    spGroup.GetText 2, i, tmp
    If Trim$(tmp) = "" Then
      sShowMessage "Please select the Expression"
      spGroup.SetFocus
      Spread_Set_Focus spGroup, i, 2
      Validate = False
      GoTo ExitHere
    End If
  Next i
  
  Validate = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Design_spGroup(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spGroup


spGroup.DblClick +=  new EventHandler(spGroup_DblClick)
spGroup.KeyDown +=  new EventHandler(spGroup_KeyDown)

End Sub

Private Sub InitializeSpreads()

	Call Design_spGroup("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Common\LSTCSTGrouping.frm", spGroup, "")

End Sub
