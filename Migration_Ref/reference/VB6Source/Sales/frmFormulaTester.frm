VERSION 5.00
Begin VB.Form frmFormulaTester 
   Caption         =   "Formula Testing Screen"
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
   Begin VB.TextBox txtUserKey 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   9555
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtGrandTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8850
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   10
      Top             =   4020
      Width           =   1470
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      Height          =   390
      Left            =   6630
      TabIndex        =   6
      Top             =   4470
      Width           =   1170
   End
   Begin VB.Frame fraFormula 
      Caption         =   "Formula"
      Height          =   870
      Left            =   2955
      TabIndex        =   1
      Top             =   570
      Width           =   6030
      Begin VB.CommandButton cmdProceed 
         Caption         =   "&Proceed"
         Height          =   390
         Left            =   4725
         TabIndex        =   8
         Top             =   390
         Width           =   1170
      End
      Begin VB.TextBox txtFLName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   3
         Top             =   435
         Width           =   3510
      End
      Begin VB.TextBox txtFSName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   2
         Top             =   435
         Width           =   960
      End
      Begin VB.Label lblFLName 
         AutoSize        =   -1  'True
         Caption         =   "Long Name"
         Height          =   195
         Left            =   1125
         TabIndex        =   5
         Top             =   225
         Width           =   825
      End
      Begin VB.Label lblFSName 
         AutoSize        =   -1  'True
         Caption         =   "Short Name"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   225
         Width           =   840
      End
   End
Begin AceSpread sprExpr
      Height          =   2430
      Left            =   1275
      TabIndex        =   0
      Top             =   1530
      Width           =   9255
      _Version        =   458752
      _ExtentX        =   16325
      _ExtentY        =   4286
      _StockProps     =   64
      ButtonDrawMode  =   4
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
      SelectBlockOptions=   0
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   390
      Left            =   5475
      TabIndex        =   7
      Top             =   4470
      Width           =   1170
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   390
      Left            =   4320
      TabIndex        =   9
      Top             =   4470
      Width           =   1170
   End
   Begin VB.Label lblUserKey 
      AutoSize        =   -1  'True
      Caption         =   "User Key"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9555
      TabIndex        =   13
      Top             =   990
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblGrandTotal 
      AutoSize        =   -1  'True
      Caption         =   "Grand Total"
      Height          =   195
      Left            =   7950
      TabIndex        =   11
      Top             =   4065
      Width           =   840
   End
End
Attribute VB_Name = "frmFormulaTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'For Formula Spread Column Nos

Private Enum eFrm
  mExName = 1
  mExBack = 2
  mExpr = 3
  mRndOff = 4
  mRndOffVal = 5
  mResult = 6
  mGrand = 7
  mAcName = 8
End Enum

Private Function GetExprResults12(ByVal UserKey As Long) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  
  Sql = "Select ExprSeq, Result from Com_tmpFormulaCalc_Hdr"
  Sql = Sql & " Where DTKey='" & UserKey & "' and FrmID=" & Val(txtFSName.Tag) & " and UserID='" & gUserID & "' and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  Do While Not rs.EOF
    sprExpr.SetText eFrm.mResult, rs("ExprSeq"), rs("Result")
    rs.MoveNext
  Loop
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub CmdClear_Click()
  Dim Sql As String

  Try
  Me.MousePointer = vbHourglass

  Call ClearTmpPromptTable
  txtFSName.Text = ""
  txtFSName.Tag = ""
  txtFLName.Text = ""
  txtGrandTotal.Text = ""
  Call sprExpr.ClearRange(1, 1, -1, -1, True)
'  Call sprParamPrompt.ClearRange(1, 1, -1, -1, True)
  sprExpr.MaxRows = 1
  fraFormula.Enabled = True
  txtFSName.SetFocus
  
  Sql = "Delete from COM_tmpFormulaCalc_Dtl Where DTKey='" & Trim$(txtUserKey.Text) & "' and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "' and Finyear='" & gFinyear & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  
  Sql = "Delete from COM_tmpFormulaCalc_Hdr Where DTKey='" & Trim$(txtUserKey.Text) & "' and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "' and Finyear='" & gFinyear & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdClose_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Function LoadExpr(ByVal FormulaID As Integer) As Boolean
  Dim Sql As String
  Dim rsExpr As AceADODB.Recordset
  Dim rCnt As Long
  
  Try
  
  LoadExpr = False
  
  Sql = " Select a.ExprSeq, Case When a.ExprRoute < 0 then 'SUB TOTAL' When a.ExprRoute='0001' Then 'TOTAL GST' else b.ExprGroupName End as ExprName,"  'isNull(b.ExprGroupName, 'SUBTOTAL') as ExprName,
  Sql = Sql & " a.RoundOff, a.RoundOffValue,"
  Sql = Sql & " Case a.GrandTotalFlag When 1 then '+' When -1 then '-' Else '' End as GTFlag,"
  Sql = Sql & " a.PostingAc, c.LedgerName as AcName"
  Sql = Sql & " from Com_FormulaMast_Dtl a Left Outer Join COM_FormulaExpr_Group b On b.Route+b.Width=a.ExprRoute and b.ModuleCode=a.ModuleCode and b.BranchID=a.BranchID and b.CompID=a.CompID"
  Sql = Sql & " Left Outer Join FAS_LedgerMaster c On c.LedgerCode=a.PostingAc and c.CompCode=a.CompID + a.BranchID"
  Sql = Sql & " where a.FrmID=" & FormulaID & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
  Sql = Sql & " Order by ExprSeq"
  Set rsExpr = New AceADODB.Recordset
  If RSOpen(rsExpr, Sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsExpr.EOF Then
    'To retrieve the record count
    rsExpr.MoveLast
rsExpr.MoveFirst
    
      sprExpr.MaxRows = rsExpr.RecordCount
      Do While Not rsExpr.EOF
        'Sequence No for Row
        rCnt = rsExpr("ExprSeq")
        
        sprExpr.SetText eFrm.mExName, rCnt, rsExpr("ExprName")
        'To Convert and set the FrontEnd Expression from BackEnd Expression
        sprExpr.SetText eFrm.mExpr, rCnt, GetFrontEndExpr(Val(txtFSName.Tag), rsExpr("ExprSeq"))
        'To set the Rounding Off Text by the options
        sprExpr.SetText eFrm.mRndOff, rCnt, GetRoundingOffText(rsExpr("RoundOff"))
        sprExpr.SetText eFrm.mRndOffVal, rCnt, IIf(rsExpr("RoundOffValue") = 0, "", rsExpr("RoundOffValue"))
        sprExpr.SetText eFrm.mGrand, rCnt, rsExpr("GTFlag")
        sprExpr.SetText eFrm.mAcName, rCnt, rsExpr("AcName")
        
        rsExpr.MoveNext
      Loop
  End If
  
  LoadExpr = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsExpr = Nothing
End Try
End Function

Private Sub cmdProceed_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If LoadExpr(Val(txtFSName.Tag)) = False Then
    GoTo ExitHere
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdStart_Click()
  
  Try
  Me.MousePointer = vbHourglass
  
  If StartCalcFormula(True) = True Then
    GoTo ExitHere
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
' On Error Resume Next
  sprExpr.MaxRows = 1
End Sub

Private Function StartCalcFormula(Optional ByVal ShowPromptScreen As Boolean = False) As Boolean
  Dim ValueList(0) As String
  Dim ItemList(0) As String
  
  Try
  StartCalcFormula = False
  
  ValueList(0) = "0" & Chr(145) & " " & Chr(145) & "0"
  ItemList(0) = "0" & Chr(145) & "0"
  
  If CalculateFormula(0, Val(txtFSName.Tag), txtUserKey, ValueList, ItemList, txtGrandTotal, sprExpr, eFrm.mExName, eFrm.mResult, , "Formula Tester", True, ShowPromptScreen) = False Then

    GoTo ExitHere
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  
  Call ClearTmpPromptTable
  If FormExit(Me, UnloadMode) = False Then
    Cancel = 1
  End If
End Sub

Private Sub txtFSName_DblClick()
  Dim Sql As String
  
  Try
  Me.MousePointer = vbHourglass

  Sql = "Select FrmID, FrmSName, FrmLName from Com_FormulaMast_Hdr"
  Sql = Sql & " Where ModuleCode=" & gModuleCode & " and isActive='1' and BranchID='" & gBranchID & "' and CompID='" & gCompID & "' Order by FrmSName"
  If FetchValue_New(False, db.ConnectionString, Sql, "011", "List of Formulae", "ShortName", "LongName") = True Then
    txtFSName.Tag = getvalue(1)
    txtFSName.Text = getvalue(2)
    txtFLName.Text = getvalue(3)
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtFSName_KeyUp(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
    Call txtFSName_DblClick
  End If
End Sub

Private Sub Design_sprExpr(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprExpr

End Sub

Private Sub InitializeSpreads()

	Call Design_sprExpr("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmFormulaTester.frm", sprExpr, "")

End Sub
