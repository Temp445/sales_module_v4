VERSION 5.00
Begin VB.Form frmFormulaMaster 
   Caption         =   "Formula Master"
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
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   4245
      Width           =   1300
   End
Begin AceSpread spFormula
      Height          =   2370
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Enter the Formula Calculation Details"
      Top             =   1755
      Width           =   11130
      _Version        =   458752
      _ExtentX        =   19632
      _ExtentY        =   4180
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
      MaxCols         =   13
      MaxRows         =   1
   End
   Begin VB.Frame fraName 
      Height          =   630
      Left            =   3165
      TabIndex        =   0
      Top             =   1020
      Width           =   4710
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Left            =   4005
         TabIndex        =   3
         ToolTipText     =   "Click to OK"
         Top             =   210
         Width           =   585
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1185
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Enter the Formula Name"
         Top             =   210
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Formula Name"
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   262
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   5190
      TabIndex        =   6
      ToolTipText     =   "Click to Clear the Data"
      Top             =   4245
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3900
      TabIndex        =   7
      ToolTipText     =   "Click to Save the Data"
      Top             =   4245
      Width           =   1300
   End
End
Attribute VB_Name = "frmFormulaMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Function ValidateData() As Boolean
  Dim tmp As Variant
  Dim i As Long
  
  Try
  ValidateData = False
  
  For i = 1 To spFormula.DataRowCnt
    spFormula.GetText 2, i, tmp
    If Trim(tmp) = "" Then
      sShowMessage "Please select the Formula Parameter"
      spFormula.SetFocus
      Spread_Set_Focus spFormula, i, 2
      GoTo ExitHere
    End If
    
    spFormula.GetText 5, i, tmp
    If Trim(tmp) = "" Then
      sShowMessage "Please select the Formula Parameter"
      spFormula.SetFocus
      Spread_Set_Focus spFormula, i, 5
      GoTo ExitHere
    End If
    
    spFormula.GetText 6, i, tmp
    If Trim(tmp) = "" Then
      sShowMessage "Please select the Operator"
      spFormula.SetFocus
      Spread_Set_Focus spFormula, i, 6
      GoTo ExitHere
    End If
    
    spFormula.GetText 8, i, tmp
    If Trim(tmp) = "" Then
      sShowMessage "Please select the Formula Parameter"
      spFormula.SetFocus
      Spread_Set_Focus spFormula, i, 8
      GoTo ExitHere
    End If
    
'    spFormula.GetText 10, i, tmp
'    If Trim(tmp) = "" Then
'      sShowMessage "Please select the Posting Account"
'      spFormula.SetFocus
'      Spread_Set_Focus spFormula, i, 10
'      GoTo ExitHere
'    End If
    
    spFormula.GetText 11, i, tmp
    If Trim(tmp) = "" Then
      sShowMessage "Please select the Rounding Off"
      spFormula.SetFocus
      Spread_Set_Focus spFormula, i, 11
      GoTo ExitHere
    End If
    
    If Trim(tmp) <> "None" Then
      spFormula.GetText 12, i, tmp
      If Trim(tmp) = "" Then
        sShowMessage "Please enter the Rounding of Value"
        spFormula.SetFocus
        Spread_Set_Focus spFormula, i, 12
        GoTo ExitHere
      End If
    End If
  Next i
    
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdClear_Click()
  Try
  Me.MousePointer = vbHourglass
  Form_Clear Me
  fraName.Enabled = True
  spFormula.Enabled = False
  cmdSave.Enabled = False
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

Private Sub cmdOk_Click()
  Try
  Me.MousePointer = vbHourglass
  
  If Trim(txtName) = "" Then
    sShowMessage "Please enter the Formula Name"
    txtName.SetFocus
    GoTo ExitHere
  End If
  
  fraName.Enabled = False
  spFormula.Enabled = True
  cmdSave.Enabled = True
  spFormula.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub cmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim tmp As Variant
  Dim i As Long
  
  Try
  Me.MousePointer = vbHourglass

  Set cmd = New AceADODB.Command
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  txtName.Tag = Val(Find_Value("COM_FormulaMast_Hdr", "isnull(max(FormulaCode),0)", "compid", gCompID)) + 1
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  cmd.CommandText = "COM_INS_FORMULA_MASTER"
  
  For i = 1 To spFormula.DataRowCnt
    commClear cmd
    cmd.Parameters.Append cmd.CreateParameter("FormulaCode", adInteger, adParamInput, , Val(txtName.Tag))
    cmd.Parameters.Append cmd.CreateParameter("FrmLName", adVarChar, adParamInput, 30, Trim$(txtName.Text))
    spFormula.GetText 1, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("Code", adInteger, adParamInput, , Val(tmp))
    spFormula.GetText 4, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("Code1", adInteger, adParamInput, , Val(tmp))
    spFormula.GetText 6, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("Operator", adChar, adParamInput, 1, Trim(tmp))
    spFormula.GetText 7, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("Code2", adInteger, adParamInput, , Val(tmp))
    spFormula.GetText 11, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("RoundingOff", adChar, adParamInput, 2, Left(Trim(tmp), 2))
    spFormula.GetText 12, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("RoundingOfValue", adInteger, adParamInput, , Val(tmp))
    cmd.Parameters.Append cmd.CreateParameter("ModuleName", adVarChar, adParamInput, 3, "SAL")
    spFormula.GetText 9, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("Posting", adInteger, adParamInput, , Val(tmp))
    cmd.Parameters.Append cmd.CreateParameter("Sequence", adInteger, adParamInput, , i)
    spFormula.GetText 13, i, tmp
    cmd.Parameters.Append cmd.CreateParameter("Printing", adBoolean, adParamInput, , IIf(Val(tmp) = 1, 1, 0))
    cmd.Parameters.Append cmd.CreateParameter("userid", adVarChar, adParamInput, 10, Trim$(gUserID))
    cmd.Parameters.Append cmd.CreateParameter("Branchid", adVarChar, adParamInput, 3, Trim$(gBranchID))
    cmd.Parameters.Append cmd.CreateParameter("Compid", adVarChar, adParamInput, 3, Trim$(gCompID))
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

  Me.MousePointer = 0
End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  Try
  
  txtName.Tag = Val(Find_Value("COM_FormulaMast_Hdr", "isnull(max(FormulaCode),0)", "compid", gCompID)) + 1
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub spFormula_DblClick(sender as object, e as DblClickEventArgs)
  Try
  Me.MousePointer = vbHourglass
  
  Dim Sql As String
  
  If Col = 2 Then
    Sql = "select Description,Code from COM_FORMULA_PARAMETER where Tag='C'"
    Sql = Sql + " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
    If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Formula Parameter", "Parameter Name") = True Then
      spFormula.SetText 2, Row, getvalue(1)
      spFormula.SetText 1, Row, getvalue(2)
    End If
  ElseIf Col = 5 Then
    Sql = "select Description,Code from COM_FORMULA_PARAMETER where "
    Sql = Sql + " Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
    Sql = Sql + " and Code <> " & Val(SpVal(spFormula, Row, 1)) & ""
    If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Formula Parameter", "Parameter Name") = True Then
      spFormula.SetText 5, Row, getvalue(1)
      spFormula.SetText 4, Row, getvalue(2)
    End If
  ElseIf Col = 8 Then
    Sql = "select Description,Code from COM_FORMULA_PARAMETER where "
    Sql = Sql + " Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
    Sql = Sql + " and Code not in (" & Val(SpVal(spFormula, Row, 1)) & "," & Val(SpVal(spFormula, Row, 4)) & ")"
    If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Formula Parameter", "Parameter Name") = True Then
      spFormula.SetText 8, Row, getvalue(1)
      spFormula.SetText 7, Row, getvalue(2)
    End If
  ElseIf Col = 10 Then
    Sql = "select LedgerName,LedgerCode from FAS_LedgerMaster where left(LrRoute,4) not in ('0206','0303')"
    Sql = Sql + " and CompCode = '" + Left(gFasCompCode, 4) + "'"
    If FetchValue_New(False, db.ConnectionString, Sql, "10", "List of Accounts", "Account Name") = True Then
      spFormula.SetText 10, Row, getvalue(1)
      spFormula.SetText 9, Row, getvalue(2)
    End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub spFormula_KeyDown(sender as object, e as KeyDownEventArgs)
  Try
  Me.MousePointer = vbHourglass
  
  If KeyCode = vbKeyF2 Then
    spFormula_DblClick spFormula.ActiveCol, spFormula.ActiveRow
  ElseIf KeyCode = 13 And spFormula.ActiveCol = 13 Then
    spFormula.MaxRows = spFormula.DataRowCnt + 1
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub Design_spFormula(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'spFormula


spFormula.DblClick +=  new EventHandler(spFormula_DblClick)
spFormula.KeyDown +=  new EventHandler(spFormula_KeyDown)

End Sub

Private Sub InitializeSpreads()

	Call Design_spFormula("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\FormulaMaster.frm", spFormula, "")

End Sub
