VERSION 5.00
Begin VB.Form frmParamMaster 
   Caption         =   "Parameter Master"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00800000&
      Height          =   1065
      Left            =   705
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   5445
      Visible         =   0   'False
      Width           =   6300
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   4425
      TabIndex        =   19
      Top             =   4830
      Width           =   1170
   End
   Begin VB.Frame fraDefVal 
      Caption         =   "Defined Values"
      Enabled         =   0   'False
      Height          =   3780
      Left            =   7140
      TabIndex        =   14
      Top             =   915
      Visible         =   0   'False
      Width           =   2970
Begin AceSpread sprDefVal
         Height          =   1680
         Left            =   105
         TabIndex        =   15
         Top             =   225
         Width           =   2745
         _Version        =   458752
         _ExtentX        =   4842
         _ExtentY        =   2963
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   1
         ScrollBars      =   2
      End
Begin AceSpread sprDefPerc
         Height          =   1680
         Left            =   105
         TabIndex        =   21
         Top             =   1995
         Width           =   2745
         _Version        =   458752
         _ExtentX        =   4842
         _ExtentY        =   2963
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   1
         ScrollBars      =   2
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "C&lear"
      Height          =   360
      Left            =   3270
      TabIndex        =   8
      Top             =   4830
      Width           =   1170
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   360
      Left            =   2115
      TabIndex        =   7
      Top             =   4830
      Width           =   1170
   End
   Begin VB.Frame fraParameter 
      Caption         =   "New Parameter"
      Height          =   3780
      Left            =   705
      TabIndex        =   0
      Top             =   915
      Width           =   6300
      Begin VB.Frame fraParamType 
         Caption         =   "Type"
         Enabled         =   0   'False
         Height          =   855
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   3390
         Begin VB.CheckBox chkDefinedPerc 
            Caption         =   "Defined Values Only"
            Height          =   210
            Left            =   1560
            TabIndex        =   26
            Top             =   555
            Width           =   1755
         End
         Begin VB.CheckBox chkDefinedValue 
            Caption         =   "Defined Values Only"
            Height          =   210
            Left            =   1560
            TabIndex        =   25
            Top             =   255
            Width           =   1755
         End
         Begin VB.OptionButton optTypePerc 
            Caption         =   "Percentage %"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   535
            Width           =   1290
         End
         Begin VB.OptionButton optTypeValue 
            Caption         =   "Value"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   255
            Value           =   -1  'True
            Width           =   750
         End
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   720
         Left            =   360
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1155
         Width           =   5565
      End
      Begin VB.TextBox txtLName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   2
         Top             =   525
         Width           =   4560
      End
      Begin VB.TextBox txtSName 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   360
         MaxLength       =   5
         TabIndex        =   1
         Top             =   525
         Width           =   960
      End
      Begin VB.Frame fraEditable 
         Height          =   855
         Left            =   3780
         TabIndex        =   23
         Top             =   1920
         Width           =   1545
         Begin VB.CheckBox chkMulti 
            Caption         =   "Multiple Type"
            Height          =   195
            Left            =   105
            TabIndex        =   27
            Top             =   555
            Width           =   1260
         End
         Begin VB.CheckBox chkEditable 
            Caption         =   "is Editable"
            Height          =   195
            Left            =   105
            TabIndex        =   24
            Top             =   255
            Width           =   1020
         End
      End
      Begin VB.Frame fraValInputMethod 
         Caption         =   "Values"
         Height          =   855
         Left            =   360
         TabIndex        =   12
         Top             =   2790
         Width           =   3390
         Begin VB.OptionButton optValue 
            Caption         =   "Prompt while Expression Building"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   255
            Width           =   2610
         End
         Begin VB.OptionButton optValue 
            Caption         =   "Prompt while Executing"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   555
            Width           =   1980
         End
      End
      Begin VB.Frame fraStatus 
         Caption         =   "Status"
         Height          =   855
         Left            =   3780
         TabIndex        =   16
         Top             =   2790
         Width           =   1545
         Begin VB.OptionButton optInactive 
            Caption         =   "Inactive"
            Height          =   195
            Left            =   60
            TabIndex        =   18
            Top             =   555
            Width           =   885
         End
         Begin VB.OptionButton optActive 
            Caption         =   "Active"
            Height          =   195
            Left            =   60
            TabIndex        =   17
            Top             =   255
            Value           =   -1  'True
            Width           =   750
         End
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   945
         Width           =   795
      End
      Begin VB.Label lblLName 
         AutoSize        =   -1  'True
         Caption         =   "* Long Name"
         Height          =   195
         Left            =   1380
         TabIndex        =   5
         Top             =   315
         Width           =   930
      End
      Begin VB.Label lblSName 
         AutoSize        =   -1  'True
         Caption         =   "* Short Name"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   315
         Width           =   945
      End
   End
End
Attribute VB_Name = "frmParamMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

'Track whether the form load by the Expression Builder / by Menu
Public byExprBuilder As Boolean    'True means by Expression Builder else by Menu

'Track whether the Parameter being edited or new one
Public isEditParameter As Boolean  'True means edit else new one

'To pass the parameter ID for Editing
Public EditParamID As Integer

'For Defined Values Spread Column No
Private Enum eDef
  mVal = 1
  mStatus = 2
End Enum

Private Enum eValInput
  mBuilding = 0
  mExecuting = 1
End Enum

Private Sub chkDefinedPerc_Click()
' On Error Resume Next
  Call ChangeDefValDisplay
End Sub

Private Sub chkDefinedValue_Click()
' On Error Resume Next
  Call ChangeDefValDisplay
End Sub

Private Sub chkEditable_Click()
  
  Try
  Me.MousePointer = vbHourglass

  If chkEditable.Value = 0 Then
    If optTypeValue Then
      chkDefinedPerc.Value = 0
    Else
      chkDefinedValue.Value = 0
    End If
  End If
  
  Call ChangeDefValDisplay
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub CmdClear_Click()
' On Error Resume Next

  txtSName.Text = ""
  txtSName.Tag = ""
  txtLName.Text = ""
  txtDesc.Text = ""
  fraParamType.Enabled = True
  chkEditable.Value = 0
  optTypeValue.Value = True
  optValue(0).Value = True
  optActive.Value = True
  sprDefVal.MaxRows = 1
  sprDefVal.ClearRange 1, 1, -1, -1, True
  Spread_Lock sprDefVal, False, 1, 1, -1, -1
  sprDefPerc.MaxRows = 1
  sprDefPerc.ClearRange 1, 1, -1, -1, True
  Spread_Lock sprDefPerc, False, 1, 1, -1, -1
  chkDefinedValue.Value = 0
  chkDefinedPerc.Value = 0
  isEditParameter = False
  txtSName.SetFocus
  
End Sub

Private Function ValInputStr() As String
  Dim tmpStr As String
' On Error Resume Next
  tmpStr = ""
  If optValue(eValInput.mBuilding).Value = True Then
    tmpStr = "PB"    'Prompt while Building
  Else
    tmpStr = "PE"    'Prompt while Executing
  End If
  ValInputStr = tmpStr
End Function

Private Sub cmdClose_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Function ValidateData() As Boolean
  
  Try
  ValidateData = False
  
  If contrl_valid(txtSName, "Enter Short Name") = False Then GoTo ExitHere
  If contrl_valid(txtLName, "Enter Long Name") = False Then GoTo ExitHere
  
  If chkEditable.Value = 1 Then
    If chkDefinedValue.Value = 1 Then
      If ValidateDefValSpread(sprDefVal) = False Then
        GoTo ExitHere
      End If
    End If

    If chkDefinedPerc.Value = 1 Then
      If ValidateDefValSpread(sprDefPerc) = False Then
        GoTo ExitHere
      End If
    End If
  Else
    If optTypeValue = True Then   'Value
      If chkDefinedValue.Value = 1 Then
        If ValidateDefValSpread(sprDefVal) = False Then
          GoTo ExitHere
        End If
      End If
    Else                  'Percentage
      If chkDefinedPerc.Value = 1 Then
        If ValidateDefValSpread(sprDefPerc) = False Then
          GoTo ExitHere
        End If
      End If
    End If
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function ValidateDefValSpread(ByVal Spr As AceSpread) As Boolean
  Dim rCnt  As Long
  Dim retVal As Variant
  
  Try
  ValidateDefValSpread = False
  
    If Spr.DataRowCnt <= 0 Then
      sShowMessage "Atleast one value should be defined."
      Spread_Set_Focus IIf(Spr.Name = "sprDefVal", sprDefVal, sprDefPerc), 1, eDef.mVal
      GoTo ExitHere
    Else
      
      For rCnt = 1 To Spr.DataRowCnt
        Spr.GetText eDef.mVal, rCnt, retVal
        If Val(retVal) <= 0 Then
          sShowMessage "The Defined Value must be non-zero value"
          Spread_Set_Focus IIf(Spr.Name = "sprDefVal", sprDefVal, sprDefPerc), rCnt, eDef.mVal
          GoTo ExitHere
        End If
      Next rCnt
      
      retVal = Spr.SearchCol(eDef.mStatus, 0, -1, "1", SearchFlagsValue)
      If retVal = -1 Then
        sShowMessage "Atleast one value should be Active"
        Spread_Set_Focus IIf(Spr.Name = "sprDefVal", sprDefVal, sprDefPerc), 1, eDef.mStatus
        GoTo ExitHere
      End If
      
    End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdSave_Click()
  Dim sql As String
  Dim rs As AceADODB.Recordset
  Dim ValInput As String
  
  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  If isEditParameter = False Then   'Check for already existing Short Name
    sql = "Select ParamID from Com_FrmParamMast_Hdr "
    sql = sql & " where ParamShortName='" & Trim$(txtSName.Text) & "'"
    sql = sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    Set rs = New AceADODB.Recordset
    If RSOpen(rs, sql) = False Then
      GoTo ExitHere
    End If
    
    If Not rs.EOF Then
      MsgBox "Parameter Short Name is already available. Please give unique name."
      GoTo ExitHere
    Else
      sql = "Select Max(ParamID) from Com_FrmParamMast_Hdr"
      sql = sql & " where ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, sql) = False Then
        GoTo ExitHere
      End If
      If Not rs.EOF Then
        If IsNull(rs(0)) Then
          txtSName.Tag = 101
        ElseIf rs(0) < 101 Then
          txtSName.Tag = 101
        Else
          txtSName.Tag = rs(0) + 1
        End If
        
      End If
    End If
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  If isEditParameter = True Then
    'Delete Header Old Data
    sql = "Delete Com_FrmParamMast_Hdr Where ParamID=" & Val(txtSName.Tag)
    sql = sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
    
    'Delete Header Old Data
    sql = "Delete Com_FrmParamMast_Dtl Where ParamID=" & Val(txtSName.Tag)
    sql = sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
  End If
  
  ValInput = ValInputStr      'Get the String using function
  'Insert Header New Data
  If SaveParameterHeader(ValInput) = False Then
    GoTo ExitHere
  End If
  
  If chkEditable.Value = 1 Then
    If chkDefinedValue.Value = 1 Then
      If SaveDefinedValues(sprDefVal) = False Then
        GoTo ExitHere
      End If
    End If
    
    If chkDefinedPerc.Value = 1 Then
      If SaveDefinedValues(sprDefPerc) = False Then
        GoTo ExitHere
      End If
    End If
  Else
    If optTypeValue = True Then   'Value
      If chkDefinedValue.Value = 1 Then
        If SaveDefinedValues(sprDefVal) = False Then
          GoTo ExitHere
        End If
      End If
    Else                  'Percentage
      If chkDefinedPerc.Value = 1 Then
        If SaveDefinedValues(sprDefPerc) = False Then
          GoTo ExitHere
        End If
      End If
    End If
  End If
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
    
  If byExprBuilder = True Then
    Call frmExprBuilder.ClearParamSpread
    Call frmExprBuilder.LoadParamSpread
    GoTo UnloadHere
  Else
    If isEditParameter = True Then
      sShowMessage "Record Update Successfully."
    Else
      sShowMessage "Record Saved Successfully."
    End If
    Call CmdClear_Click
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
  Me.MousePointer = vbDefault
  Exit Sub
UnloadHere:
  Unload Me
End Try
End Sub

Private Function SaveParameterHeader(ByVal fValInput As String) As Boolean
  Dim sql As String
  
  Try
  SaveParameterHeader = False

  sql = "Execute Com_Ins_FrmParamMast_Hdr "
  sql = sql & " @ModuleCode=" & gModuleCode & ","
  sql = sql & " @ParamID=" & Val(txtSName.Tag) & ","
  sql = sql & " @ParamShortName='" & Trim$(txtSName.Text) & "',"
  sql = sql & " @ParamLongName='" & Trim$(txtLName.Text) & "',"
  sql = sql & " @ParamDesc='" & Trim$(txtDesc.Text) & "',"
  sql = sql & " @IsTypePercentage='" & IIf(optTypePerc = True, "1", "0") & "',"
  sql = sql & " @IsTypeEditable='" & chkEditable.Value & "',"
  sql = sql & " @ValInputMethod='" & fValInput & "',"
  
  If chkDefinedPerc.Value = 0 And chkDefinedValue.Value = 0 Then
    sql = sql & " @ValuesDefinedType='None',"
  ElseIf chkDefinedPerc.Value = 1 And chkDefinedValue.Value = 1 Then
    sql = sql & " @ValuesDefinedType='Both',"
  Else
    If chkDefinedPerc.Value = 1 Then
      sql = sql & " @ValuesDefinedType='Perc',"
    Else
      sql = sql & " @ValuesDefinedType='Value',"
    End If
  End If

  sql = sql & " @isMultiple='" & chkMulti.Value & "',"
  sql = sql & " @IsActive='" & IIf(optActive = True, "1", "0") & "',"
  sql = sql & " @UserID='" & gUserID & "',"
  sql = sql & " @BranchID='" & gBranchID & "',"
  sql = sql & " @CompID='" & gCompID & "'"
  If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function SaveDefinedValues(ByVal Spr As AceSpread) As Boolean
  Dim sql As String
  Dim rCnt As Long
  Dim DefValVar As String
  
  Try
  SaveDefinedValues = False

      For rCnt = 1 To Spr.DataRowCnt
        DefValVar = GetSprText(Spr, eDef.mVal, rCnt)
        If Trim$(DefValVar) <> "" Then
          sql = "Execute Com_Ins_FrmParamMast_Dtl "
          sql = sql & " @ModuleCode=" & gModuleCode & ","
          sql = sql & " @ParamID=" & Val(txtSName.Tag) & ","
          If Spr.Name = "sprDefVal" Then
            sql = sql & " @IsTypePercentage='0',"
          Else
            sql = sql & " @IsTypePercentage='1',"
          End If
          sql = sql & " @ParamValue='" & Val(DefValVar) & "',"
          sql = sql & " @IsActive='" & GetSprText(Spr, eDef.mStatus, rCnt) & "',"
          sql = sql & " @BranchID='" & gBranchID & "',"
          sql = sql & " @CompID='" & gCompID & "'"
          If DBExecuteSql(db, Sql) = False Then
 GoTo ExitHere
End If
        End If
      Next rCnt

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
  
  Try

  If byExprBuilder = True Then
    'If from ExprBuilder it should close after saving
    cmdSave.Caption = "Save && Close"
  Else
    cmdSave.Caption = "&Save"
  End If

  If isEditParameter = True Then
    'Open for Editing Parameter
    cmdClear.Enabled = False
  Else
    'Open for New Parameter
    cmdClear.Enabled = True
    fraParameter.Caption = "New Parameter"
    fraDefVal.Height = 1920
    fraDefVal.Enabled = True
    chkMulti.Enabled = True
    fraParamType.Enabled = True
    optValue(0).Value = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  If byExprBuilder = False Then
    If FormExit(frmParamMaster, UnloadMode) = False Then
      Cancel = 1
    End If
  End If
End Sub

Private Sub optTypePerc_Click()
  
  Try
  Me.MousePointer = vbHourglass

  Call ChangeDefValDisplay
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optTypeValue_Click()
  
  Try
  Me.MousePointer = vbHourglass

  Call ChangeDefValDisplay
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub optValue_Click(ByVal Index As Integer)
'  Dim PrevOption As String
'  Dim Sql As String
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  PrevOption = Trim$(fraValInputMethod.Tag)
'  Select Case Index
'    Case 0            'Supplied thru Code
'      'Not necessary since it is handled by source code
'    Case 1            'Defined Values
'      If PrevOption = "PB" Or PrevOption = "PE" Then
'        'It should show the values already used.
'        Sql = "Select"
'
'      End If
'      fraDefVal.Visible = True
'    Case 2, 3         'Prompt while building / executing
'      fraDefVal.Visible = False       'DefValues are ignored
'  End Select
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
  
End Sub

Private Sub sprDefVal_EditMode(sender as object, e as EditModeEventArgs)
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 0 Then
    If Col = eDef.mVal Then
      If GetSprText(sprDefVal, eDef.mVal, Row) <> "" Then
        sprDefVal.MaxRows = sprDefVal.DataRowCnt + 1
        If sprDefVal.DataRowCnt = Row Then
          sprDefVal.SetText eDef.mStatus, Row, "1"       'Set Active Status default
        End If
      End If
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprDefPerc_EditMode(sender as object, e as EditModeEventArgs)
  
  Try
  Me.MousePointer = vbHourglass

  If Mode = 0 Then
    If Col = eDef.mVal Then
      If Trim$(GetSprText(sprDefPerc, eDef.mVal, Row)) <> "" Then
        sprDefPerc.MaxRows = sprDefPerc.DataRowCnt + 1
        If sprDefPerc.DataRowCnt = Row Then
          sprDefPerc.SetText eDef.mStatus, Row, "1"       'Set Active Status default
        End If
      End If
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtSName_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  KeyAscii = ToUpper(KeyAscii)
  KeyAscii = CheckAlphaNumeric(KeyAscii, "@", "#", "%")  ', "[", "]", "(", ")")
End Sub

Private Sub txtSName_DblClick()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  sql = "Select ParamShortName, ParamLongName, ParamID from Com_FrmParamMast_Hdr"
  sql = sql & " Where ValInputMethod<>'SC' and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "110", "Parameters", "Short Name", "Long Name") = True Then
    Call LoadValuesForEdit(Val(getvalue(3)))
    isEditParameter = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Public Sub LoadValuesForEdit(ByVal ParamID As Integer)
  Dim sql As String
  Dim rsParam As AceADODB.Recordset
  Dim rsPMDtl As AceADODB.Recordset
  
  Try
  
  fraParameter.Caption = "Edit Parameter"
  fraParamType.Enabled = False
  
  sql = "Select ParamID, ParamShortName, ParamLongName, ParamDesc, IsTypePercentage, IsTypeEditable, ValInputMethod, ValuesDefinedType, isMultiple, isActive"
  sql = sql & " from Com_FrmParamMast_Hdr"
  sql = sql & " where ModuleCode=" & gModuleCode & " and ParamID=" & ParamID & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Set rsParam = New AceADODB.Recordset
  If RSOpen(rsParam, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsParam.EOF Then
    Call CmdClear_Click
  
    fraDefVal.Height = 1920
    txtSName.Tag = ParamID
    txtSName.Text = rsParam("ParamShortName")
    txtLName.Text = rsParam("ParamLongName")
    txtDesc.Text = rsParam("ParamDesc")
    optTypePerc.Value = rsParam("IsTypePercentage")
    chkEditable.Value = IIf(rsParam("IsTypeEditable"), "1", "0")
    chkMulti.Value = IIf(rsParam("IsMultiple"), "1", "0")
    fraValInputMethod.Tag = rsParam("ValInputMethod")   'Previous Value Input Method
    
    If Trim$(rsParam("ValuesDefinedType")) <> "None" Then
      If LoadDefValues(ParamID) = False Then       'Loading Defined Values
        sShowMessage "Unable to check/load the Defined Values"
        GoTo ExitHere
      End If
    End If
    
    Select Case rsParam("ValInputMethod")
      Case "PB"
        optValue(0).Value = True
      Case "PE"
        optValue(1).Value = True
    End Select
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsParam = Nothing
  Set rsPMDtl = Nothing
End Try
End Sub

Private Function LoadDefValues(ByVal ParamID As Integer) As Boolean
  Dim sql As String
  Dim rsPMDtl As AceADODB.Recordset
  
  Try
  LoadDefValues = False
  
  sql = "Select IsTypePercentage, ParamValue, isActive from Com_FrmParamMast_Dtl"
  sql = sql & " Where ModuleCode=" & gModuleCode & " and ParamID=" & ParamID & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  sql = sql & " Order by Convert(float, ParamValue)"
  Set rsPMDtl = New AceADODB.Recordset
  If RSOpen(rsPMDtl, sql) = False Then
    GoTo ExitHere
  End If
  
  If Not rsPMDtl.EOF Then
    Do While Not rsPMDtl.EOF
      If rsPMDtl("IsTypePercentage") = "0" Then
        sprDefVal.SetText eDef.mVal, sprDefVal.MaxRows, rsPMDtl("ParamValue")
        sprDefVal.SetText eDef.mStatus, sprDefVal.MaxRows, IIf(rsPMDtl("IsActive"), "1", "0")
        sprDefVal.MaxRows = sprDefVal.MaxRows + 1
      Else
        sprDefPerc.SetText eDef.mVal, sprDefPerc.MaxRows, rsPMDtl("ParamValue")
        sprDefPerc.SetText eDef.mStatus, sprDefPerc.MaxRows, IIf(rsPMDtl("IsActive"), "1", "0")
        sprDefPerc.MaxRows = sprDefPerc.MaxRows + 1
      End If
      rsPMDtl.MoveNext
    Loop
    
    If sprDefVal.DataRowCnt > 0 Then
      chkDefinedValue.Value = 1
    End If
    
    If sprDefPerc.DataRowCnt > 0 Then
      chkDefinedPerc.Value = 1
    End If
    
    Call LockPreviousValues(sprDefVal)
    Call LockPreviousValues(sprDefPerc)
    Call ChangeDefValDisplay
  End If
  
  fraDefVal.Enabled = True
  
  LoadDefValues = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rsPMDtl = Nothing
End Try
End Function

Private Function LockPreviousValues(ByVal Spr As AceSpread) As Boolean
  
  Try
  LockPreviousValues = False

    Spr.Tag = Spr.MaxRows - 1     'No of prev stored records
    Spr.LockBackColor = &HFFFF
    Spr.Col = eDef.mVal
Spr.Col2 = eDef.mVal
Spr.Row = 1
Spr.Row2 = Spr.MaxRows - 1
    Spr.BlockMode = True
    Spr.Lock = True
    Spr.BlockMode = False

  LockPreviousValues = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Function

Private Sub ChangeDefValDisplay()
' On Error Resume Next
  If chkDefinedValue.Value = 0 And chkDefinedPerc.Value = 0 Then
    fraDefVal.Visible = False
  ElseIf chkDefinedValue.Value = 1 And chkDefinedPerc.Value = 1 Then
    fraDefVal.Height = 3630
    sprDefVal.Top = 225
    sprDefVal.Visible = True
    sprDefPerc.Top = 1950
    sprDefPerc.Visible = True
    fraDefVal.Visible = True
  Else
    fraDefVal.Height = 1920
    If chkDefinedValue.Value = 1 Then
      sprDefVal.Top = 225
      sprDefVal.Visible = True
      sprDefPerc.Visible = False
    Else
      sprDefPerc.Top = 225
      sprDefPerc.Visible = True
      sprDefVal.Visible = False
    End If
    fraDefVal.Visible = True
  End If
End Sub

Private Sub Design_sprDefVal(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprDefVal


sprDefVal.EditMode +=  new EventHandler(sprDefVal_EditMode)

End Sub

Private Sub Design_sprDefPerc(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprDefPerc


sprDefPerc.EditMode +=  new EventHandler(sprDefPerc_EditMode)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprDefVal("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmParamMaster.frm", sprDefVal, "")

	Call Design_sprDefPerc("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmParamMaster.frm", sprDefPerc, "")

End Sub
