VERSION 5.00
Begin VB.Form frmPromptValueEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Values Entry Screen for Formula Calculation"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPartyCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   330
      TabIndex        =   14
      Top             =   7740
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   390
      Left            =   6930
      TabIndex        =   13
      Top             =   7575
      Visible         =   0   'False
      Width           =   1170
   End
Begin AceSpread sprMultiParamPrompt
      Height          =   3150
      Left            =   330
      TabIndex        =   8
      Top             =   4365
      Width           =   7755
      _Version        =   458752
      _ExtentX        =   13679
      _ExtentY        =   5556
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
   End
   Begin VB.TextBox txtScreen 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   330
      MaxLength       =   30
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   330
      Width           =   3330
   End
   Begin VB.TextBox txtFSName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   7125
      MaxLength       =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   330
      Width           =   960
   End
   Begin VB.TextBox txtFLName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3690
      MaxLength       =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   330
      Width           =   3390
   End
Begin AceSpread sprParamPrompt
      Height          =   3150
      Left            =   330
      TabIndex        =   0
      Top             =   885
      Width           =   7755
      _Version        =   458752
      _ExtentX        =   13679
      _ExtentY        =   5556
      _StockProps     =   64
      BackColorStyle  =   1
      ButtonDrawMode  =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   8
      ScrollBars      =   2
      UserResize      =   0
   End
   Begin VB.Frame fraCommand 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   2955
      TabIndex        =   10
      Top             =   7530
      Width           =   2505
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Ca&ncel"
         Height          =   390
         Left            =   1245
         TabIndex        =   11
         Top             =   180
         Width           =   1170
      End
      Begin VB.CommandButton cmdFinish 
         Caption         =   "&Finish"
         Height          =   390
         Left            =   90
         TabIndex        =   12
         Top             =   180
         Width           =   1170
      End
   End
   Begin VB.Label lblPartyCode 
      Caption         =   "Party Code"
      Height          =   195
      Left            =   330
      TabIndex        =   15
      Top             =   7545
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblMultiParameter 
      AutoSize        =   -1  'True
      Caption         =   "Give Values to Multi Parameters"
      Height          =   195
      Left            =   315
      TabIndex        =   9
      Top             =   4095
      Width           =   2250
   End
   Begin VB.Label lblParameter 
      AutoSize        =   -1  'True
      Caption         =   "Give Values to Parameters"
      Height          =   195
      Left            =   330
      TabIndex        =   7
      Top             =   675
      Width           =   1875
   End
   Begin VB.Label lblScreen 
      AutoSize        =   -1  'True
      Caption         =   "Screen Name"
      Height          =   195
      Left            =   330
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblFSName 
      AutoSize        =   -1  'True
      Caption         =   "Short Name"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   7125
      TabIndex        =   4
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblFLName 
      AutoSize        =   -1  'True
      Caption         =   "Formula Name"
      Height          =   195
      Left            =   3690
      TabIndex        =   3
      Top             =   120
      Width           =   1020
   End
End
Attribute VB_Name = "frmPromptValueEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Const cReSizeValue As Integer = 3350
Private Const cTypePerc As Double = 8#
Private Const cTypeEdit As Double = 6.25
Private Const cValue As Double = 9.13
Private Const cValueEdit As Double = 7.63

Private isFrmTest As Boolean  'To store whether it called from formula tester
Private ParamColCnt As Integer  'Per Multiple Parameter column Count
'Private byDataFill As Boolean

'Public FromScreen As eFormList

Private isCancelPressed As Boolean
Private NoDataFound As Boolean

Private Enum eParamSpr
  mParam = 1
  mLongName = 2
  mIsPerc = 3
  mIsTypeEditable = 4
  mValue = 5
  mIsValueEditable = 6
  mParamID = 7
  mValDefinedType = 8
End Enum

Private Enum eMultiParamSpr
  mIsPerc = 1
  mIsTypeEditable = 2
  mValue = 3
  mSame = 4
  mIsValueEditable = 5
End Enum

Private Enum eSprName
  mParam = 0
  mMultiParam = 1
End Enum

Private Sub LockSpread()
  Dim iCnt As Long
  
  Try
  
  If Trim$(txtScreen.Text) = "Formula Tester" Or Left$(Trim$(txtScreen.Text), 10) = "Invoice - " Then
    Call Spread_Lock(sprParamPrompt, True, eParamSpr.mIsTypeEditable, 1, eParamSpr.mIsTypeEditable, sprParamPrompt.MaxRows)
    
      Call Spread_Lock(sprMultiParamPrompt, True, iCnt + 1, 1, iCnt + 1, sprMultiParamPrompt.MaxRows)
      For iCnt = 2 To sprMultiParamPrompt.MaxCols Step ParamColCnt
        sprMultiParamPrompt.Col = iCnt + 1
sprMultiParamPrompt.ColHidden = True

'        .Col = iCnt + 4: .ColHidden = True
      Next iCnt
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub PrepareSpread(ByVal MultiItemList() As String,Optional ByVal isMultiShow As Boolean = False)
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim ShowMultiParam As Boolean
  Dim RowLimit As Long
  Dim iCnt As Long
  Dim ItemList() As String

  Try
  
  If Trim$(txtScreen.Text) = "Formula Tester" Or Left$(Trim$(txtScreen.Text), 10) = "Invoice - " Then
'    With sprExprPrompt
'      .Col = eExprSpr.mIsValueEditable: .ColHidden = True
'    End With
    
      sprParamPrompt.Col = eParamSpr.mIsTypeEditable
sprParamPrompt.ColHidden = True

      sprParamPrompt.Col = eParamSpr.mIsValueEditable
sprParamPrompt.ColHidden = True    
    If Trim$(txtScreen.Text) = "Formula Tester" Then
      isFrmTest = True
    Else
      isFrmTest = False
    End If
    ParamColCnt = 4
  Else
    isFrmTest = False
    ParamColCnt = 5
  End If
  
  ShowMultiParam = isMultiShow
  
  If ShowMultiParam = True Then
    If isFrmTest = True Then
      'Required to check the no of multiple entry rows
      RowLimit = 0
      Do While RowLimit <= 0
        RowLimit = Val(InputBox("No. of Rows you required for Multi Parameters."))
      Loop
      sprMultiParamPrompt.MaxRows = RowLimit
    Else
      Sql = "Select Max(MaxCnt) as MaxCnt from"
      Sql = Sql & " (Select a.Block, Count(a.KeyCol) MaxCnt"
      Sql = Sql & " from COM_tmpFormulaCalc_Dtl a Left Outer Join COM_FrmParamMast_Hdr b"
      Sql = Sql & " On a.ModuleCode=b.ModuleCode and a.Block=b.ParamID and a.BranchID=b.BranchID and a.CompID=b.CompID"
      Sql = Sql & " Where b.isMultiple='1' and a.DTKey='" & Trim$(txtScreen.Tag) & "' and a.FrmID=" & Val(txtFSName.Tag)
      Sql = Sql & " and a.BlockType='Param'"
      Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
      Sql = Sql & " Group by a.Block) x"
      Set rs = New AceADODB.Recordset
      If RSOpen(rs, Sql) = False Then
        GoTo ExitHere
      End If
      
      If Not rs.EOF Then
        If rs(0) > 0 Then
          RowLimit = rs(0)
        Else
          ShowMultiParam = False
        End If
      Else
        ShowMultiParam = False
      End If
    End If
    
  End If
  
  If ShowMultiParam = False Then
    'Hide Multi Parameter Spread & resize the form
    Call HidePromptSpread(eSprName.mMultiParam)
  Else
      sprMultiParamPrompt.MaxRows = RowLimit
      'Span the Column Itemcode header
      sprMultiParamPrompt.AddCellSpan SpreadHeader, SpreadHeader, 1, 2
      sprMultiParamPrompt.AddCellSpan 1, SpreadHeader, 1, 2
      
      If Trim$(txtScreen.Text) <> "Formula Tester" Then
        For iCnt = 0 To UBound(MultiItemList)
          Erase ItemList
          ItemList = Split(MultiItemList(iCnt), Chr(145), 2, vbTextCompare)
          sprMultiParamPrompt.SetText 1, Val(ItemList(0)), Trim$(ItemList(1))
  '        .SetText 1, iCnt + 1, Trim$(ItemList(1)) 'by ramana on 16/11/06
  '        Spread_Set_CellTag sprMultiParamPrompt, Val(ItemList(0)), 1, iCnt + 1 'by ramana on 16/11/06
        Next iCnt
      
        sprMultiParamPrompt.MaxRows = sprMultiParamPrompt.DataRowCnt
      End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Sub

Private Sub HidePromptSpread(ByVal fSpread As eSprName)

  Try

  Select Case fSpread
    Case eSprName.mMultiParam
      sprMultiParamPrompt.Visible = False
      lblMultiParameter.Visible = False
      fraCommand.Top = fraCommand.Top - cReSizeValue
      Me.Height = Me.Height - cReSizeValue
    Case eSprName.mParam
      sprParamPrompt.Visible = False
      lblParameter.Visible = False
      
      sprMultiParamPrompt.Top = sprMultiParamPrompt.Top - cReSizeValue
      lblMultiParameter.Top = lblMultiParameter.Top - cReSizeValue
      fraCommand.Top = fraCommand.Top - cReSizeValue
      Me.Height = Me.Height - cReSizeValue
  End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub ShowDefinedValueHelp(ByRef fSpr As AceSpread,ByVal  fCol As Long,ByVal  fRow As Long,Optional ByVal isByEditMode As Boolean = False)
  Dim retVal As Variant
  Dim Sql As String
  Dim sParamID As Integer
  Dim sIsPerc As Integer
  Dim SCol As Long
  
  Try
  
    If isByEditMode Then
      SCol = fCol - 2
    Else
      SCol = fCol
    End If
  
    If fSpr.name = "sprParamPrompt" Then
      fSpr.GetText eParamSpr.mParamID, fRow, retVal
      sParamID = Val(retVal)
      
      fSpr.GetText eParamSpr.mIsPerc, fRow, retVal
      sIsPerc = Val(retVal)
    ElseIf fSpr.name = "sprMultiParamPrompt" Then
      sParamID = Val(Spread_Get_CellTag(fSpr, (SCol - 1) + eMultiParamSpr.mIsPerc, SpreadHeader))
      
      fSpr.GetText (SCol - 1) + eMultiParamSpr.mIsPerc, fRow, retVal
      sIsPerc = Val(retVal)
    Else
      GoTo ExitHere
    End If
    
    Sql = "Select ParamValue from Com_FrmParamMast_Dtl "
    Sql = Sql & " where ParamID=" & sParamID
    Sql = Sql & " and isTypePercentage='" & sIsPerc & "'"
    Sql = Sql & " and ModuleCode = " & gModuleCode & ""
    Sql = Sql & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
    Sql = Sql & " Order by Convert(float, ParamValue)"
    If FetchValue_New(False, db.ConnectionString, Sql, "1", "Defined Values List", "Values") = True Then
      fSpr.SetText SCol + 2, fRow, Val(getvalue(1))
      fSpr.EditMode = False
    Else
      fSpr.SetText SCol + 2, fRow, ""
    End If
    
    
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_Load()

	Call InitializeSpreads()
Exit Sub
' On Error Resume Next
  isCancelPressed = False
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
' On Error Resume Next
  
  If UnloadMode = vbFormControlMenu Then
    isCancelPressed = True
  End If
End Sub

Private Sub SearchHideOrShow(ByRef fpSpr As AceSpread,ByVal  SearchCol As Long,ByVal  SearchTerm As String,ByVal  MakeHidden As Boolean,Optional ByVal StartRow As Long = 1,Optional ByVal EndRow As Long = -2)
  Dim rCnt As Long
  Dim retVal As Variant
  
  Try

    For rCnt = StartRow To IIf(EndRow = -2, fpSpr.DataRowCnt, EndRow)
      fpSpr.GetText SearchCol, rCnt, retVal
      If Trim$(retVal) = SearchTerm Then
        fpSpr.Row = rCnt
fpSpr.RowHidden = MakeHidden
      End If
    Next rCnt
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

'Private Sub sprExprPrompt_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
'  Dim retVal As Variant
'
'  On Error GoTo ErrHandler
'  Me.MousePointer = vbHourglass
'
'  With sprExprPrompt
'    If Col = eExprSpr.mValue Then
'      If Mode = 0 Then
'        If ChangeMade = True Then
'          .GetText eExprSpr.mValue, Row, retVal
'          If Trim$(retVal) = "" Then
'            .GetText eExprSpr.mExprSeq, Row, retVal
''            Call SearchHideOrShow(sprParamPrompt, eParamSpr.mExprSeq, Val(retVal), False)
'          Else
'            .GetText eExprSpr.mExprSeq, Row, retVal
''            Call SearchHideOrShow(sprParamPrompt, eParamSpr.mExprSeq, Val(retVal), True)
'          End If
'        End If
'      End If
'    End If
'  End With
'
'ExitHere:
'  On Error GoTo 0
'  Me.MousePointer = vbDefault
'  Exit Sub
'ErrHandler:
'  Call ErrMsgShow
'  Resume ExitHere
'End Sub

'Private Sub sprExprPrompt_KeyUp(KeyCode As Integer, Shift As Integer)
'
'  On Error Resume Next
'  With sprExprPrompt
'    If .ActiveCol = eExprSpr.mValue Then                'In Value Cell
'      If .EditMode = False Then                         'When not in Edit Mode
'        If KeyCode = vbKeyDelete Then                   'Delete Key Pressed
'          .SetText eExprSpr.mValue, .ActiveRow, ""      'Set Empty String
'          Call sprExprPrompt_EditMode(eExprSpr.mValue, .ActiveRow, 0, True)
'        End If
'      End If
'    End If
'  End With
'
'End Sub

Private Sub sprParamPrompt_ButtonClicked(sender as object, e as ButtonClickedEventArgs)
  Dim tmpVar As Variant
  Dim retVal As Variant
  Dim retRow As Integer
  
  Try
  Me.MousePointer = vbHourglass
  
    Select Case Col
      Case eParamSpr.mIsPerc
        sprParamPrompt.GetText eParamSpr.mValDefinedType, Row, tmpVar
        sprParamPrompt.GetText eParamSpr.mIsPerc, Row, retVal
        If Val(retVal) = 0 Then     'Value
          If Trim$(tmpVar) = "Both" Or Trim$(tmpVar) = "Value" Then
            Call ShowDefinedValueHelp(sprParamPrompt, Col, Row)
          Else
            sprParamPrompt.SetText eParamSpr.mValue, Row, ""
          End If
        Else                        'Percentage
          If Trim$(tmpVar) = "Both" Or Trim$(tmpVar) = "Perc" Then
            Call ShowDefinedValueHelp(sprParamPrompt, Col, Row)
          Else
            sprParamPrompt.SetText eParamSpr.mValue, Row, ""
          End If
        End If
        
'        If Trim$(tmpVar) <> "None" Then
'          Call ShowDefinedValueHelp(sprParamPrompt, Col, Row)
'        Else
'          .SetText eParamSpr.mValue, Row, ""
'        End If
'      Case eParamSpr.mSameValue
'        .GetText eParamSpr.mParamID, Row, tmpVar
'        retRow = .SearchCol(eParamSpr.mParamID, Row, -1, Trim$(tmpVar), SearchFlagsValue)
'        Do While retRow <> -1
'          .SetText eParamSpr.mValue, retRow, SpVal(sprParamPrompt, Row, eParamSpr.mValue)
'          retRow = .SearchCol(eParamSpr.mParamID, retRow, -1, Trim$(tmpVar), SearchFlagsValue)
'        Loop
    End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub sprMultiParamPrompt_ButtonClicked(sender as object, e as ButtonClickedEventArgs)
  Dim tmpVar As Variant
  Dim CurValue As Double
  Dim rCnt As Integer
  Dim CurCol As Long
  Dim ParamFirstCol As Long
  
  Try
  Me.MousePointer = vbHourglass
  
  CurCol = ((Col - 1) Mod ParamColCnt)
  CurCol = IIf(CurCol = 0, ParamColCnt, CurCol)
  
  ParamFirstCol = (Col - (CurCol - 1))
  
    Select Case CurCol
      Case eMultiParamSpr.mIsPerc
        'To Get Val Input Method
        tmpVar = Spread_Get_CellTag(sprMultiParamPrompt, ParamFirstCol, SpreadHeader + 1)
        If Trim$(tmpVar) <> "None" Then
          Call ShowDefinedValueHelp(sprMultiParamPrompt, Col, Row)
        Else
          sprMultiParamPrompt.SetText (ParamFirstCol - 1) + eMultiParamSpr.mValue, Row, ""
          
          sprMultiParamPrompt.GetText ParamFirstCol, Row, tmpVar
          If Val(tmpVar) = 1 Then
            Spread_Set_Max sprMultiParamPrompt, 100, (ParamFirstCol - 1) + eMultiParamSpr.mValue, Row, (ParamFirstCol - 1) + eMultiParamSpr.mValue, Row
          Else
            Spread_Set_Max sprMultiParamPrompt, 99999999, (ParamFirstCol - 1) + eMultiParamSpr.mValue, Row, (ParamFirstCol - 1) + eMultiParamSpr.mValue, Row
          End If
          
        End If
      Case eMultiParamSpr.mSame
        sprMultiParamPrompt.GetText (ParamFirstCol - 1) + eMultiParamSpr.mValue, Row, tmpVar
        If Trim$(tmpVar) <> "" Then
          CurValue = Val(tmpVar)
          For rCnt = 1 To sprMultiParamPrompt.MaxRows
            sprMultiParamPrompt.GetText (ParamFirstCol - 1) + eMultiParamSpr.mValue, rCnt, tmpVar
            If Trim$(tmpVar) = "" Then
              sprMultiParamPrompt.SetText (ParamFirstCol - 1) + eMultiParamSpr.mValue, rCnt, CurValue
            End If
          Next rCnt
        End If
    End Select
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Public Function GetPromptValuesHelp(ByVal fPartyCode As Integer,ByVal  fFormulaID As Integer,ByVal  fUserKey As String,ByVal  fPrevUserKey As String,ByVal  fCallScreenName As String,ByVal  MultiItemList() As String,Optional ByVal isSupplyThruCodeInclude As Boolean = False,Optional ByVal LockRefType As String = "",Optional ByVal LockRefNo As Integer = 0,Optional ByVal LockRefDate As Date = "1900-01-01",Optional ByVal ValRefType As String = "",Optional ByVal ValRefNo As Integer = 0,Optional ByVal ValRefDate As Date = "1900-01-01",Optional ByVal VendorCode As Integer = 0) As VbTriState



'        Optional MultiItemSpread As fpSpread, Optional MultiItemCol As Long = 1, 

  Try
  GetPromptValuesHelp = vbFalse
  
  If LoadFormulaName(fPartyCode, fFormulaID, fUserKey, fPrevUserKey, fCallScreenName) = False Then
    GoTo ExitHere
  End If
  
  If LoadPromptInSpread(fPartyCode, fFormulaID, MultiItemList, isSupplyThruCodeInclude, fPrevUserKey, LockRefType, LockRefNo, LockRefDate, ValRefType, ValRefNo, ValRefDate) = False Then
    GoTo ExitHere
  Else
    If NoDataFound = True Then
      GetPromptValuesHelp = vbUseDefault
      Unload Me
      GoTo ExitHere
    Else
      Me.MousePointer = vbDefault
      Me.Show vbModal
      If isCancelPressed = True Then
        sShowMessage "Calculation canceled by User."
        GoTo ExitHere
      Else
          Call ClearTmpPromptTable(fPrevUserKey)
      End If
    End If
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function FillMultiParamDetails(ByVal fParamFirstCol As Long,ByVal  fRow As Long,ByVal  isTypePerc As Boolean,ByVal  isTypeEdit As Boolean,ByVal  isTypeEditLock As Boolean,ByVal  ValueUsed As Double,Optional ByVal isValueEdit As Boolean = False) As Boolean
 
  Try
  FillMultiParamDetails = False
  
    sprMultiParamPrompt.SetText fParamFirstCol, fRow, IIf(isTypePerc, "1", "0")
    sprMultiParamPrompt.SetText fParamFirstCol + 1, fRow, IIf(isTypeEdit, "1", "0")
    If isTypeEditLock = True Then
      Call Spread_Lock(sprMultiParamPrompt, False, fParamFirstCol, fRow, fParamFirstCol + 1, fRow)
      Call Spread_BkColour(sprMultiParamPrompt, "&HFFFFFF", fParamFirstCol, fRow, fParamFirstCol + 1)
    Else
      Call Spread_Lock(sprMultiParamPrompt, True, fParamFirstCol, fRow, fParamFirstCol + 1, fRow)
      Call Spread_BkColour(sprMultiParamPrompt, "&H555555", fParamFirstCol, fRow, fParamFirstCol + 1)
    End If
    If ValueUsed <> 0 Then
      sprMultiParamPrompt.SetText fParamFirstCol + 2, fRow, ValueUsed
    End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Function LoadPromptInSpread(ByVal fPartyCode As Integer,ByVal  fFormulaID As Integer,ByVal  MultiItemList() As String,Optional ByVal isSupplyThruCodeInclude As Boolean = False,Optional ByVal fUserKey As String,Optional ByVal LockRefType As String,Optional ByVal LockRefNo As Integer,Optional ByVal LockRefDate As Date,Optional ByVal ValRefType As String,Optional ByVal ValRefNo As Integer,Optional ByVal ValRefDate As Date) As Boolean


  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim rCnt As Long
  Dim tmpVal As String
  Dim cCnt As Long
  Dim cParam As Integer
  Dim i As Long
  Dim retVal As Variant
  
  Dim MultiDataFound As Boolean
  Dim ParamDataFound As Boolean
  
  Try
  LoadPromptInSpread = False
  
  Sql = "Execute COM_GetMultiParameter_Prompt @ModuleCode=" & gModuleCode & ", @FrmID=" & fFormulaID & ", @SCParamInclude='" & IIf(isSupplyThruCodeInclude, 1, 0) & "', @UserKey='" & fUserKey & "',"
  Sql = Sql & " @LockRefType='" & LockRefType & "', @LockRefNo=" & LockRefNo & ", @LockRefDate='" & Format(LockRefDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @ValRefType='" & ValRefType & "', @ValRefNo=" & ValRefNo & ", @ValRefDate='" & Format(ValRefDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @VendorCode=" & fPartyCode & ","
  Sql = Sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "', @Finyear='" & gFinyear & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    Call PrepareSpread(MultiItemList, False)   'Hide Multi Parameter Spread
    MultiDataFound = False
  Else
    Call PrepareSpread(MultiItemList, True)    'show Multi Parameter Spread
    
    rs.MoveLast
rs.MoveFirst
    
      cParam = 0
rCnt = 0
      Do While Not rs.EOF
        If cParam <> rs("ParamID") Then
          '-------------Add Parameter Details column Starts Here
          cCnt = sprMultiParamPrompt.MaxCols + 1
          sprMultiParamPrompt.MaxCols = sprMultiParamPrompt.MaxCols + ParamColCnt
          
          sprMultiParamPrompt.AddCellSpan cCnt, SpreadHeader, ParamColCnt, 1
          sprMultiParamPrompt.SetText cCnt, SpreadHeader, "[" & rs("ParamShortName") & "]" & " " & rs("ParamLongName")
          Spread_Set_CellTag sprMultiParamPrompt, rs("ParamID"), cCnt, SpreadHeader
          Spread_Set_CellTag sprMultiParamPrompt, rs("ValuesDefinedType"), cCnt, SpreadHeader + 1
          
          'Column Width
          sprMultiParamPrompt.ColWidth(cCnt) = cTypePerc
          sprMultiParamPrompt.ColWidth(cCnt + 1) = cTypeEdit
          sprMultiParamPrompt.ColWidth(cCnt + 2) = cValue
          sprMultiParamPrompt.ColWidth(cCnt + 3) = cTypeEdit
          If isFrmTest = False Then
            sprMultiParamPrompt.ColWidth(cCnt + 4) = cValueEdit
          End If
          
          'Column Type
          sprMultiParamPrompt.Col = cCnt
sprMultiParamPrompt.Row = -1
sprMultiParamPrompt.CellType = CellTypeCheckBox
sprMultiParamPrompt.TypeCheckType = TypeCheckTypeNormal
sprMultiParamPrompt.TypeCheckCenter = True

          sprMultiParamPrompt.Col = cCnt + 1
sprMultiParamPrompt.Row = -1
sprMultiParamPrompt.CellType = CellTypeCheckBox
sprMultiParamPrompt.TypeCheckType = TypeCheckTypeNormal
sprMultiParamPrompt.TypeCheckCenter = True

          sprMultiParamPrompt.Col = cCnt + 2
sprMultiParamPrompt.Row = -1
sprMultiParamPrompt.CellType = CellTypeNumber
sprMultiParamPrompt.TypeNumberDecPlaces = 2
sprMultiParamPrompt.TypeNumberMin = 0

          sprMultiParamPrompt.Col = cCnt + 3
sprMultiParamPrompt.Row = -1
sprMultiParamPrompt.CellType = CellTypeButton
sprMultiParamPrompt.TypeButtonText = "Apply"
sprMultiParamPrompt.TypeButtonAlign = TypeButtonAlignLeft
          If isFrmTest = False Then
            sprMultiParamPrompt.Col = cCnt + 4
sprMultiParamPrompt.Row = -1
sprMultiParamPrompt.CellType = CellTypeCheckBox
sprMultiParamPrompt.TypeCheckType = TypeCheckTypeNormal
sprMultiParamPrompt.TypeCheckCenter = True
          End If
          
          'Column Heading
          sprMultiParamPrompt.SetText cCnt, SpreadHeader + 1, "is Type Percentage"
          sprMultiParamPrompt.SetText cCnt + 1, SpreadHeader + 1, "is Type Editable"
          sprMultiParamPrompt.SetText cCnt + 2, SpreadHeader + 1, "Value"
          sprMultiParamPrompt.SetText cCnt + 3, SpreadHeader + 1, "Same to Rest"
          If isFrmTest = False Then
            sprMultiParamPrompt.SetText cCnt + 4, SpreadHeader + 1, "is Value Editable"
          End If
          
          Call Spread_Lock(sprMultiParamPrompt, True, cCnt + 1, 1, cCnt + 1, sprMultiParamPrompt.MaxRows)
          sprMultiParamPrompt.Col = cCnt
sprMultiParamPrompt.Row = -1
sprMultiParamPrompt.Text = 1
          '-------------Add Parameter Details column Ends Here
        
          cParam = rs("ParamID")
        End If
        
        Do While cParam = rs("ParamID")
          rCnt = sprMultiParamPrompt.SearchCol(1, 0, -1, CheckNull(rs("KeyCol"), 0), SearchFlagsNone)  'Search in ItemCode
          
          If rCnt > 0 Then
            If FillMultiParamDetails(cCnt, rCnt, rs("tmpTypePercentage"), rs("tmpTypeEditable"), rs("LockTypeEditable"), rs("tmpValueUsed")) = False Then
              GoTo ExitHere
            End If
          End If
          rs.MoveNext
          If rs.EOF Then
 Exit Do
End If
        Loop
        
'        For i = 1 To sprMultiParamPrompt.DataRowCnt
'          If Val(Spread_Get_CellTag(sprMultiParamPrompt, 1, i)) = Val(rs("MultiSeq")) Then
'            If FillMultiParamDetails(cCnt, i, rs("tmpTypePercentage"), rs("tmpTypeEditable"), rs("LockTypeEditable"), rs("tmpValueUsed")) = False Then
'              GoTo ExitHere
'            End If
'          End If
'        Next i
'        rs.MoveNext

'      'If the available entry of Last Parameter short for the total rows then fill the previous entry to all
'      If rCnt <> .MaxRows And rCnt <> 0 Then
'        rs.MovePrevious
'        Do While rCnt < .MaxRows
'          rCnt = rCnt + 1
'          If FillMultiParamDetails(cCnt, rCnt, rs("tmpTypePercentage"), rs("tmpTypeEditable"), rs("LockTypeEditable"), rs("tmpValueUsed")) = False Then
'            GoTo ExitHere
'          End If
'        Loop
'        rs.MoveNext
'      End If
      Loop    
    MultiDataFound = True
  End If
  
  'Parameter Prompt Spread fill
  Sql = "Execute COM_GetParameter_Prompt @ModuleCode=" & gModuleCode & ", @FrmID=" & fFormulaID & ", @PartyCode=" & Val(txtPartyCode) & ", @SCParamInclude='" & IIf(isSupplyThruCodeInclude, 1, 0) & "', @UserKey='" & fUserKey & "',"
  Sql = Sql & " @LockRefType='" & LockRefType & "', @LockRefNo=" & LockRefNo & ", @LockRefDate='" & Format(LockRefDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @ValRefType='" & ValRefType & "', @ValRefNo=" & ValRefNo & ", @ValRefDate='" & Format(ValRefDate, "yyyy-MM-dd") & "',"
  Sql = Sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "', @Finyear='" & gFinyear & "'"
  
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  
  If rs.EOF Then
    ParamDataFound = False
    Call HidePromptSpread(eSprName.mParam)
  Else
    rs.MoveLast
rs.MoveFirst
    NoDataFound = False
    ParamDataFound = True
      sprParamPrompt.MaxRows = rs.RecordCount
      For rCnt = 1 To rs.RecordCount
        sprParamPrompt.SetText eParamSpr.mParam, rCnt, rs("ParamShortName")
        sprParamPrompt.SetText eParamSpr.mLongName, rCnt, rs("ParamLongName")
        sprParamPrompt.SetText eParamSpr.mParamID, rCnt, rs("ParamID")
        sprParamPrompt.SetText eParamSpr.mValDefinedType, rCnt, rs("ValuesDefinedType")
        sprParamPrompt.SetText eParamSpr.mIsPerc, rCnt, IIf(rs("tmpTypePercentage"), "1", "0")
        sprParamPrompt.SetText eParamSpr.mIsTypeEditable, rCnt, IIf(rs("tmpTypeEditable"), "1", "0")
        If rs("LockTypeEditable") Then
          Call Spread_Lock(sprParamPrompt, False, eParamSpr.mIsPerc, rCnt, eParamSpr.mIsTypeEditable, rCnt)
          Call Spread_BkColour(sprParamPrompt, "&HFFFFFF", eParamSpr.mIsPerc, rCnt, eParamSpr.mIsTypeEditable)
        Else
          Call Spread_Lock(sprParamPrompt, True, eParamSpr.mIsPerc, rCnt, eParamSpr.mIsTypeEditable, rCnt)
          Call Spread_BkColour(sprParamPrompt, "&H555555", eParamSpr.mIsPerc, rCnt, eParamSpr.mIsTypeEditable)
        End If
        sprParamPrompt.GetText eParamSpr.mValue, rCnt, retVal
        If Trim$(retVal) = "" Then
          sprParamPrompt.SetText eParamSpr.mValue, rCnt, rs("tmpValueUsed")
        End If
        
        rs.MoveNext
      Next rCnt
  End If
  
  If MultiDataFound = False And ParamDataFound = False Then
    NoDataFound = True
'  Else
'    'Expression Prompt Spread Fill
'    Sql = "Execute COM_GetExpression_Prompt @ModuleCode=" & gModuleCode & ", @FrmID=" & fFormulaID & ", @SCParamInclude='" & IIf(isSupplyThruCodeInclude, 1, 0) & "', @UserKey='" & fUserKey & "',"
'    Sql = Sql & " @LockRefType='" & LockRefType & "', @LockRefNo=" & LockRefNo & ", @LockRefDate='" & Format(LockRefDate, "yyyy-MM-dd") & "',"
'    Sql = Sql & " @ValRefType='" & ValRefType & "', @ValRefNo=" & ValRefNo & ", @ValRefDate='" & Format(ValRefDate, "yyyy-MM-dd") & "',"
'    Sql = Sql & " @BranchID='" & gBranchID & "', @CompID='" & gCompID & "', @Finyear='" & gFinyear & "'"
'
'    Set rs = New ADODB.Recordset
'    If RSOpen(rs, Sql) = False Then
'      GoTo ExitHere
'    End If
'
'    If Not rs.EOF Then
'      rs.MoveLast: rs.MoveFirst
'      With sprExprPrompt
'        .MaxRows = rs.RecordCount
'        For rCnt = 1 To rs.RecordCount
'          .SetText eExprSpr.mExprName, rCnt, rs("ExprName")
'          .SetText eExprSpr.mExpression, rCnt, GetFrontEndExpr(fFormulaID, rs("ExprSeq"))
'          tmpVal = CheckNull(rs("Value"))
'          .SetText eExprSpr.mValue, rCnt, tmpVal
'          .SetText eExprSpr.mIsValueEditable, rCnt, IIf(CheckNull(rs("isValueEditable"), "0") = "1", "1", "0")
'          .SetText eExprSpr.mExprSeq, rCnt, rs("ExprSeq")
'          rs.MoveNext
'        Next rCnt
'      End With
'    End If
    
'    Call LockSpread
  End If
  
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function LoadFormulaName(ByVal fPartyCode As Integer,ByVal  fFormulaID As Integer,ByVal  fUserKey As String,ByVal  fPrevUserKey As String,ByVal  fCallScreenName As String) As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  
  Try
  LoadFormulaName = False
  
  Sql = "Select FrmLName, FrmSName from Com_FormulaMast_Hdr"
  Sql = Sql & " where FrmID=" & fFormulaID
  Sql = Sql & " and ModuleCode=" & gModuleCode & " and BranchID='" & gBranchID & "' and CompID='" & gCompID & "'"
  Set rs = New AceADODB.Recordset
  If RSOpen(rs, Sql) = False Then
    GoTo ExitHere
  End If
  If Not rs.EOF Then
    txtScreen.Text = fCallScreenName
    txtScreen.Tag = fUserKey
    
    txtFLName.Text = rs("FrmLName")
    txtFLName.Tag = fPrevUserKey
    
    txtFSName.Text = rs("FrmSName")
    txtFSName.Tag = fFormulaID
    
    txtPartyCode.Text = fPartyCode
  Else
    sShowMessage "No Records Found"
    GoTo ExitHere
  End If

Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Function ValidateData() As Boolean
  Dim Sql As String
  Dim rs As AceADODB.Recordset
  Dim cCnt As Long
  Dim rCnt As Long
  Dim retVal As Variant
  Dim ParamDataEmpty As Boolean
  Dim Pos As Long
  
  Try
  ValidateData = False
  
    If sprParamPrompt.DataRowCnt <> 0 Then
      ParamDataEmpty = False
      For rCnt = 1 To sprParamPrompt.DataRowCnt
        sprParamPrompt.GetText eParamSpr.mValue, rCnt, retVal
        If Len(Trim$(retVal)) = 0 Then
          ParamDataEmpty = True
          Exit For
        End If
      Next rCnt
    
      If ParamDataEmpty = True Then
        Sql = "Select dbo.fn_GetExprName(a.ModuleCode, a.FrmID, a.ExprSeq, a.BranchID, a.CompID) as ExprName,  a.ExprSeq, a.Block, "
        Sql = Sql & " (Select Count(b.Block) from COM_FormulaMast_Block_Dtl b"
        Sql = Sql & " Where b.Block=a.Block and b.ExprSeq=a.ExprSeq and b.FrmID=a.FrmID and b.ModuleCode=a.ModuleCode and b.BranchID=a.BranchID and b.CompID=a.CompID) as ParamCnt"
        Sql = Sql & " from COM_FormulaMast_Block_Dtl a"
        Sql = Sql & " where a.BlockType='Param' and a.FrmID=" & txtFSName.Tag
        Sql = Sql & " and a.ModuleCode=" & gModuleCode & " and a.BranchID='" & gBranchID & "' and a.CompID='" & gCompID & "'"
    
        Set rs = New AceADODB.Recordset
        If RSOpen(rs, Sql) = False Then
          GoTo ExitHere
        End If
        
        If rs.EOF Then
          sShowMessage "Validation Failed."
          GoTo ExitHere
        End If
        
'        For rCnt = 1 To .DataRowCnt
'          .GetText eParamSpr.mValue, rCnt, retVal
'          If Len(Trim$(retVal)) = 0 Then
'            .GetText eParamSpr.mParamID, rCnt, retVal
'
'            rs.Filter = adFilterNone
'            rs.Filter = "Block=" & Val(retVal)
'            Sql = ""
'            If Not rs.EOF Then
'              Do While Not rs.EOF
'                Pos = sprExprPrompt.SearchCol(eExprSpr.mExprName, 1, -1, rs("ExprName"), SearchFlagsNone)
'                If Pos <> -1 Then
'                  Sql = IIf(Sql = "", "'", ", '") & rs("ExprName") & "'"
'                End If
'                rs.MoveNext
'              Loop
'              If Trim$(Sql) <> "" Then
'                .GetText eParamSpr.mParam, rCnt, retVal
'                sShowMessage Trim$(retVal) & " Parameter used in the following Expressions (" & Sql & ") for which the value not supplied. Give value for the Expression or the Parameter."
'                GoTo ExitHere
'              End If
'            End If
'          End If
'        Next rCnt
          
      End If
      
    End If
  
'  With sprMultiParamPrompt
'    For cCnt = 4 To .MaxCols Step 4
'      For rCnt = 1 To .DataRowCnt
'        .GetText cCnt, rCnt, retVal
'        If Trim$(retVal) = "" Then
'          .GetText cCnt - 2, SpreadHeader, retVal
'          sShowMessage "Value not supplied for the Multi Parameter '" & Trim$(retVal) & "' in the Row " & rCnt
'          GoTo ExitHere
'        End If
'      Next rCnt
'    Next cCnt
'  End With
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set rs = Nothing
End Try
End Function

Private Sub cmdCancel_Click()
' On Error Resume Next
  isCancelPressed = True
  Unload Me
End Sub

Private Sub cmdClear_Click()
' On Error Resume Next
    sprParamPrompt.ClearRange 5, 1, 5, sprParamPrompt.DataRowCnt, True
End Sub

Private Sub cmdFinish_Click()
  Dim cCnt As Long
  Dim rCnt As Long
  Dim retVal As Variant
  Dim sFrmID As Integer
  Dim sExprSeq As Integer
  Dim sParamID As Integer
  Dim sIsPerc As Boolean
  Dim sIsTypeEdit As Boolean
  Dim SValue As Double
  Dim sIsValueEdit As Boolean
  Dim isInvoiceScreen As Boolean
  
  Try
  Me.MousePointer = vbHourglass
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  If Left$(Trim$(txtScreen.Text), 10) = "Invoice - " Then
    isInvoiceScreen = True
  Else
    isInvoiceScreen = False
  End If
  
  sFrmID = Val(txtFSName.Tag)  'Formula ID
'  With sprExprPrompt
'    For rCnt = 1 To .DataRowCnt
'      .GetText eExprSpr.mValue, rCnt, retVal
'      If Trim$(retVal) <> "" Then
'        .GetText eExprSpr.mExprSeq, rCnt, retVal
'        sExprSeq = Val(retVal)   'Expr Seq
'        sIsPerc = False
'        sIsTypeEdit = False
'        .GetText eExprSpr.mValue, rCnt, retVal
'        SValue = Val(retVal)   'Value Used
'
'        If isInvoiceScreen Then
'          sIsValueEdit = False
'        Else
'          .GetText eExprSpr.mIsValueEditable, rCnt, retVal
'          sIsValueEdit = IIf(retVal = "", False, True)
'        End If
'
'        If InsertPromptValues(Trim$(txtScreen.Tag), sFrmID, "Expr", CStr(sExprSeq), 0, sIsPerc, sIsTypeEdit, SValue, sIsValueEdit) = False Then
'          GoTo ExitHere
'        End If
'
'      End If
'    Next rCnt
'  End With
  
    For rCnt = 1 To sprParamPrompt.DataRowCnt
      sprParamPrompt.Row = rCnt               'If the Row is not visible then it will be handled by the expression values
      If sprParamPrompt.RowHidden = False Then
        sprParamPrompt.GetText eParamSpr.mParamID, rCnt, retVal
        sParamID = Val(retVal)       'ParamID
        sprParamPrompt.GetText eParamSpr.mIsPerc, rCnt, retVal
        sIsPerc = IIf(Trim$(retVal) = "1", True, False) 'Is Type Percentage
        
        If isInvoiceScreen Then
          sIsTypeEdit = False
          sIsValueEdit = False
        Else
          sprParamPrompt.GetText eParamSpr.mIsTypeEditable, rCnt, retVal
          sIsTypeEdit = IIf(Trim$(retVal) = "1", True, False) 'Is Type Editable
          sprParamPrompt.GetText eParamSpr.mIsValueEditable, rCnt, retVal
          sIsValueEdit = IIf(Trim$(retVal) = "1", True, False)
        End If
        
        sprParamPrompt.GetText eParamSpr.mValue, rCnt, retVal
        SValue = Val(retVal)       'Value Used
        
        If InsertPromptValues(Trim$(txtScreen.Tag), sFrmID, "Param", CStr(sParamID), "", sIsPerc, sIsTypeEdit, SValue, sIsValueEdit) = False Then
          GoTo ExitHere
        End If
      End If
    Next rCnt  
    For cCnt = 1 To sprMultiParamPrompt.MaxCols - 1 Step ParamColCnt
      sParamID = Val(Spread_Get_CellTag(sprMultiParamPrompt, cCnt + eMultiParamSpr.mIsPerc, SpreadHeader))
      For rCnt = 1 To sprMultiParamPrompt.DataRowCnt
        sprMultiParamPrompt.GetText cCnt + eMultiParamSpr.mIsPerc, rCnt, retVal
        sIsPerc = IIf(Trim$(retVal) = "1", True, False) 'Is Type Percentage
        
        If isInvoiceScreen Then
          sIsTypeEdit = False
          sIsValueEdit = False
        Else
          sprMultiParamPrompt.GetText cCnt + eMultiParamSpr.mIsTypeEditable, rCnt, retVal
          sIsTypeEdit = IIf(Trim$(retVal) = "1", True, False) 'Is Type Editable
          sprMultiParamPrompt.GetText cCnt + eMultiParamSpr.mIsValueEditable, rCnt, retVal
          sIsValueEdit = IIf(Trim$(retVal) = "1", True, False) 'Is Value Editable
        End If
        
        sprMultiParamPrompt.GetText cCnt + eMultiParamSpr.mValue, rCnt, retVal
        SValue = Val(retVal)       'Value Used
        
        sprMultiParamPrompt.GetText 1, rCnt, retVal
        If Trim$(retVal) = "" Then retVal = rCnt
        If InsertPromptValues(Trim$(txtScreen.Tag), sFrmID, "Param", CStr(sParamID), Trim$(retVal), sIsPerc, sIsTypeEdit, SValue, sIsValueEdit) = False Then
          GoTo ExitHere
        End If
      
      Next rCnt
    Next cCnt
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  GoTo UnloadHere
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
  Exit Sub
UnloadHere:
  Unload Me
End Try
End Sub

Private Sub sprParamPrompt_KeyUp(sender as object, e as KeyUpEventArgs)
  Dim tmpVar As Variant
  Dim retVal As Variant
  Dim fRow As Long

' On Error Resume Next
    fRow = sprParamPrompt.ActiveRow
    If sprParamPrompt.ActiveCol = eParamSpr.mValue Then                'In Value Cell
      If sprParamPrompt.EditMode = False Then                         'When not in Edit Mode
        If KeyCode = vbKeyDelete Then                   'Delete Key Pressed
          sprParamPrompt.SetText eParamSpr.mValue, fRow, ""      'Set Empty String
        End If
      Else                                                'For Calling Defined Values Help
        sprParamPrompt.GetText eParamSpr.mValDefinedType, fRow, tmpVar
        sprParamPrompt.GetText eParamSpr.mIsPerc, fRow, retVal
        If Val(retVal) = 0 Then     'Value
          If Trim$(tmpVar) = "Both" Or Trim$(tmpVar) = "Value" Then
            Call ShowDefinedValueHelp(sprParamPrompt, sprParamPrompt.ActiveCol, fRow, True)
            KeyCode = 0
          End If
        Else                        'Percentage
          If Trim$(tmpVar) = "Both" Or Trim$(tmpVar) = "Perc" Then
            Call ShowDefinedValueHelp(sprParamPrompt, sprParamPrompt.ActiveCol, fRow, True)
            KeyCode = 0
          End If
        End If
      End If
    End If
End Sub

Private Sub sprMultiParamPrompt_KeyUp(sender as object, e as KeyUpEventArgs)
  Dim retVal As Variant
  Dim fRow As Long
  Dim fCol As Long

' On Error Resume Next
    fRow = sprMultiParamPrompt.ActiveRow
    fCol = sprMultiParamPrompt.ActiveCol
    If (fCol - 1) Mod ParamColCnt = eMultiParamSpr.mValue Then              'In Value Cell
      If sprMultiParamPrompt.EditMode = False Then                         'When not in Edit Mode
        If KeyCode = vbKeyDelete Then                   'Delete Key Pressed
          sprMultiParamPrompt.SetText fCol, fRow, ""      'Set Empty String
        End If
      Else                                                'For Calling Defined Values Help
        fCol = fCol - eMultiParamSpr.mValue + 1           'Col Set as isTypePercentage
        retVal = Trim$(Spread_Get_CellTag(sprMultiParamPrompt, fCol, SpreadHeader + 1))
        If Trim$(retVal) <> "None" Then
          Call ShowDefinedValueHelp(sprMultiParamPrompt, fCol, fRow, True)
          KeyCode = 0
        End If
      End If
    End If
End Sub

Private Sub txtScreen_Change()

  Try

  If Trim$(txtScreen.Text) = "Formula Tester" Then
    isFrmTest = True
    ParamColCnt = 4
    
      sprMultiParamPrompt.Col = 1
sprMultiParamPrompt.ColHidden = True
  Else
    isFrmTest = False
    ParamColCnt = 5
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)


Finally

End Try
End Sub

Private Sub Design_sprMultiParamPrompt(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprMultiParamPrompt


sprMultiParamPrompt.ButtonClicked +=  new EventHandler(sprMultiParamPrompt_ButtonClicked)
sprMultiParamPrompt.KeyUp +=  new EventHandler(sprMultiParamPrompt_KeyUp)

End Sub

Private Sub Design_sprParamPrompt(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprParamPrompt


sprParamPrompt.ButtonClicked +=  new EventHandler(sprParamPrompt_ButtonClicked)
sprParamPrompt.KeyUp +=  new EventHandler(sprParamPrompt_KeyUp)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprMultiParamPrompt("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmPromptValueEntry.frm", sprMultiParamPrompt, "")

	Call Design_sprParamPrompt("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Sales\frmPromptValueEntry.frm", sprParamPrompt, "")

End Sub
