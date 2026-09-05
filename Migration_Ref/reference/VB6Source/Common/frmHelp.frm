VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMulti 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   5685
      TabIndex        =   10
      Top             =   15
      Width           =   4395
      Begin VB.CheckBox chkSelected 
         Caption         =   "Selected Items Only"
         Height          =   240
         Left            =   2640
         TabIndex        =   14
         Top             =   690
         Width           =   1725
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear All"
         Height          =   330
         Left            =   1350
         TabIndex        =   13
         Top             =   660
         Width           =   1180
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select All"
         Height          =   330
         Left            =   75
         TabIndex        =   12
         Top             =   660
         Width           =   1180
      End
      Begin VB.Label lblMulti 
         AutoSize        =   -1  'True
         Caption         =   "Multi Selection Enabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   210
         Left            =   1395
         TabIndex        =   11
         Top             =   225
         Width           =   1905
      End
   End
   Begin VB.Timer RunTime1 
      Interval        =   65535
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fraFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   15
      Width           =   5505
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find &Next - F3"
         Height          =   315
         Left            =   4095
         TabIndex        =   4
         Top             =   195
         Width           =   1275
      End
      Begin VB.TextBox txtFind 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   795
         MaxLength       =   50
         TabIndex        =   2
         Top             =   195
         Width           =   3225
      End
      Begin MSComCtl2.DTPicker dtpFind 
         Height          =   315
         Left            =   1020
         TabIndex        =   3
         Top             =   195
         Visible         =   0   'False
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   141361155
         CurrentDate     =   37630
         MinDate         =   32874
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   255
         TabIndex        =   1
         Top             =   210
         Width           =   465
      End
   End
   Begin VB.Frame fraSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      TabIndex        =   5
      Top             =   615
      Width           =   5505
      Begin VB.ComboBox cmbFind 
         Height          =   330
         ItemData        =   "frmHelp.frx":0000
         Left            =   855
         List            =   "frmHelp.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   195
         Width           =   1755
      End
      Begin VB.ComboBox cmbSearchMode 
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "frmHelp.frx":0004
         Left            =   3720
         List            =   "frmHelp.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   195
         Width           =   1725
      End
      Begin VB.Label lblFindIn 
         AutoSize        =   -1  'True
         Caption         =   "&Search in"
         Height          =   210
         Left            =   105
         TabIndex        =   6
         Top             =   255
         Width           =   690
      End
      Begin VB.Label lblSearchMode 
         AutoSize        =   -1  'True
         Caption         =   "Search &Mode"
         Height          =   210
         Left            =   2685
         TabIndex        =   8
         Top             =   255
         Width           =   960
      End
   End
   Begin VB.Frame fraSpread 
      Height          =   4440
      Left            =   120
      TabIndex        =   16
      Top             =   1530
      Width           =   9960
Begin AceSpread sprHelp
         Height          =   4155
         Left            =   75
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   210
         Width           =   9810
         _Version        =   458752
         _ExtentX        =   17304
         _ExtentY        =   7329
         _StockProps     =   64
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
         MaxRows         =   1
         SelectBlockOptions=   2
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   510
      Left            =   4890
      TabIndex        =   20
      Top             =   6345
      Width           =   3420
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   510
      Left            =   1485
      TabIndex        =   19
      Top             =   6345
      Width           =   3420
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Close this Form by Pressing Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   1275
      Width           =   3135
   End
   Begin VB.Label lblNoRec 
      AutoSize        =   -1  'True
      Caption         =   "Number of Records :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      TabIndex        =   18
      Top             =   6060
      Width           =   1785
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private adoHelp As AceADODB.Recordset

Dim  SortOrder() As Integer
Dim  FindIndex() As Integer
Dim  ColIndex() As Integer

Dim  isFormLoaded As Boolean
Dim  ColCap() As String
Dim  FindContStatus As Boolean

Dim cmpOperator(1 To 2) As String
Dim isDontSort As Boolean
Dim localOrderbyCol As String
Dim LastCol As Long
Dim lastColOrderAsc As Boolean
Dim sRow As Long
Dim sMaxCols As Long

Private Sub chkSelected_Click()
  Dim rCnt As Long
    If chkSelected Then
      For rCnt = 1 To sprHelp.DataRowCnt
        sprHelp.Row = rCnt
        If sprHelp.SelModeSelected = False Then
          sprHelp.RowHidden = True
        End If
      Next rCnt
      sprHelp.Enabled = False
fraFind.Enabled = False
fraSearch.Enabled = False

      cmdSelect.Enabled = False
CmdClear.Enabled = False
    Else
      For rCnt = 1 To sprHelp.DataRowCnt
        sprHelp.Row = rCnt
        sprHelp.RowHidden = False
      Next rCnt
      sprHelp.Enabled = True
fraFind.Enabled = True
fraSearch.Enabled = True

      cmdSelect.Enabled = True
CmdClear.Enabled = True
    End If
End Sub

Private Sub cmbFind_Click()
  Dim DtType As Integer
    
  If isFormLoaded Then
    If LastCol = Val(FindIndex(cmbFind.ListIndex) + 1) Then
      Call SortSpread(sprHelp, FindIndex(cmbFind.ListIndex) + 1, IIf(lastColOrderAsc, False, True))
      lastColOrderAsc = IIf(lastColOrderAsc, False, True)
    Else
      Call SortSpread(sprHelp, FindIndex(cmbFind.ListIndex) + 1, True)
      lastColOrderAsc = True
    End If
    LastCol = Val(FindIndex(cmbFind.ListIndex) + 1)
     'sprHelp.UserColAction = UserColActionSort
  End If
  
  DtType = adoHelp.Fields(FindIndex(cmbFind.ListIndex)).Type
  Select Case DtType
    Case 2, 3, 4, 5, 14, 20, 131
      txtFind.Visible = True
dtpFind.Visible = False

      cmbSearchMode.Enabled = False
cmbSearchMode.ListIndex = 0
    Case 129, 200
      txtFind.Visible = True
dtpFind.Visible = False

      cmbSearchMode.Enabled = True
cmbSearchMode.ListIndex = 0
    Case 133, 134, 135
      txtFind.Visible = False
dtpFind.Visible = True

      cmbSearchMode.Enabled = False
cmbSearchMode.ListIndex = 0
    Case 11
      txtFind.Visible = True
dtpFind.Visible = False

      cmbSearchMode.Enabled = False
cmbSearchMode.ListIndex = 0
  End Select
  
  If txtFind.Visible = True And dtpFind.Visible = False Then
    Call txtFind_Change ': txtFind.SetFocus
  ElseIf txtFind.Visible = False And dtpFind.Visible = True Then
    Call dtpFind_Change ': dtpFind.SetFocus
  End If
End Sub

Private Sub cmdCancel_Click()
  Dim hCol As Integer
  Erase getvalue
  ReDim getvalue(0) As Variant
  getvalue(hCol) = ""
  sMaxCols = 0
  Unload Me
End Sub

Private Sub cmdClear_Click()
  Dim rCnt As Long
    For rCnt = 1 To sprHelp.DataRowCnt
      sprHelp.Row = rCnt
      sprHelp.SelModeSelected = False
      sprHelp.SetText sprHelp.MaxCols, rCnt, ""
    Next rCnt
End Sub

Private Sub cmdFindNext_Click()
    FindContStatus = True
  If txtFind.Visible = True Then
    Call txtFind_Change
  ElseIf dtpFind.Visible = True Then
    Call dtpFind_Change
  End If
End Sub

Private Sub cmdOK_Click()
Dim    rSelCnt As Long
Dim  DataCnt As Long

  Dim SelOrder As Long
  Dim retVal As Variant
Dim    i As Long
Dim  sCol As Long

  
    If sprHelp.OperationMode = OperationModeMulti Then
      If sprHelp.SelModeSelCount > 0 Then
        If localOrderbyCol <> "" Then
          sCol = 1
          For i = 1 To Len(localOrderbyCol) Step 2
            sprHelp.SortKey(sCol) = Val(Mid(localOrderbyCol, i, 2))
            sprHelp.SortKeyOrder(sCol) = SortKeyOrderAscending
            sCol = sCol + 1
          Next i
          sprHelp.Sort 1, 1, -1, -1, SortByRow
          localOrderbyCol = ""
        End If
        ReDim getvalue(1 To sprHelp.SelModeSelCount, 1 To sMaxCols) As Variant
        DataCnt = 0
rSelCnt = 0
        Do While DataCnt <= sprHelp.SelModeSelCount
          rSelCnt = sprHelp.GetMultiSelItem(rSelCnt)
          If rSelCnt = -1 Then Exit Do
          DataCnt = DataCnt + 1
          sprHelp.GetText sprHelp.MaxCols, rSelCnt, retVal
          SelOrder = Val(retVal)
          Call ReturnMultiValues(rSelCnt, DataCnt, SelOrder)
        Loop
      Else
        sMaxCols = 0
        Call cmdCancel_Click
Exit Sub
      End If
    Else
      ReDim getvalue(1 To sMaxCols) As Variant
      Call ReturnValues(sprHelp.SelModeIndex)
    End If
    sMaxCols = 0
  Unload Me
End Sub

Private Sub cmdSelect_Click()
  Dim rCnt As Long
  If isFormLoaded Then
      For rCnt = 1 To sprHelp.DataRowCnt
        sprHelp.Row = rCnt
        sprHelp.SelModeSelected = True
        If Trim$(SpVal(sprHelp, rCnt, sprHelp.MaxCols)) = "" Then
          sprHelp.SetText sprHelp.MaxCols, rCnt, sprHelp.SelectionCount
        End If
      Next rCnt
  End If
End Sub

Private Sub dtpFind_Change()
  Dim sRow As Long
    
' On Error Resume Next
 
  sRow = 0

  If isFormLoaded Then
    If cmbFind.ListIndex >= 0 Then
      If cmbSearchMode.Enabled = True And cmbSearchMode.ListIndex = 1 Then
        sRow = sprHelp.SearchCol(FindIndex(cmbFind.ListIndex) + 1, sRow, sprHelp.MaxRows, Trim$(dtpFind), SearchFlagsGreaterOrEqual + SearchFlagsPartialMatch)
      Else
        sRow = sprHelp.SearchCol(FindIndex(cmbFind.ListIndex) + 1, sRow, sprHelp.MaxRows, Trim$(dtpFind), SearchFlagsPartialMatch)
      End If
      
      If sRow < 1 Then
        sprHelp.SetActiveCell 1, 1
      Else
        sprHelp.SetActiveCell 1, sRow
      End If
    End If
    FindContStatus = False
    DoEvents
    sprHelp.Refresh
  End If
End Sub

Private Sub Form_Activate()
  Call FindTextFocusSet
  sRow = 0
  'Added By Saravanan For reseting the Counter Timer
  'Add This In Form Activate Event (Help.frm)
  '*************************************************
'  MDIQC.RunTime.Interval = 0
'  TimeCnt = 0
'  MDIQC.RunTime.Interval = 65535
'  Unload frmPopUp
  ''************************************************
End Sub

'Add this In HElp Form
'*************************
'Private Sub RunTime1_Timer()
''  TCnt = TCnt + 1
''  If TCnt = 3 Then Unload Me
'End Sub
Private Sub Form_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
    Select Case KeyCode
      Case vbKeyF3
        Call cmdFindNext_Click
      Case vbKeyDown
        If sprHelp.ActiveRow < sprHelp.MaxRows Then
          sprHelp.SetActiveCell sprHelp.ActiveCol, sprHelp.ActiveRow + 1
        End If
      Case vbKeyUp
        If sprHelp.ActiveRow > 1 Then
          sprHelp.SetActiveCell sprHelp.ActiveCol, sprHelp.ActiveRow - 1
        End If
    End Select
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
  End If
  sMaxCols = 0
End Sub

Private Sub sprHelp_AfterUserSort(sender as object, e as AfterUserSortEventArgs)
' On Error Resume Next

  If Col > 0 Then
    cmbFind.ListIndex = ColIndex(Col)
    Call FindTextFocusSet
  End If
  DoEvents
End Sub

Private Sub sprHelp_Click(sender as object, e as ClickEventArgs)
  Dim s As Long
  Dim getcnt As Long
  Dim sRow As Long
  Dim retVal As Variant

  If Row = 0 And Col > 0 Then
    cmbFind.ListIndex = ColIndex(Col)
    
    'sprHelp.UserColAction = UserColActionSort
    
    If LastCol = Val(FindIndex(cmbFind.ListIndex) + 1) Then
      Call SortSpread(sprHelp, FindIndex(cmbFind.ListIndex) + 1, IIf(lastColOrderAsc, False, True))
      lastColOrderAsc = IIf(lastColOrderAsc, False, True)
    Else
      Call SortSpread(sprHelp, FindIndex(cmbFind.ListIndex) + 1, True)
      lastColOrderAsc = True
    End If
    LastCol = Val(FindIndex(cmbFind.ListIndex) + 1)
    Call FindTextFocusSet

      sprHelp.ClearSelection
      For sRow = 1 To sprHelp.DataRowCnt
        sprHelp.GetText sprHelp.MaxCols, sRow, retVal
        If Trim$(retVal) <> "" Then
          sprHelp.Row = sRow
          sprHelp.SelModeSelected = True
        End If
      Next sRow
  End If
  
  If Trim$(SpVal(sprHelp, Row, sprHelp.MaxCols)) = "" Then
    If Row > 0 Then sprHelp.SetText sprHelp.MaxCols, Row, sprHelp.SelectionCount
  Else
    getcnt = Val(SpVal(sprHelp, Row, sprHelp.MaxCols))
    sprHelp.SetText sprHelp.MaxCols, Row, ""
    sRow = 0
    For s = 1 To sprHelp.SelectionCount
      sRow = sprHelp.GetMultiSelItem(sRow)
      If Val(SpVal(sprHelp, sRow, sprHelp.MaxCols)) > getcnt Then
        sprHelp.SetText sprHelp.MaxCols, sRow, (Val(SpVal(sprHelp, sRow, sprHelp.MaxCols)) - 1)
        If Val(SpVal(sprHelp, sRow, sprHelp.MaxCols)) = 0 Then
          sprHelp.SetText sprHelp.MaxCols, sRow, ""
        End If
      End If
      
    Next s
  End If
  DoEvents
  
End Sub

Private Sub sprHelp_DblClick(sender as object, e as DblClickEventArgs)
  If Row <> 0 Then
    If sprHelp.OperationMode = OperationModeSingle Then
      ReDim getvalue(1 To sMaxCols) As Variant
      Call ReturnValues(Row)
      Unload Me
    End If
  End If
End Sub

Private Sub txtFind_Change()
  Dim DType As CellTypeConstants
    
' On Error Resume Next
  
  If Len(Trim$(txtFind)) = 0 Then
    sprHelp.SetActiveCell 1, 1
    Exit Sub
  End If
  
  If FindContStatus = False Then
    sRow = -1
  End If
  
  sprHelp.Row = 1
sprHelp.Col = FindIndex(cmbFind.ListIndex) + 1
  DType = sprHelp.CellType
  
  If cmbFind.ListIndex >= 0 Then
    If DType = CellTypeNumber Then
      If cmbSearchMode.Enabled = True And cmbSearchMode.ListIndex = 1 Then
        sRow = sprHelp.SearchCol(FindIndex(cmbFind.ListIndex) + 1, sRow, sprHelp.MaxRows, Val(txtFind), SearchFlagsValue + SearchFlagsPartialMatch)
      Else
        sRow = sprHelp.SearchCol(FindIndex(cmbFind.ListIndex) + 1, sRow, sprHelp.MaxRows, Val(txtFind), SearchFlagsValue + SearchFlagsPartialMatch)
      End If
    Else
      If cmbSearchMode.Enabled = True And cmbSearchMode.ListIndex = 1 Then
        sRow = sprHelp.SearchCol(FindIndex(cmbFind.ListIndex) + 1, sRow, sprHelp.MaxRows, Trim$(txtFind), SearchFlagsGreaterOrEqual + SearchFlagsPartialMatch)
      Else
        sRow = sprHelp.SearchCol(FindIndex(cmbFind.ListIndex) + 1, sRow, sprHelp.MaxRows, Trim$(txtFind), SearchFlagsGreaterOrEqual)
      End If
    End If
    
    If sRow < 1 Then
      sprHelp.SetActiveCell 1, 1
    Else
      sprHelp.SetActiveCell 1, sRow
    End If
  End If
  
  FindContStatus = False
  DoEvents
  sprHelp.Refresh
End Sub

Private Sub txtFind_KeyPress(ByVal KeyAscii As Integer)
  Select Case Chr(KeyAscii)
    Case "%", "_", "[", "]", "^", "*", "?"
      MsgBox "The following characters " & "% _ [ ] ^ * ?" & " are not allowed in find string"
      KeyAscii = 0
  End Select
End Sub

Public Function GetHelp(ByVal isMultiSelect As Boolean,ByVal  mydbConStr As String,ByVal  SqlStr As String,ByVal  ColVisStatus As String,ByVal  HelpCaption As String,ByVal  ColCaption() As Variant,Optional ByVal OrderbyCol As String = "",Optional ByVal isAlreadySorted As Boolean = False) As Boolean
Dim    i As Integer
Dim  visColCnt As Integer

Dim    ColVisCntStr As String
Dim  tempStr As String

  Dim FindColIndex As Integer
  Dim FirstVisibleCol As Long
  
  Try
  Me.MousePointer = vbHourglass

  GetHelp = False
  
  isFormLoaded = False
  
  Erase getvalue
  ReDim getvalue(0) As Variant
  getvalue(0) = ""
  
  Set adoHelp = New AceADODB.Recordset
  If RSOpen(adoHelp, SqlStr) = False Then
    GoTo ExitHere
  End If
  
  If adoHelp.RecordCount = 0 Then
    MsgBox "Record(S) Not found", vbOKOnly, App.Title
    GoTo ExitHere
  End If
  
  Call ADORS_Spread_DataFill(adoHelp, sprHelp, False)
  
'  sprHelp.DAutoSizeCols = DAutoSizeColsMax
  lblNoRec = "Number of Records: " & adoHelp.RecordCount
  lblTitle = HelpCaption
  
  If Len(ColVisStatus) = 0 Then
    Call cmdCancel_Click
  ElseIf Len(ColVisStatus) > adoHelp.Fields.Count Then
    ColVisStatus = Mid$(ColVisStatus, 1, adoHelp.Fields.Count)
  End If
  localOrderbyCol = OrderbyCol
  
  ColVisCntStr = ""
FirstVisibleCol = 0
  For i = 1 To Len(ColVisStatus)
    tempStr = Mid$(ColVisStatus, i, 1)
    If tempStr = "0" Or tempStr = "1" Then
      ColVisCntStr = ColVisCntStr & tempStr
    ElseIf tempStr = "?" Or tempStr = "_" Then
      ColVisCntStr = ColVisCntStr & "1"
    ElseIf tempStr = "*" Or tempStr = "%" Then
      ColVisCntStr = ColVisCntStr & String((adoHelp.Fields.Count) - Len(ColVisCntStr), "1")
      Exit For
    Else
      ColVisCntStr = ColVisCntStr & "1"
    End If
  Next i
  
  ColVisCntStr = ColVisCntStr & String((adoHelp.Fields.Count) - Len(ColVisCntStr), "0")
  
  For i = 1 To Len(ColVisCntStr)
    tempStr = Mid$(ColVisCntStr, i, 1)
    If Val(tempStr) = 1 Then
      FirstVisibleCol = i
      Exit For
    End If
  Next i
  
    ReDim SortOrder(1 To sprHelp.MaxCols) As Integer 'SortOrder(ColNo)  is 0-nothing, 1-ascending & is 2-descending
    ReDim ColCap(1 To sprHelp.MaxCols) As String
    ReDim FindIndex(0 To sprHelp.MaxCols - 1) As Integer
    ReDim ColIndex(1 To sprHelp.MaxCols) As Integer
    
    Erase SortOrder
    cmbFind.Clear
    
    If InStr(1, ColVisCntStr, "1", vbTextCompare) Then
      visColCnt = 1
      For i = 1 To sprHelp.MaxCols
        If Mid$(ColVisCntStr, i, 1) = 1 Then
          If UBound(ColCaption) + 1 >= visColCnt Then
            If Trim$(ColCaption(visColCnt - 1)) <> "" Then
              ColCap(visColCnt) = Trim$(ColCaption(visColCnt - 1))
            Else
              ColCap(visColCnt) = adoHelp.Fields(i - 1).Name
            End If
            sprHelp.SetText i, 0, ColCap(visColCnt)
            cmbFind.AddItem ColCap(visColCnt)
            FindIndex(visColCnt - 1) = i - 1
          Else
            ColCap(visColCnt) = adoHelp.Fields(i - 1).Name
            sprHelp.SetText i, 0, ColCap(visColCnt)
            cmbFind.AddItem ColCap(visColCnt)
            FindIndex(visColCnt - 1) = i - 1
          End If
          ColIndex(i) = visColCnt - 1
          visColCnt = visColCnt + 1
        Else
          ColIndex(i) = -1
          sprHelp.Col = i
sprHelp.ColHidden = True
        End If
      Next
      
      ReDim Preserve FindIndex(0 To visColCnt - 2) As Integer
      ReDim SortOrder(1 To visColCnt - 1) As Integer
      ReDim Preserve ColCap(1 To visColCnt - 1) As String
      
      cmbSearchMode.ListIndex = 0
      If isMultiSelect Then
        sprHelp.OperationMode = OperationModeMulti
        lblMulti.Caption = "Multi Selection Mode"
        fraMulti.Enabled = True
      Else
        sprHelp.OperationMode = OperationModeSingle
        lblMulti.Caption = "Single Selection Mode"
        fraMulti.Enabled = False
      End If
      
      sprHelp.Row = 1
sprHelp.RowHeight(0) = 15
      
      dtpFind.Value = CDate(Format(SetServerDate, "dd/mm/yyyy"))
      If dtpFind < FinStart Then GoTo ExitHere
      
      FindContStatus = False
      
      Call ColWidthFix
      
      ' Don't change the following if else block. it is not just same coding but line changed. it operates differently.
      'If InStr(1, SqlStr, "Order By", vbTextCompare) < 1 Or isAlreadySorted = False Then
      If isAlreadySorted = False Then
        isFormLoaded = True
        cmbFind.ListIndex = 0
      Else
        cmbFind.ListIndex = 0
        isFormLoaded = True
      End If
            
      'add by ramana for new col for sort on Selected Row on 2013-10-15 11:00am
      sMaxCols = sprHelp.MaxCols
      sprHelp.MaxCols = sprHelp.MaxCols + 1
      Spread_HideCol sprHelp, True, sprHelp.MaxCols, sprHelp.MaxCols
      
      If adoHelp.RecordCount = 1 Then
        sprHelp.Row = 1
        sprHelp.SelModeSelected = True
        sprHelp.SetText sprHelp.MaxCols, 1, 1
      End If
      
      Me.MousePointer = vbDefault
' On Error Resume Next

      Me.Show vbModal
      Try:
    Else
      Call cmdCancel_Click
      GoTo ExitHere
    End If
    
  
  GetHelp = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub ColHeadFill()
  Dim cCnt As Integer
  For cCnt = 1 To UBound(ColCap)
    sprHelp.SetText FindIndex(cCnt - 1) + 1, 0, ColCap(cCnt)
  Next
End Sub

Private Sub ColWidthFix()
  Dim i As Integer
    For i = 0 To UBound(FindIndex)
      sprHelp.ColWidth(FindIndex(i) + 1) = sprHelp.MaxTextColWidth(FindIndex(i) + 1) + 1
    Next i
End Sub

Private Sub FindTextFocusSet()
  If txtFind.Visible = True And dtpFind.Visible = False Then
    txtFind.SetFocus
  ElseIf txtFind.Visible = False And dtpFind.Visible = True Then
    dtpFind.SetFocus
  End If
End Sub

Private Sub ReturnMultiValues(ByVal sRow As Long,ByVal  dCnt As Long,ByVal  SelOrder As Long)
Dim    hCol As Integer
Dim  sData As Variant

  For hCol = 1 To sMaxCols
    sprHelp.GetText hCol, sRow, sData
    getvalue(SelOrder, hCol) = sData
  Next hCol
End Sub

Private Sub ReturnValues(ByVal sRow As Long)
'  Dim sam(1, 12) As Variant
'  sprHelp.GetArray 1, 1, sam
Dim    hCol As Integer
Dim  sData As Variant

  For hCol = 1 To sMaxCols
    sprHelp.GetText hCol, sRow, sData
    getvalue(hCol) = sData
  Next hCol
End Sub

Private Sub Design_sprHelp(ByVal FormPath As String,ByRef  SourceSpread As AceSpread,ByVal  SpreadIndex As String)
		Call Set_SpreadAllCoding_InTextFile(FormPath, Me.Name, SourceSpread, SpreadIndex) 'sprHelp


sprHelp.AfterUserSort +=  new EventHandler(sprHelp_AfterUserSort)
sprHelp.Click +=  new EventHandler(sprHelp_Click)
sprHelp.DblClick +=  new EventHandler(sprHelp_DblClick)

End Sub

Private Sub InitializeSpreads()

	Call Design_sprHelp("D:\OutOfSync\_2_UnicornERP_Convertion_VB6ToNextJS_Antigravity\Export_VB6_Sales\PrePost_Sales\Sales\Common\frmHelp.frm", sprHelp, "")

End Sub
