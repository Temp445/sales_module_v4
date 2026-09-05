VERSION 5.00
Begin VB.Form frmTermsMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Terms Master"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
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
   ScaleHeight     =   4125
   ScaleWidth      =   6810
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cbexit 
      Caption         =   "E&xit"
      Height          =   400
      Left            =   6180
      TabIndex        =   5
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   2985
      Width           =   1300
   End
   Begin VB.CommandButton cbclear 
      Caption         =   "&Clear"
      Height          =   400
      Left            =   4890
      TabIndex        =   4
      ToolTipText     =   "Click to Clear the Data"
      Top             =   2985
      Width           =   1300
   End
   Begin VB.CommandButton cbsave 
      Caption         =   "&Save"
      Height          =   400
      Left            =   3600
      TabIndex        =   3
      ToolTipText     =   "Click to Save the Data"
      Top             =   2985
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   2775
      TabIndex        =   6
      Top             =   1515
      Width           =   5610
      Begin VB.ComboBox cmbStatus 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         ItemData        =   "TermsMaster.frx":0000
         Left            =   945
         List            =   "TermsMaster.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         ToolTipText     =   "Select the Status"
         Top             =   960
         Width           =   1395
      End
      Begin VB.ComboBox cmbTerms 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "TermsMaster.frx":0020
         Left            =   945
         List            =   "TermsMaster.frx":003C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select the Term"
         Top             =   210
         Width           =   2100
      End
      Begin VB.TextBox tbDesc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   945
         MaxLength       =   100
         TabIndex        =   1
         ToolTipText     =   "Enter the Description. Press F2 or Double Click for Help."
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Status"
         Height          =   210
         Left            =   450
         TabIndex        =   9
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   105
         TabIndex        =   8
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Terms"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   465
         TabIndex        =   7
         Top             =   270
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmTermsMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub GenerateCode()
  Try
  
  If CmbTerms <> "" Then
    tbDesc.Tag = Val(Find_Value("TERMSMASTER", "isnull(max(convert(int,Code)),0)", "TypeFlg", Trim$(CmbTerms))) + 1
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Function ValidateData() As Boolean
  Dim sql As String
  Dim snap As AceADODB.Recordset
  
  Try
  ValidateData = False
  
  If contrl_valid(tbDesc, "Enter the Description") = False Then GoTo ExitHere
  
  sql = " select Code from TERMSMASTER where typeflg = '" + Trim$(CmbTerms.Text) + "' and Description = '" + Trim$(tbDesc.Text) + "'"
  sql = sql + " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
  Set snap = New AceADODB.Recordset
  If snap.State = adStateOpen Then snap.Close
  snap.Open sql, db, adOpenStatic, adLockReadOnly
  If Not snap.EOF Then
    sShowMessage "Duplicate Entry not allowed"
    tbDesc.SetFocus
    GoTo ExitHere
  End If

 ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
End Try
End Function

Private Sub cbclear_Click()
  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  cmbStatus.Enabled = False
  CmbTerms.ListIndex = 0
  cmbStatus.ListIndex = 0
  Call cmbTerms_Click
  CmbTerms.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub cbExit_Click()
' On Error Resume Next
  Unload Me
End Sub

Private Sub cbsave_Click()
  Dim sql As String
  Dim snap As AceADODB.Recordset
  Dim cmd As AceADODB.Command
  
  Try
  Me.MousePointer = vbHourglass

  Set snap = New AceADODB.Recordset
  Set cmd = New AceADODB.Command
  
  If ValidateData = False Then
    GoTo ExitHere
  End If
  
  sql = " select Code from TERMSMASTER where typeflg = '" + Trim$(CmbTerms.Text) + "' and Code = '" + Trim$(tbDesc.Tag) + "'"
  sql = sql + " and Branchid = '" & gBranchID & "' and Compid = '" & gCompID & "'"
  If snap.State = adStateOpen Then snap.Close
  snap.Open sql, db, adOpenStatic, adLockReadOnly
  snap.ActiveConnection = Nothing

  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  If snap.EOF Then
    cmd.CommandText = "insTERMSMASTER"
  Else
    cmd.CommandText = "upTERMSMASTER"
  End If
  commClear cmd
  cmd.Parameters.Append cmd.CreateParameter("TypeFlg", adChar, adParamInput, 50, Trim$(CmbTerms.Text))
  cmd.Parameters.Append cmd.CreateParameter("Code", adVarChar, adParamInput, 6, Trim$(tbDesc.Tag))
  cmd.Parameters.Append cmd.CreateParameter("Description", adVarChar, adParamInput, 100, Trim$(tbDesc.Text))
  cmd.Parameters.Append cmd.CreateParameter("Status", adBoolean, adParamInput, 1, IIf(cmbStatus = "Active", True, False))
  Call AddCommonCmdParameters(cmd, , False)
  cmd.Execute
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  sShowMessage "Record Saved"
  
  Call cbclear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Set cmd = Nothing
  Me.MousePointer = Default
End Try
End Sub

Private Sub cmbTerms_Click()
  
  Try
  Me.MousePointer = vbHourglass

  GenerateCode
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub Form_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  If KeyAscii = vbKeyReturn Then
 SendKeys "{tab}"
End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub Form_Load()
  Try

  CmbTerms.ListIndex = 0
  cmbStatus.ListIndex = 0
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_QueryUnload(ByVal Cancel As Integer,ByVal  UnloadMode As Integer)
  If FormExit(Me, UnloadMode) = False Then
    Cancel = True
  End If
End Sub

Private Sub tbDesc_DblClick()
  Dim sql As String
  Try
  Me.MousePointer = vbHourglass

  sql = "Select Code,description, Status from TERMSMASTER where typeflg = '" + Trim$(CmbTerms.Text) + "' "
  sql = sql + " and Branchid='" & gBranchID & "' and Compid ='" & gCompID & "'"
  If FetchValue_New(False, db.ConnectionString, sql, "*", "List of Codes", "Code", "Description") = True Then
    tbDesc.Tag = getvalue(1)
    tbDesc.Text = getvalue(2)
    If getvalue(3) = 0 Then
      cmbStatus.Text = "InActive"
    Else
      cmbStatus.Text = "Active"
    End If
    cmbStatus.Enabled = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

Private Sub tbDesc_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then
 tbDesc_DblClick
End If
End Sub

Private Sub tbDesc_KeyPress(ByVal KeyAscii As Integer)
  Try
  Me.MousePointer = vbHourglass

  KeyAscii = Asc(UCase(Chr(KeyAscii)))
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = 0
End Try
End Sub

