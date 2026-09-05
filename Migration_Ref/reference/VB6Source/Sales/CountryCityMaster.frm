VERSION 5.00
Begin VB.Form frmCountryCityMaster 
   Caption         =   "Country / City Master"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CountryCityMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   8595
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCat 
      ClipControls    =   0   'False
      Height          =   1545
      Left            =   2633
      TabIndex        =   6
      Top             =   1815
      Width           =   5895
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         MaxLength       =   50
         TabIndex        =   2
         ToolTipText     =   "Enter the City. Press F2 or Double Click for Help"
         Top             =   1080
         Width           =   4950
      End
      Begin VB.ComboBox cmbType 
         Appearance      =   0  'Flat
         Height          =   330
         ItemData        =   "CountryCityMaster.frx":000C
         Left            =   840
         List            =   "CountryCityMaster.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Select the Country or City"
         Top             =   210
         Width           =   2100
      End
      Begin VB.TextBox txtCountry 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   840
         MaxLength       =   50
         TabIndex        =   1
         ToolTipText     =   "Enter the Country. Press F2 or Double Click for Help"
         Top             =   615
         Width           =   4950
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4500
         MaxLength       =   10
         TabIndex        =   8
         Top             =   210
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         Height          =   210
         Left            =   4035
         TabIndex        =   11
         Top             =   255
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblCity 
         AutoSize        =   -1  'True
         Caption         =   "* City"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   435
         TabIndex        =   10
         Top             =   1110
         Width           =   375
      End
      Begin VB.Label lblCountry 
         AutoSize        =   -1  'True
         Caption         =   "* Country"
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   135
         TabIndex        =   9
         Top             =   675
         Width           =   675
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "&Type"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   450
         TabIndex        =   7
         Top             =   270
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6225
      TabIndex        =   5
      ToolTipText     =   "Click to Exit from the Current Screen"
      Top             =   3450
      Width           =   1300
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   4930
      TabIndex        =   4
      ToolTipText     =   "Click to Clear the Data"
      Top             =   3450
      Width           =   1300
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   400
      Left            =   3645
      TabIndex        =   3
      ToolTipText     =   "Click to Save the Data"
      Top             =   3450
      Width           =   1300
   End
End
Attribute VB_Name = "frmCountryCityMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Sub CmdClear_Click()

  Try
  Me.MousePointer = vbHourglass

  Form_Clear Me
  CmdSave.Enabled = False
  cmbType.ListIndex = 0
  cmbType_Click
  cmbType.Enabled = True
  txtCountry.Tag = ""
txtCity.Tag = ""
  cmbType.SetFocus
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmdExit_Click()
  Try
  
  Unload Me
  
  Exit Sub
Catch ex As Exception
  Call ErrMsgShow(ex)
  
End Sub

Private Function ValidateData() As Boolean
    
  Try
  ValidateData = False
  
  If Trim$(txtCountry.Text) = "" Then
    Select Case Trim$(cmbType)
      Case "Country"
        sShowMessage "Enter the Country Name"
      Case "City"
        sShowMessage "Enter the City Name"
    End Select
    txtCountry.SetFocus
    GoTo ExitHere
  End If
  If StrComp(Trim$(cmbType.Text), "City", vbTextCompare) = 0 Then
    If Val(txtCountry.Tag) <= 0 Then
      sShowMessage "You can create City only for already created Country"
      txtCountry.SetFocus
      GoTo ExitHere
    End If
  End If
  
  ValidateData = True
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Function

Private Sub cmdSave_Click()
  Dim cmd As AceADODB.Command
  Dim snap As AceADODB.Recordset
Dim    sql As String
Dim  isInsert As Boolean
Dim  isCategory As Boolean

  Dim errMsg As String
  
  Try
  Me.MousePointer = vbHourglass

  If ValidateData = False Then GoTo ExitHere
  
  If DBStartTransaction = False Then
 GoTo ExitHere
End If
  
  Set snap = New AceADODB.Recordset
  Set cmd = New AceADODB.Command
  cmd.ActiveConnection = db
  cmd.CommandType = adCmdStoredProc
  commClear cmd
  If StrComp(Trim$(cmbType.Text), "Country", vbTextCompare) = 0 Then
    isCategory = True
    If Val(txtCountry.Tag) > 0 Then
      cmd.CommandText = "COM_Up_CountryCity_Master"
      isInsert = False
    Else
      cmd.CommandText = "COM_Ins_CountryCity_Master"
      isInsert = True
      sql = " select CountryName from COM_CountryCity_Master where ParentCode=0"
      sql = sql & " and CountryName = '" + Trim$(txtCountry) + "'"
      If RSOpen(snap, sql) = False Then GoTo ExitHere
      If Not snap.EOF Then
        errMsg = "Duplicate Description Not allowed"
        Call DBRollBackWithMessage(, errMsg, mCritical)
        txtCountry.SetFocus
        GoTo ExitHere
      End If
      sql = "Select isNull(Max(CountryCode), 0)+1 as CountryCode From COM_CountryCity_Master"
      sql = sql & " Where  ParentCode = 0 "
    End If
  ElseIf StrComp(Trim$(cmbType.Text), "City", vbTextCompare) = 0 Then
    isCategory = False
    If Val(txtCountry.Tag) > 0 And Val(txtCity.Tag) > 0 Then
      cmd.CommandText = "COM_Up_CountryCity_Master"
      isInsert = False
    Else
      cmd.CommandText = "COM_Ins_CountryCity_Master"
      isInsert = True
      sql = " select CountryName from COM_CountryCity_Master where ParentCode=" & Val(txtCountry.Tag) & " "
      sql = sql + " and CountryCode=" & Val(txtCity.Tag) & " and CountryName = '" + Trim$(txtCountry) + "' "
      If RSOpen(snap, sql) = False Then GoTo ExitHere
      If Not snap.EOF Then
        errMsg = "Duplicate Description Not allowed"
        Call DBRollBackWithMessage(, errMsg, mCritical)
        txtCity.SetFocus
        GoTo ExitHere
      End If
      sql = "Select isNull(Max(CountryCode), 0)+1 as CountryCode From COM_CountryCity_Master"
      sql = sql & " Where  ParentCode=" & Val(txtCountry.Tag) & ""
    End If
  End If
  If isInsert = True Then
    If RSOpen(snap, sql) = False Then GoTo ExitHere
    cmd.Parameters.Append cmd.CreateParameter("CountryCode", adInteger, adParamInput, , snap("CountryCode"))
  Else
    cmd.Parameters.Append cmd.CreateParameter("CountryCode", adInteger, adParamInput, , IIf(isCategory, Val(txtCountry.Tag), Val(txtCity.Tag)))
  End If
  cmd.Parameters.Append cmd.CreateParameter("CountryName", adVarChar, adParamInput, 50, IIf(isCategory, Left$(Trim$(txtCountry.Text), 50), Left$(Trim$(txtCity.Text), 50)))
  cmd.Parameters.Append cmd.CreateParameter("ParentCode", adInteger, adParamInput, , IIf(isCategory, 0, Val(txtCountry.Tag)))
  cmd.Parameters.Append cmd.CreateParameter("UserID", adVarChar, adParamInput, 10, gUserID)
  cmd.Execute
  
  If DBCommitTransaction = False Then
 GoTo ExitHere
End If
  
  Set snap = Nothing
Set cmd = Nothing
  If isInsert = True Then
    sShowMessage "Record Saved"
  Else
    sShowMessage "Record Updated"
  End If
  
  Call CmdClear_Click
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Set snap = Nothing
  Set cmd = Nothing
  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub cmbType_Click()
  Try

  If StrComp(Trim$(cmbType.Text), "Country", vbTextCompare) = 0 Then
    lblCity.Enabled = False
txtCity.Enabled = False

    txtCity.Text = ""
txtCity.Tag = ""
txtCountry.Locked = False
    txtCountry = ""
  Else
    txtCountry = ""
    lblCity.Enabled = True
    txtCity.Enabled = True
    txtCountry.Locked = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub Form_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
  Try
    
  If KeyCode = 13 Then
 SendKeys "{tab}"
End If
Catch ex As Exception
  Call ErrMsgShow(ex)
  
Finally

End Try
End Sub

Private Sub Form_Load()
  Try
  
  cmbType.ListIndex = 0
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

Private Sub txtCode_KeyUp(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  Call CheckSave
End Sub

Private Sub txtCountry_DblClick()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  sql = "Select CountryName,CountryCode From COM_CountryCity_Master where ParentCode=0 "
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of Country", "Country Name") = True Then
    If Trim$(getvalue(1)) = "" Then GoTo ExitHere
    txtCountry.Tag = getvalue(2)
    txtCountry.Text = getvalue(1)
    If StrComp(Trim$(cmbType.Text), "City", vbTextCompare) = 0 Then
      lblCity.Enabled = True
txtCity.Enabled = True
txtCode.Enabled = False
    Else
      CmdSave.Enabled = True
    End If
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtCountry_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then Call txtCountry_DblClick
End Sub

Private Sub txtCountry_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  KeyAscii = ToUpper(KeyAscii)
End Sub

Private Sub txtCountry_KeyUp(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  Call CheckSave
End Sub

Private Sub CheckSave()

  Try
  
  If StrComp(Trim$(cmbType.Text), "Country", vbTextCompare) = 0 Then
    If Trim$(txtCountry) <> "" Then CmdSave.Enabled = True Else CmdSave.Enabled = False
  Else
    If Trim$(txtCountry) <> "" And Trim$(txtCity) <> "" Then CmdSave.Enabled = True Else CmdSave.Enabled = False
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

End Try
End Sub

Private Sub txtCity_DblClick()
  Dim sql As String
  
  Try
  Me.MousePointer = vbHourglass

  If Val(txtCountry.Tag) <= 0 Then
    MsgBox "Please Select the Country."
    txtCountry.SetFocus
GoTo ExitHere
  End If
  
  sql = "Select CountryName,CountryCode From COM_CountryCity_Master where ParentCode=" & Val(txtCountry.Tag) & ""
  Me.MousePointer = vbHourglass
  If FetchValue_New(False, db.ConnectionString, sql, "10", "List of City", "City Name") = True Then
    If Trim$(getvalue(1)) = "" Then GoTo ExitHere
    txtCity.Tag = getvalue(2)
    txtCity.Text = getvalue(1)
    lblCity.Enabled = True
txtCity.Enabled = True
    CmdSave.Enabled = True
  End If
Catch ex As Exception
  Call ErrMsgShow(ex)

Finally

  Me.MousePointer = vbDefault
End Try
End Sub

Private Sub txtCity_KeyDown(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  If KeyCode = vbKeyF2 Then Call txtCity_DblClick
End Sub

Private Sub txtCity_KeyPress(ByVal KeyAscii As Integer)
' On Error Resume Next
  KeyAscii = ToUpper(KeyAscii)
End Sub

Private Sub txtCity_KeyUp(ByVal KeyCode As Integer,ByVal  Shift As Integer)
' On Error Resume Next
  Call CheckSave
End Sub

